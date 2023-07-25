/*
 * IdMUnit - Automated Testing Framework for Identity Management Solutions
 * Copyright (c) 2005-2023 TriVir, LLC
 *
 * This program is licensed under the terms of the GNU General Public License
 * Version 2 (the "License") as published by the Free Software Foundation, and
 * the TriVir Licensing Policies (the "License Policies").  A copy of the License
 * and the Policies were distributed with this program.
 *
 * The License is available at:
 * http://www.gnu.org/copyleft/gpl.html
 *
 * The Policies are available at:
 * http://www.idmunit.org/licensing/index.html
 *
 * Unless required by applicable law or agreed to in writing, this program is
 * distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS
 * OF ANY KIND, either express or implied.  See the License and the Policies
 * for specific language governing the use of this program.
 *
 * www.TriVir.com
 * TriVir LLC
 * 13890 Braddock Road
 * Suite 310
 * Centreville, Virginia 20121
 *
 */

package com.trivir.idmunit.cli.converter;

import com.trivir.idmunit.cli.converter.model.*;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelWriter {

    private static final int EXCEL_WIDTH_CONSTANT = 256;

    private int numOperationConfigHeaders;
    private final Workbook workbook;
    private Sheet sheet = null;
    private IdmUnitTest idmUnitTest = null;
    private int nextRow = 0;
    private int maxAttrSize = 0;
    private final Map<String, Map<String, Integer>> connectorAttrIndices = new LinkedHashMap<>();
    private CellStyle titleCellStyle = null;
    private CellStyle boldCellStyle = null;
    private CellStyle delimiterCellStyle = null;
    private CellStyle opConfigHeaderStyle = null;
    private CellStyle connectorAttrHeaderStyle = null;
    private CellStyle commentCellStyle = null;
    private CellStyle borderedCellStyle = null;

    public ExcelWriter(Workbook workbook) {
        this.workbook = workbook;
        initStyles();
    }

    public void writeTest(IdmUnitTest idmTest) {
        idmUnitTest = idmTest;
        numOperationConfigHeaders = 7;
        if (idmUnitTest.getHasIsCriticalConfigHeader() != null) {
            numOperationConfigHeaders += 1;
        }
        if (idmUnitTest.getHasRepeatOpRangeConfigHeader() != null) {
            numOperationConfigHeaders += 1;
        }
        sheet = workbook.createSheet(idmUnitTest.getName());
        sheet.setDefaultRowHeightInPoints(15);

        maxAttrSize = idmUnitTest.getConnectors().stream()
            .map(x -> x.getAttributes().stream().map(ConnectorAttribute::getGroupNum).max(Integer::compare).orElse(0))
            .max(Integer::compare)
            .orElse(0);
        populateConnectorIndicesMap();

        writeTitle(idmUnitTest.getTitle());
        writeDescription(idmUnitTest.getDesc());
        writeDelimiterRow();
        writeFirstHeaderRow();
        idmUnitTest.getConnectors().forEach(this::writeTarget);
        writeDelimiterRow();
        idmUnitTest.getOperations().forEach(operation -> {
            if (operation.getOperation().trim().equals("comment")) {
                writeCommentRow(operation);
            } else {
                writeOperation(operation);
            }
        });
        writeDelimiterRow();

        sheet.setColumnWidth(0, 35 * EXCEL_WIDTH_CONSTANT);
        for (int i = 1; i < numOperationConfigHeaders; i++) {
            sheet.setColumnWidth(i, 15 * EXCEL_WIDTH_CONSTANT);
        }
        for (int i = numOperationConfigHeaders; i < getNumColumns(); i++) {
            sheet.setColumnWidth(i, 35 * EXCEL_WIDTH_CONSTANT);
        }
        sheet = null;
        nextRow = 0;
        maxAttrSize = 0;
    }

    private void initStyles() {
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 14);

        titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setFont(titleFont);

        Font boldFont = workbook.createFont();
        boldFont.setBold(true);

        boldCellStyle = workbook.createCellStyle();
        boldCellStyle.setFont(boldFont);

        delimiterCellStyle = workbook.createCellStyle();
        delimiterCellStyle.setFont(boldFont);
        delimiterCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
        delimiterCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        opConfigHeaderStyle = workbook.createCellStyle();
        opConfigHeaderStyle.setFont(boldFont);
        opConfigHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        opConfigHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        opConfigHeaderStyle.setBorderBottom(BorderStyle.THIN);
        opConfigHeaderStyle.setBorderTop(BorderStyle.THIN);
        opConfigHeaderStyle.setBorderLeft(BorderStyle.THIN);
        opConfigHeaderStyle.setBorderRight(BorderStyle.THIN);

        connectorAttrHeaderStyle = workbook.createCellStyle();
        connectorAttrHeaderStyle.setFont(boldFont);
        connectorAttrHeaderStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
        connectorAttrHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        connectorAttrHeaderStyle.setBorderBottom(BorderStyle.THIN);
        connectorAttrHeaderStyle.setBorderTop(BorderStyle.THIN);
        connectorAttrHeaderStyle.setBorderLeft(BorderStyle.THIN);
        connectorAttrHeaderStyle.setBorderRight(BorderStyle.THIN);

        commentCellStyle = workbook.createCellStyle();
        commentCellStyle.setFont(boldFont);
        commentCellStyle.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.index);
        commentCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        commentCellStyle.setBorderBottom(BorderStyle.THIN);
        commentCellStyle.setBorderTop(BorderStyle.THIN);
        commentCellStyle.setBorderLeft(BorderStyle.THIN);
        commentCellStyle.setBorderRight(BorderStyle.THIN);

        borderedCellStyle = workbook.createCellStyle();
        borderedCellStyle.setBorderBottom(BorderStyle.THIN);
        borderedCellStyle.setBorderTop(BorderStyle.THIN);
        borderedCellStyle.setBorderLeft(BorderStyle.THIN);
        borderedCellStyle.setBorderRight(BorderStyle.THIN);
    }

    private int getNumColumns() {
        return numOperationConfigHeaders + maxAttrSize + 1;
    }

    private void populateConnectorIndicesMap() {
        for (Connector connector : idmUnitTest.getConnectors()) {
            Map<String, Integer> indicesMap = new HashMap<>();
            for (int i = 0; i < connector.getAttributes().size(); i++) {
                ConnectorAttribute attr = connector.getAttributes().get(i);
                indicesMap.put(attr.getName(), numOperationConfigHeaders + attr.getGroupNum());
            }
            connectorAttrIndices.put(connector.getName(), indicesMap);
        }
    }

    private void writePaddingCellsFrom(int startingIndex, Row row, CellStyle cellStyle) {
        for (int i = startingIndex; i < getNumColumns(); i++) {
            Cell paddingCell = row.createCell(i, CellType.BLANK);
            paddingCell.setCellStyle(cellStyle);
        }
    }

    private void writeTitle(String title) {
        Row titleRow = sheet.createRow(nextRow++);
        titleRow.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());

        Cell cell = titleRow.createCell(0, CellType.STRING);
        cell.setCellValue(title);
        cell.setCellStyle(titleCellStyle);
    }

    private void writeDescription(String desc) {
        Row descRow = sheet.createRow(nextRow++);
        descRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        Cell cell = descRow.createCell(0, CellType.STRING);
        cell.setCellValue(desc);
        cell.setCellStyle(boldCellStyle);
    }

    private void writeDelimiterRow() {
        Row delimiterRow = sheet.createRow(nextRow++);
        delimiterRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        writePaddingCellsFrom(0, delimiterRow, delimiterCellStyle);

        Cell cell = delimiterRow.getCell(0);
        cell.setCellValue("---");
    }

    private void writeFirstHeaderRow() {
        Row firstHeaderRow = sheet.createRow(nextRow++);
        firstHeaderRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        writePaddingCellsFrom(0, firstHeaderRow, connectorAttrHeaderStyle);

        Cell commentCell = firstHeaderRow.getCell(0);
        commentCell.setCellValue(OperationConfigHeader.COMMENT.getExcelHeader());
        commentCell.setCellStyle(opConfigHeaderStyle);

        Cell operationCell = firstHeaderRow.getCell(1);
        operationCell.setCellValue(OperationConfigHeader.OPERATION.getExcelHeader());
        operationCell.setCellStyle(opConfigHeaderStyle);

        Cell targetCell = firstHeaderRow.getCell(2);
        targetCell.setCellValue(OperationConfigHeader.TARGET.getExcelHeader());
        targetCell.setCellStyle(opConfigHeaderStyle);

        Cell waitIntervalCell = firstHeaderRow.getCell(3);
        waitIntervalCell.setCellValue(OperationConfigHeader.WAIT_INTERVAL.getExcelHeader());
        waitIntervalCell.setCellStyle(opConfigHeaderStyle);

        Cell retryCountCell = firstHeaderRow.getCell(4);
        retryCountCell.setCellValue(OperationConfigHeader.RETRY_COUNT.getExcelHeader());
        retryCountCell.setCellStyle(opConfigHeaderStyle);

        Cell disableStepCell = firstHeaderRow.getCell(5);
        disableStepCell.setCellValue(OperationConfigHeader.DISABLE_STEP.getExcelHeader());
        disableStepCell.setCellStyle(opConfigHeaderStyle);

        Cell expectFailureCell = firstHeaderRow.getCell(6);
        expectFailureCell.setCellValue(OperationConfigHeader.EXPECT_FAILURE.getExcelHeader());
        expectFailureCell.setCellStyle(opConfigHeaderStyle);

        if (idmUnitTest.getHasIsCriticalConfigHeader() != null) {
            Cell isCriticalCell = firstHeaderRow.getCell(7);
            isCriticalCell.setCellValue(OperationConfigHeader.IS_CRITICAL.getExcelHeader());
            isCriticalCell.setCellStyle(opConfigHeaderStyle);
        }

        if (idmUnitTest.getHasRepeatOpRangeConfigHeader() != null) {
            Cell repeatOpRangeCell;
            if (idmUnitTest.getHasIsCriticalConfigHeader() == null) {
                repeatOpRangeCell = firstHeaderRow.getCell(7);
            } else {
                repeatOpRangeCell = firstHeaderRow.getCell(8);
            }
            repeatOpRangeCell.setCellValue(OperationConfigHeader.REPEAT_OP_RANGE.getExcelHeader());
            repeatOpRangeCell.setCellStyle(opConfigHeaderStyle);
        }
    }

    private void writeTarget(Connector connector) {
        Row targetRow = sheet.createRow(nextRow++);
        targetRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        writePaddingCellsFrom(0, targetRow, connectorAttrHeaderStyle);

        Cell commentCell = targetRow.getCell(0);
        commentCell.setCellStyle(opConfigHeaderStyle);

        Cell operationCell = targetRow.getCell(1);
        operationCell.setCellStyle(opConfigHeaderStyle);

        Cell targetCell = targetRow.getCell(2);
        targetCell.setCellValue(connector.getName());
        targetCell.setCellStyle(opConfigHeaderStyle);

        Cell waitIntervalCell = targetRow.getCell(3);
        waitIntervalCell.setCellStyle(opConfigHeaderStyle);

        Cell retryCountCell = targetRow.getCell(4);
        retryCountCell.setCellStyle(opConfigHeaderStyle);

        Cell disableStepCell = targetRow.getCell(5);
        disableStepCell.setCellStyle(opConfigHeaderStyle);

        Cell expectFailureCell = targetRow.getCell(6);
        expectFailureCell.setCellStyle(opConfigHeaderStyle);

        if (idmUnitTest.getHasIsCriticalConfigHeader() != null) {
            Cell isCriticalCell = targetRow.getCell(7);
            isCriticalCell.setCellStyle(opConfigHeaderStyle);
        }

        if (idmUnitTest.getHasRepeatOpRangeConfigHeader() != null) {
            Cell repeatOpRangeCell;
            if (idmUnitTest.getHasIsCriticalConfigHeader() == null) {
                repeatOpRangeCell = targetRow.getCell(7);
            } else {
                repeatOpRangeCell = targetRow.getCell(8);
            }
            repeatOpRangeCell.setCellStyle(opConfigHeaderStyle);
        }

        for (int i = 0; i < connector.getAttributes().size(); i++) {
            ConnectorAttribute attr = connector.getAttributes().get(i);
            Cell attrCell = targetRow.getCell(connectorAttrIndices.get(connector.getName()).get(attr.getName()));
            attrCell.setCellStyle(connectorAttrHeaderStyle);
            attrCell.setCellValue(attr.getName());
        }
    }

    private void writeCommentRow(Operation operation) {
        Row commentRow = sheet.createRow(nextRow++);
        commentRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        Cell commentCell = commentRow.createCell(0, CellType.STRING);
        commentCell.setCellStyle(commentCellStyle);
        commentCell.setCellValue(operation.getComment());

        Cell operationCell = commentRow.createCell(1, CellType.STRING);
        operationCell.setCellStyle(commentCellStyle);
        operationCell.setCellValue(operation.getOperation());

        writePaddingCellsFrom(2, commentRow, commentCellStyle);
    }

    private void writeOperation(Operation operation) {
        Row operationRow = sheet.createRow(nextRow++);
        operationRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints());

        writePaddingCellsFrom(0, operationRow, borderedCellStyle);

        Cell commentCell = operationRow.getCell(0);
        commentCell.setCellValue(operation.getComment());
        commentCell.setCellStyle(borderedCellStyle);

        Cell operationCell = operationRow.getCell(1);
        operationCell.setCellValue(operation.getOperation());
        operationCell.setCellStyle(borderedCellStyle);

        Cell targetCell = operationRow.getCell(2);
        targetCell.setCellValue(operation.getTarget());
        targetCell.setCellStyle(borderedCellStyle);

        Cell waitIntervalCell = operationRow.getCell(3);
        waitIntervalCell.setCellValue(operation.getWaitInterval());
        waitIntervalCell.setCellStyle(borderedCellStyle);

        Cell retryCountCell = operationRow.getCell(4);
        retryCountCell.setCellValue(operation.getRetryCount());
        retryCountCell.setCellStyle(borderedCellStyle);

        Cell disableStepCell = operationRow.getCell(5);
        disableStepCell.setCellValue(operation.getDisabled());
        disableStepCell.setCellStyle(borderedCellStyle);

        Cell expectFailureCell = operationRow.getCell(6);
        expectFailureCell.setCellValue(operation.getFailureExpected());
        expectFailureCell.setCellStyle(borderedCellStyle);

        if (idmUnitTest.getHasIsCriticalConfigHeader() != null) {
            Cell isCriticalCell = operationRow.getCell(7);
            isCriticalCell.setCellStyle(borderedCellStyle);
        }

        if (idmUnitTest.getHasRepeatOpRangeConfigHeader() != null) {
            Cell repeatOpRangeCell;
            if (idmUnitTest.getHasIsCriticalConfigHeader() == null) {
                repeatOpRangeCell = operationRow.getCell(7);
            } else {
                repeatOpRangeCell = operationRow.getCell(8);
            }
            repeatOpRangeCell.setCellStyle(borderedCellStyle);
        }

        if (operation.getData() == null) {
            return;
        }

        for (int i = 0; i < operation.getData().size(); i++) {
            if (!connectorAttrIndices.containsKey(operation.getTarget())) {
                throw new RuntimeException("Operation with undefined target specified");
            }
            OperationData data = operation.getData().get(i);
            Integer index = connectorAttrIndices.get(operation.getTarget()).get(data.getAttribute());
            if (index == null) {
                throw new RuntimeException("Operation data for undefined target attribute specified");
            }
            Cell dataCell = operationRow.getCell(index);
            dataCell.setCellStyle(borderedCellStyle);
            if (operation.getData().get(i).getMeta() != null && operation.getData().get(i).getMeta().contains("excel:isFormula")) {
                dataCell.setCellFormula(operation.getData().get(i).getValue().get(0));
            } else {
                dataCell.setCellValue(String.join("|", operation.getData().get(i).getValue()));
            }
        }
    }

}
