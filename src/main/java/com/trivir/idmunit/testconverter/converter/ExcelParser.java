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

package com.trivir.idmunit.testconverter.converter;

import com.trivir.idmunit.testconverter.converter.model.*;
import lombok.Getter;
import lombok.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ExcelParser {

    private static final String SECTION_DELIMITER_ROW_VALUE = "---";
    private static final String COMMENT_OPERATION_VALUE = "comment";

    @Getter
    private String sheetName;

    @Getter
    private final LintMessages lintMessages;

    private boolean sheetHasIsCriticalOpConfigHeader = false;
    private boolean sheetHasRepeatOpRangeOpConfigHeader = false;

    public ExcelParser(boolean verbose) {
        lintMessages = new LintMessages(verbose);
    }

    public IdmUnitTest parseSheet(Sheet sheet) throws IdmUnitTestConverterException {
        lintMessages.clear();
        sheetHasIsCriticalOpConfigHeader = false;
        sheetHasRepeatOpRangeOpConfigHeader = false;
        this.sheetName = sheet.getSheetName();
        // Get row sections
        List<Row> sectionDelimiterRows = StreamSupport.stream(sheet.spliterator(), false)
            .filter(x -> ExcelUtils.parseCellAsString(ExcelUtils.getCellOrBlank(x, 0)).equals(SECTION_DELIMITER_ROW_VALUE))
            .collect(Collectors.toList());
        if (sectionDelimiterRows.size() < 3) {
            throw lintMessages.errorTooFewSectionDelimiterRows(sectionDelimiterRows.size());
        } else if (sectionDelimiterRows.size() > 3) {
            lintMessages.warnTooManySectionDelimiterRows(sectionDelimiterRows.size());
        }
        sectionDelimiterRows.forEach(this::checkSectionDelimiterRow);
        RowGroups rowGroups = parseRowGroups(sheet, sectionDelimiterRows);
        if (rowGroups.getConnectorRows().isEmpty()) {
            throw lintMessages.errorNoRowsInConnectorsSection();
        }
        // Test Details Section
        String testName = sheet.getSheetName();
        TestDetails testDetails = parseTestDetails(rowGroups.getTestDetailsRows());
        // Connectors Section
        FirstRowHeaders firstRowHeaders = parseFirstConnectorRow(rowGroups.getConnectorRows().get(0));
        List<Row> connectorRows = rowGroups.getConnectorRows().subList(1, rowGroups.getConnectorRows().size());
        Map<String, List<CellWrapper>> connectorAttributesMap = parseConnectorAttributes(connectorRows, firstRowHeaders.getOperationConfigHeaders());
        // Operations Section
        List<Map<String, CellWrapper>> operationDataList = parseOperations(rowGroups.getOperationRows(), firstRowHeaders, connectorAttributesMap);
        // Unknown Rows
        rowGroups.unknownRows.forEach(this::checkUnknownRow);
        // Map into Java structure for easy JSON conversion
        List<Connector> connectors = connectorAttributesMap.entrySet().stream()
            .map(entry -> mapConnector(entry.getKey(), entry.getValue()))
            .collect(Collectors.toList());
        List<Operation> operations = operationDataList.stream()
            .map(this::mapOperationData)
            .collect(Collectors.toList());
        List<Connector> inferredConnectors = operations.stream()
            .map(Operation::getTarget)
            .filter(x -> x != null && !x.trim().isEmpty())
            .filter(x -> !connectorAttributesMap.containsKey(x))
            .map(x -> mapConnector(x, firstRowHeaders.getDefaultConnectorAttributes()))
            .collect(Collectors.toList());
        connectors.addAll(inferredConnectors);
        normalizeConnectorAttrGroupNums(connectors);
        IdmUnitTest idmUnitTest = new IdmUnitTest();
        idmUnitTest.setName(testName);
        idmUnitTest.setTitle(testDetails.getTitle());
        idmUnitTest.setDesc(testDetails.getDescription());
        idmUnitTest.setConnectors(connectors);
        idmUnitTest.setOperations(operations);
        idmUnitTest.setHasIsCriticalConfigHeader(sheetHasIsCriticalOpConfigHeader ? true : null);
        idmUnitTest.setHasRepeatOpRangeConfigHeader(sheetHasRepeatOpRangeOpConfigHeader ? true : null);
        return idmUnitTest;
    }

    private void checkSectionDelimiterRow(Row row) {
        StreamSupport.stream(row.spliterator(), false)
            .skip(1)
            .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
            .forEach(lintMessages::warnCellWithValueOnSectionDelimiterRow);
    }

    private RowGroups parseRowGroups(Sheet sheet, List<Row> delimiterRows) {
        Map<Integer, List<Row>> groups = StreamSupport.stream(sheet.spliterator(), false)
            .filter(row -> delimiterRows.stream().noneMatch(x -> row.getRowNum() == x.getRowNum()))
            .collect(Collectors.groupingBy(x -> determineRowGroup(x, delimiterRows)));
        return new RowGroups(
            groups.getOrDefault(0, new ArrayList<>()),
            groups.getOrDefault(1, new ArrayList<>()),
            groups.getOrDefault(2, new ArrayList<>()),
            groups.getOrDefault(3, new ArrayList<>())
        );
    }

    private int determineRowGroup(Row row, List<Row> delimiterRows) {
        for (int i = 0; i < delimiterRows.size(); i++) {
            if (row.getRowNum() < delimiterRows.get(i).getRowNum()) {
                return i;
            }
        }
        return delimiterRows.size();
    }

    private TestDetails parseTestDetails(List<Row> titleRows) {
        if (titleRows.isEmpty()) {
            lintMessages.warnNoRowsInTestDetailsSection();
            return new TestDetails("", "");
        }
        String title = ExcelUtils.parseCellAsString(ExcelUtils.getCellOrBlank(titleRows.get(0), 0));
        if (title.trim().isEmpty()) {
            lintMessages.warnNoTitle();
        }
        String description = "";
        if (titleRows.size() > 1) {
            description = ExcelUtils.parseCellAsString(ExcelUtils.getCellOrBlank(titleRows.get(1), 0));
        }
        List<Cell> badCells = titleRows.stream()
            .map(row -> StreamSupport.stream(row.spliterator(), false)
                .filter(cell -> !ExcelUtils.parseCellAsString(cell).trim().isEmpty())
                .filter(cell -> cell.getRowIndex() > 1 || cell.getColumnIndex() > 0)
                .collect(Collectors.toList()))
            .flatMap(Collection::stream)
            .collect(Collectors.toList());
        badCells.forEach(lintMessages::warnExtraCellWithValueInTestDetailsSection);
        return new TestDetails(title, description);
    }

    private FirstRowHeaders parseFirstConnectorRow(Row row) {
        Map<Boolean, List<CellWrapper>> headers = StreamSupport.stream(row.spliterator(), false)
            .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
            .map(CellWrapper::new)
            .collect(Collectors.partitioningBy(x -> x.getValue().startsWith(OperationConfigHeader.PREFIX)));
        List<CellWrapper> operationConfigHeaders = headers.get(true).stream()
            // No connector attributes should start with the Operation Config Prefix
            .peek(x -> {
                if (!OperationConfigHeader.isKnownExcelOpConfigHeader(x.getValue())) {
                    lintMessages.warnUnknownHeaderWithOperationConfigPrefix(x.getCell(), OperationConfigHeader.PREFIX);
                }
            })
            .filter(x -> OperationConfigHeader.isKnownExcelOpConfigHeader(x.getValue()))
            .collect(Collectors.toList());
        // Check for operation config headers: Error if no Target, mark if there is IsCritical or RepeatOpRange
        boolean hasTargetHeader = false;
        for (CellWrapper configHeaderCell : operationConfigHeaders) {
            if (configHeaderCell.getValue().equals(OperationConfigHeader.TARGET.getExcelHeader())) {
                hasTargetHeader = true;
            } else if (configHeaderCell.getValue().equals(OperationConfigHeader.IS_CRITICAL.getExcelHeader())) {
                sheetHasIsCriticalOpConfigHeader = true;
            } else if (configHeaderCell.getValue().equals(OperationConfigHeader.REPEAT_OP_RANGE.getExcelHeader())) {
                sheetHasRepeatOpRangeOpConfigHeader = true;
            }
        }
        if (!hasTargetHeader) {
            throw lintMessages.errorNoTargetOperationConfigHeader();
        }
        return new FirstRowHeaders(operationConfigHeaders, headers.get(false));
    }

    private Map<String, List<CellWrapper>> parseConnectorAttributes(List<Row> connectorRows, List<CellWrapper> operationConfigHeaders) {
        Map<String, List<CellWrapper>> connectorAttrsMap = new LinkedHashMap<>();
        int targetColIndex = operationConfigHeaders.stream()
            .filter(x -> x.getValue().equals(OperationConfigHeader.TARGET.getExcelHeader()))
            .findFirst()
            // Should already have thrown error if target column was not defined
            .orElseThrow(lintMessages::errorNoTargetOperationConfigHeader)
            .getCell()
            .getColumnIndex();
        List<Integer> operationConfigHeaderColIndices = operationConfigHeaders.stream()
            .filter(x -> !x.getValue().equals(OperationConfigHeader.TARGET.getExcelHeader()))
            .map(x -> x.getCell().getColumnIndex())
            .collect(Collectors.toList());
        for (Row row : connectorRows) {
            String connectorName = ExcelUtils.parseCellAsString(ExcelUtils.getCellOrBlank(row, targetColIndex));
            if (connectorName.trim().isEmpty()) {
                lintMessages.warnConnectorRowWithNoName(row);
                continue;
            }
            List<CellWrapper> attrs = StreamSupport.stream(row.spliterator(), false)
                .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
                .peek(x -> {
                    if (operationConfigHeaderColIndices.contains(x.getColumnIndex())) {
                        lintMessages.warnConnectorAttributeUnderOperationConfigHeader(x);
                    }
                })
                .filter(x -> !operationConfigHeaderColIndices.contains(x.getColumnIndex()) && x.getColumnIndex() != targetColIndex)
                .map(CellWrapper::new)
                .collect(Collectors.toList());
            if (connectorAttrsMap.containsKey(connectorName)) {
                int originalRowNum = connectorAttrsMap.get(connectorName).get(0).getCell().getRowIndex();
                lintMessages.warnConnectorRowWithSameName(row, connectorName, originalRowNum);
            }
            connectorAttrsMap.put(connectorName, attrs);
        }
        return connectorAttrsMap;
    }

    private List<Map<String, CellWrapper>> parseOperations(List<Row> operationRows, FirstRowHeaders firstRowHeaders, Map<String, List<CellWrapper>> connectorAttrsMap) {
        List<Row> blankRows = operationRows.stream()
            .filter(this::isRowBlank)
            .collect(Collectors.toList());
        // No blank rows allowed
        Optional<Row> blankRow = blankRows.stream().findFirst();
        if (blankRow.isPresent()) {
            throw lintMessages.errorBlankOperationRow(blankRow.get());
        }
        List<Map<String, CellWrapper>> operationDataList = new ArrayList<>();
        for (Row row : operationRows) {
            Map<String, CellWrapper> operationData = new LinkedHashMap<>();
            // Collect all cells under Operation Config Headers in this row
            for (CellWrapper header : firstRowHeaders.getOperationConfigHeaders()) {
                Cell cell = ExcelUtils.getCellOrBlank(row, header.getCell().getColumnIndex());
                operationData.put(header.getValue(), new CellWrapper(cell));
            }
            // Handle comment operation
            CellWrapper operationCell = operationData.get(OperationConfigHeader.OPERATION.getExcelHeader());
            if (operationCell != null && operationCell.getValue().trim().equals(COMMENT_OPERATION_VALUE)) {
                operationDataList.add(parseCommentOperation(row, operationData));
                continue;
            }
            // Ensure target connector is defined for this operation
            CellWrapper targetConnector = operationData.get(OperationConfigHeader.TARGET.getExcelHeader());
            if (targetConnector == null || targetConnector.getValue().trim().isEmpty()) {
                throw lintMessages.errorOperationRowWithNoTargetDefined(row);
            }
            // Use default connector attributes if target connector was not defined in Connectors Section
            List<CellWrapper> connectorAttrs = connectorAttrsMap.get(targetConnector.getValue());
            if (connectorAttrs == null) {
                connectorAttrs = firstRowHeaders.getDefaultConnectorAttributes();
            }
            // Parse all operation data in cells under the target connectors' attrs
            for (CellWrapper attr : connectorAttrs) {
                Cell cell = ExcelUtils.getCellOrBlank(row, attr.getCell().getColumnIndex());
                if (ExcelUtils.parseCellAsString(cell).trim().isEmpty()) {
                    continue;
                }
                // Warn if two values defined for the same connector attribute
                if (operationData.containsKey(attr.getValue())) {
                    Cell originalCell = operationData.get(attr.getValue()).getCell();
                    lintMessages.warnOperationDataForDuplicateAttr(cell, originalCell, attr.getValue());
                }
                operationData.put(attr.getValue(), new CellWrapper(cell));
            }
            // Warn if the row contains non-blank cells in columns with no header (operation config or connector attr)
            List<Integer> knownColumns = connectorAttrs.stream()
                .map(header -> header.getCell().getColumnIndex())
                .collect(Collectors.toList());
            knownColumns.addAll(
                firstRowHeaders.getOperationConfigHeaders().stream()
                    .map(header -> header.getCell().getColumnIndex())
                    .collect(Collectors.toList())
            );
            StreamSupport.stream(row.spliterator(), false)
                .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
                .filter(x -> !knownColumns.contains(x.getColumnIndex()))
                .forEach(lintMessages::warnNonBlankCellInColumnWithNoHeader);
            operationDataList.add(operationData);
        }
        return operationDataList;
    }

    private boolean isRowBlank(Row row) {
        return StreamSupport.stream(row.spliterator(), false)
            .allMatch(cell -> ExcelUtils.parseCellAsString(cell).trim().isEmpty());
    }

    private Map<String, CellWrapper> parseCommentOperation(Row row, Map<String, CellWrapper> operationData) {
        CellWrapper commentCell = operationData.get(OperationConfigHeader.COMMENT.getExcelHeader());
        // Warn if no comment is defined
        if (commentCell == null || commentCell.getValue().trim().isEmpty()) {
            lintMessages.warnCommentOperationWithNoCommentDefined(row);
        }
        // All cells should be blank on a comment operation row except under Operation and Comment config headers
        List<Integer> colIndicesToIgnore = new ArrayList<>();
        colIndicesToIgnore.add(operationData.get(OperationConfigHeader.OPERATION.getExcelHeader()).getCell().getColumnIndex());
        if (commentCell != null) {
            colIndicesToIgnore.add(operationData.get(OperationConfigHeader.COMMENT.getExcelHeader()).getCell().getColumnIndex());
        }
        StreamSupport.stream(row.spliterator(), false)
            .filter(x -> !colIndicesToIgnore.contains(x.getColumnIndex()))
            .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
            .forEach(lintMessages::warnNonBlankCellOnCommentOperationRow);
        Map<String, CellWrapper> returnMap = new HashMap<>();
        returnMap.put(OperationConfigHeader.COMMENT.getExcelHeader(), operationData.get(OperationConfigHeader.COMMENT.getExcelHeader()));
        returnMap.put(OperationConfigHeader.OPERATION.getExcelHeader(), operationData.get(OperationConfigHeader.OPERATION.getExcelHeader()));
        return returnMap;
    }

    private void checkUnknownRow(Row row) {
        StreamSupport.stream(row.spliterator(), false)
            .filter(x -> !ExcelUtils.parseCellAsString(x).trim().isEmpty())
            .forEach(lintMessages::warnNonBlankCellInRowAfterOperationSection);
    }

    private Connector mapConnector(String connectorName, List<CellWrapper> attrs) {
        final Connector connector = new Connector();
        connector.setName(connectorName);

        final List<ConnectorAttribute> attributes = attrs.stream()
            .map(x -> {
                ConnectorAttribute attribute = new ConnectorAttribute();
                attribute.setName(x.getValue());
                attribute.setGroupNum(x.getCell().getColumnIndex());
                return attribute;
            })
            .collect(Collectors.toList());
        connector.setAttributes(attributes);

        return connector;
    }

    private Operation mapOperationData(Map<String, CellWrapper> operationData) {
        Operation operation = new Operation();
        if (operationData.get(OperationConfigHeader.COMMENT.getExcelHeader()) != null) {
            operation.setComment(operationData.get(OperationConfigHeader.COMMENT.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.OPERATION.getExcelHeader()) != null) {
            operation.setOperation(operationData.get(OperationConfigHeader.OPERATION.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.TARGET.getExcelHeader()) != null) {
            operation.setTarget(operationData.get(OperationConfigHeader.TARGET.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.WAIT_INTERVAL.getExcelHeader()) != null) {
            operation.setWaitInterval(operationData.get(OperationConfigHeader.WAIT_INTERVAL.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.RETRY_COUNT.getExcelHeader()) != null) {
            operation.setRetryCount(operationData.get(OperationConfigHeader.RETRY_COUNT.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.DISABLE_STEP.getExcelHeader()) != null) {
            operation.setDisabled(operationData.get(OperationConfigHeader.DISABLE_STEP.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.EXPECT_FAILURE.getExcelHeader()) != null) {
            operation.setFailureExpected(operationData.get(OperationConfigHeader.EXPECT_FAILURE.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.IS_CRITICAL.getExcelHeader()) != null) {
            operation.setCritical(operationData.get(OperationConfigHeader.IS_CRITICAL.getExcelHeader()).getValue());
        }
        if (operationData.get(OperationConfigHeader.REPEAT_OP_RANGE.getExcelHeader()) != null) {
            operation.setRepeatRange(operationData.get(OperationConfigHeader.REPEAT_OP_RANGE.getExcelHeader()).getValue());
        }
        final List<OperationData> data = new ArrayList<>();
        for (final Map.Entry<String, CellWrapper> entry : operationData.entrySet()) {
            if (entry.getKey().startsWith(OperationConfigHeader.PREFIX)) {
                continue;
            }
            final OperationData opData = new OperationData();
            opData.setAttribute(entry.getKey());
            opData.setValue(Collections.singletonList(entry.getValue().getValue()));
            if (entry.getValue().cell.getCellType() == CellType.FORMULA) {
                if (opData.getMeta() == null) {
                    opData.setMeta(new ArrayList<>());
                }
                opData.getMeta().add("excel:isFormula");
            }
            data.add(opData);
        }
        operation.setData(data.isEmpty() ? null : data);
        return operation;
    }

    void normalizeConnectorAttrGroupNums(List<Connector> connectors) {
        List<Integer> groupNums = connectors.stream()
            .filter(x -> x.getAttributes() != null)
            .flatMap(x -> x.getAttributes().stream().map(ConnectorAttribute::getGroupNum))
            .distinct()
            .sorted()
            .collect(Collectors.toList());
        Map<Integer, Integer> normalizerMap = new HashMap<>();
        for (int i = 0; i < groupNums.size(); i++) {
            normalizerMap.put(groupNums.get(i), i);
        }
        connectors.stream()
            .filter(x -> x.getAttributes() != null)
            .forEach(x -> x.getAttributes()
                .forEach(y -> y.setGroupNum(normalizerMap.get(y.getGroupNum()))));
    }

    @Value
    private static class RowGroups {
        List<Row> testDetailsRows;
        List<Row> connectorRows;
        List<Row> operationRows;
        List<Row> unknownRows;
    }

    @Value
    private static class TestDetails {
        String title;
        String description;
    }

    @Value
    private static class FirstRowHeaders {
        List<CellWrapper> operationConfigHeaders;
        List<CellWrapper> defaultConnectorAttributes;
    }

    @Value
    private static class CellWrapper {
        Cell cell;

        public String getValue() {
            return ExcelUtils.parseCellAsString(cell);
        }
    }
}
