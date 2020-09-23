/*
 * Copyright (c) 2015-2020 TriVir LLC - All Rights Reserved
 *
 *  This software is proprietary and confidential.
 *  Unauthorized copying of this file, via any medium, is strictly prohibited.
 */

package com.trivir.idmunit.converter;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class IdMUnitTestConverter {

    static final int COMMENT_COL_INDEX = 0;
    static final int OPERATION_COL_INDEX = 1;
    static final int TARGET_COL_INDEX = 2;
    static final int WAIT_INTERVAL_COL_INDEX = 3;
    static final int RETRY_COUNT_COL_INDEX = 4;
    static final int DISABLE_STEP_COL_INDEX = 5;
    static final int EXPECT_FAILURE_COL_INDEX = 6;

    private static final String SECTION_DIVIDER_INDICATOR = "---";

    private static final String ID_KEY = "id";
    private static final String TITLE_KEY = "title";
    private static final String DESCRIPTION_KEY = "description";
    private static final String OPERATIONS_KEY = "operations";

    private static final String COMMENT_KEY = "comment";
    private static final String OPERATION_KEY = "operation";
    private static final String TARGET_KEY = "target";
    private static final String WAIT_INTERVAL_KEY = "waitInterval";
    private static final String RETRY_COUNT_KEY = "retryCount";
    private static final String DISABLED_KEY = "disabled";
    private static final String EXPECT_FAILURE_KEY = "expectFailure";
    private static final String DATA_KEY = "data";

    private final ObjectMapper objectMapper;

    public IdMUnitTestConverter() {
        objectMapper = new ObjectMapper();
    }

    public static void main(String[] args) {
    }

    ObjectNode convertSheet(Sheet sheetToConvert) {
        ObjectNode ret = objectMapper.createObjectNode();
        ret.put(ID_KEY, sheetToConvert.getSheetName());
        ret.put(TITLE_KEY, sheetToConvert.getRow(0).getCell(0).getStringCellValue());
        if (!sheetToConvert.getRow(1).getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
            ret.put(DESCRIPTION_KEY, sheetToConvert.getRow(1).getCell(0).getStringCellValue());
        }
        Map<String, Map<Integer, String>> headerInformationMap = new HashMap<>();
        ArrayNode operationsArray = objectMapper.createArrayNode();
        boolean headersParsed = false;
        boolean isHeaderRow = false;
        for (Iterator<Row> rowIterator = sheetToConvert.rowIterator(); rowIterator.hasNext(); ) {
            Row r = rowIterator.next();
            if (!headersParsed) {
                if (!isHeaderRow) {
                    if (r.getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
                        isHeaderRow = true;
                        rowIterator.next();
                        continue;
                    } else {
                        continue;
                    }
                }
                if (r.getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
                    isHeaderRow = false;
                    headersParsed = true;
                } else {
                    Map<Integer, String> headerMap = new HashMap<>();
                    for (Iterator<Cell> cellIterator = r.cellIterator(); cellIterator.hasNext(); ) {
                        Cell c = cellIterator.next();
                        if (c.getColumnIndex() > EXPECT_FAILURE_COL_INDEX) {
                            headerMap.put(c.getColumnIndex(), c.getStringCellValue());
                        }
                    }
                    headerInformationMap.put(r.getCell(TARGET_COL_INDEX).getStringCellValue(), headerMap);
                }
            } else {
                if (r.getCell(0) != null && !r.getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
                    operationsArray.add(convertRow(r, headerInformationMap.get(r.getCell(TARGET_COL_INDEX).getStringCellValue())));
                }
            }
        }
        ret.set(OPERATIONS_KEY, operationsArray);
        // Iterate over the header rows and marshall the data so that we can use it when iterating over the operation rows.

        return ret;
    }

    ObjectNode convertRow(Row rowToConvert, Map<Integer, String> headerInformation) {
        ObjectNode ret = objectMapper.createObjectNode();
        ret.put(COMMENT_KEY, rowToConvert.getCell(COMMENT_COL_INDEX).getStringCellValue());
        ret.put(OPERATION_KEY, rowToConvert.getCell(OPERATION_COL_INDEX).getStringCellValue());
        ObjectNode data = objectMapper.createObjectNode();

        if (!ret.get(OPERATION_KEY).asText().equals("comment")) {
            ret.put(TARGET_KEY, rowToConvert.getCell(TARGET_COL_INDEX).getStringCellValue());
            ret.put(WAIT_INTERVAL_KEY, getIntFromCell(rowToConvert.getCell(WAIT_INTERVAL_COL_INDEX)));
            ret.put(RETRY_COUNT_KEY, getIntFromCell(rowToConvert.getCell(RETRY_COUNT_COL_INDEX)));
            ret.put(DISABLED_KEY, getBooleanFromCell(rowToConvert.getCell(DISABLE_STEP_COL_INDEX)));
            ret.put(EXPECT_FAILURE_KEY, getBooleanFromCell(rowToConvert.getCell(EXPECT_FAILURE_COL_INDEX)));
            // Iterate over the columns. If the index is 0-6, these are operation information. Otherwise it is data that we will use the header to convert.
            // headerInformation.get(column.getColumnIndex)) gives us the key that we will use for the JSON.

            for (int i = 7; i <= rowToConvert.getLastCellNum(); i++) {
                ArrayNode a = objectMapper.createArrayNode();
                Cell cell = rowToConvert.getCell(i);
                if (cell != null && !cell.getStringCellValue().isEmpty()) {
                    a.add(cell.getStringCellValue());
                    data.set(headerInformation.get(i), a);
                }
            }
        }
        // Split the value in the column and add the values to the array.
        // Add the array to the object.
        ret.set(DATA_KEY, data);
        return ret;
    }

    private int getIntFromCell(Cell c) {
        if (c.getCellType() == CellType.NUMERIC) {
            return (int)c.getNumericCellValue();
        } else if (c.getCellType() == CellType.STRING) {
            return Integer.parseInt(c.getStringCellValue());
        } else {
            throw new IllegalArgumentException(String.format("Unknown cell type for boolean %s", c.getCellType()));
        }
    }

    private boolean getBooleanFromCell(Cell c) {
        if (c.getCellType() == CellType.BOOLEAN) {
            return c.getBooleanCellValue();
        } else if (c.getCellType() == CellType.STRING) {
            return Boolean.parseBoolean(c.getStringCellValue());
        } else if (c.getCellType() == CellType.FORMULA) {
            return Boolean.parseBoolean(c.getCellFormula());
        } else {
            throw new IllegalArgumentException(String.format("Unknown cell type for boolean %s", c.getCellType()));
        }
    }

    Workbook loadWorkbook(Path pathToSheet) throws IOException {
        return new HSSFWorkbook(Files.newInputStream(pathToSheet));
    }
}
