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
import java.util.LinkedHashMap;
import java.util.Map;

import static com.trivir.idmunit.converter.IdMUnitHeader.*;

public class IdMUnitTestConverter {

    private static final String SECTION_DIVIDER_INDICATOR = "---";

    private static final String ID_KEY = "id";
    private static final String TITLE_KEY = "title";
    private static final String DESCRIPTION_KEY = "description";
    private static final String OPERATIONS_KEY = "operations";

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
        Map<IdMUnitHeader, Integer> idmUnitHeaderMap = new HashMap<>();
        Map<String, Map<Integer, String>> headerInformationMap = new LinkedHashMap<>();
        ArrayNode operationsArray = objectMapper.createArrayNode();
        boolean headersParsed = false;
        boolean isHeaderRow = false;
        for (Iterator<Row> rowIterator = sheetToConvert.rowIterator(); rowIterator.hasNext(); ) {
            Row r = rowIterator.next();
            if (!headersParsed) {
                if (!isHeaderRow) {
                    if (r.getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
                        isHeaderRow = true;
                        //rowIterator.next();
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
                        if (IdMUnitHeader.isHeader(c.getStringCellValue())) {
                            idmUnitHeaderMap.put(IdMUnitHeader.fromSheetHeader(c.getStringCellValue()), c.getColumnIndex());
                        } else {
                            headerMap.put(c.getColumnIndex(), c.getStringCellValue());
                        }
                    }
                    headerInformationMap.put(r.getCell(idmUnitHeaderMap.get(Target)).getStringCellValue(), headerMap);
                }
            } else {
                if (r.getCell(0) != null && !r.getCell(0).getStringCellValue().equals(SECTION_DIVIDER_INDICATOR)) {
                    if (headerInformationMap.size() == 1) {
                        operationsArray.add(convertRow(r, idmUnitHeaderMap, headerInformationMap.values().stream().findFirst().get()));
                    } else {
                        operationsArray.add(convertRow(r, idmUnitHeaderMap, headerInformationMap.get(r.getCell(idmUnitHeaderMap.get(Target)).getStringCellValue())));
                    }
                } else {
                    break;
                }
            }
        }
        ret.set(OPERATIONS_KEY, operationsArray);
        // Iterate over the header rows and marshall the data so that we can use it when iterating over the operation rows.

        return ret;
    }

    ObjectNode convertRow(Row rowToConvert, Map<IdMUnitHeader, Integer> idmUnitHeaderMap, Map<Integer, String> headerInformation) {
        ObjectNode ret = objectMapper.createObjectNode();
        ret.put(Comment.getJsonKey(), rowToConvert.getCell(idmUnitHeaderMap.get(Comment)).getStringCellValue());
        ret.put(Operation.getJsonKey(), rowToConvert.getCell(idmUnitHeaderMap.get(Operation)).getStringCellValue());
        ObjectNode data = objectMapper.createObjectNode();

        if (!ret.get(Operation.getJsonKey()).asText().equals("comment")) {
            ret.put(Target.getJsonKey(), getIdMUnitHeaderStringFromCell(idmUnitHeaderMap.get(Target), rowToConvert));
            ret.put(WaitInterval.getJsonKey(), getIdMUnitHeaderIntFromCell(idmUnitHeaderMap.get(WaitInterval), rowToConvert));
            ret.put(RetryCount.getJsonKey(), getIdMUnitHeaderIntFromCell(idmUnitHeaderMap.get(RetryCount), rowToConvert));
            ret.put(DisableStep.getJsonKey(), getIdMUnitHeaderBooleanFromCell(idmUnitHeaderMap.get(DisableStep), rowToConvert));
            ret.put(ExpectFailure.getJsonKey(), getIdMUnitHeaderBooleanFromCell(idmUnitHeaderMap.get(ExpectFailure), rowToConvert));
            // Iterate over the columns. If the index is 0-6, these are operation information. Otherwise it is data that we will use the header to convert.
            // headerInformation.get(column.getColumnIndex)) gives us the key that we will use for the JSON.

            int maxIdMUnitHeaderIndex = idmUnitHeaderMap.values().stream().max(Integer::compareTo).orElseThrow(() -> new IllegalArgumentException("There are no IdMUnit headers."));

            for (int i = maxIdMUnitHeaderIndex + 1; i <= rowToConvert.getLastCellNum(); i++) {
                ArrayNode a = objectMapper.createArrayNode();
                Cell cell = rowToConvert.getCell(i);
                if (cell != null && !getStringFromCell(cell).isEmpty()) {
                    a.add(getStringFromCell(cell));
                    data.set(headerInformation.get(i), a);
                }
            }
        }
        // Split the value in the column and add the values to the array.
        // Add the array to the object.
        ret.set(DATA_KEY, data);
        return ret;
    }

    private String getIdMUnitHeaderStringFromCell(Integer columnIndex, Row rowToConvert) {
        if (columnIndex == null) {
            return "";
        } else {
            return rowToConvert.getCell(columnIndex).getStringCellValue();
        }
    }

    private int getIdMUnitHeaderIntFromCell(Integer columnIndex, Row rowToConvert) {
        if (columnIndex == null || rowToConvert.getCell(columnIndex) == null) {
            return 0;
        } else {
            return getIntFromCell(rowToConvert.getCell(columnIndex));
        }
    }

    private String getStringFromCell(Cell c) {
        if (c.getCellType() == CellType.STRING) {
            return c.getStringCellValue();
        } else if (c.getCellType() == CellType.NUMERIC) {
            return String.valueOf(c.getNumericCellValue());
        } else if (c.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(c.getBooleanCellValue());
        } else if (c.getCellType() == CellType.BLANK) {
            return "";
        } else {
            throw new IllegalArgumentException(String.format("Unknown cell type for boolean %s", c.getCellType()));
        }
    }

    private int getIntFromCell(Cell c) {
        if (c.getCellType() == CellType.NUMERIC) {
            return (int)c.getNumericCellValue();
        } else if (c.getCellType() == CellType.STRING) {
            String cellValue = c.getStringCellValue();
            if (cellValue.trim().isEmpty()) {
                return 0;
            } else {
                return Integer.parseInt(c.getStringCellValue());
            }
        } else if (c.getCellType() == CellType.BLANK) {
            return 0;
        } else {
            throw new IllegalArgumentException(String.format("Unknown cell type for boolean %s", c.getCellType()));
        }
    }

    private boolean getIdMUnitHeaderBooleanFromCell(Integer columnIndex, Row rowToConvert) {
        if (columnIndex == null) {
            return false;
        } else {
            return getBooleanFromCell(rowToConvert.getCell(columnIndex));
        }
    }

    private boolean getBooleanFromCell(Cell c) {
        if (c == null) {
            return false;
        } else if (c.getCellType() == CellType.BOOLEAN) {
            return c.getBooleanCellValue();
        } else if (c.getCellType() == CellType.STRING) {
            return Boolean.parseBoolean(c.getStringCellValue());
        } else if (c.getCellType() == CellType.FORMULA) {
            return Boolean.parseBoolean(c.getCellFormula());
        } else if (c.getCellType() == CellType.BLANK) {
            return false;
        } else {
            throw new IllegalArgumentException(String.format("Unknown cell type for boolean %s", c.getCellType()));
        }
    }

    Workbook loadWorkbook(Path pathToSheet) throws IOException {
        return new HSSFWorkbook(Files.newInputStream(pathToSheet));
    }
}
