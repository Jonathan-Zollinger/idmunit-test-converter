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


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;

public class ExcelUtils {

    public static Workbook loadWorkbook(Path workbookPath) throws IOException {
        if (workbookPath.toString().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
            return new XSSFWorkbook(Files.newInputStream(workbookPath));
        }
        return new HSSFWorkbook(Files.newInputStream(workbookPath));
    }

    public static Workbook createWorkbook(Path workbookPath) {
        if (workbookPath.toString().endsWith(".xlsx")) {
            return new XSSFWorkbook();
        }
        return new HSSFWorkbook();
    }

    public static String parseCellAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // This is written these ways to match the IdmUnit Core parser
                return Integer.toString(Double.valueOf(cell.getNumericCellValue()).intValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                throw new UnsupportedOperationException(String.format("Cannot parse cell %s as string", cell.getAddress()));
        }
    }

    public static Cell getCellOrBlank(Row row, int colIndex) {
        return row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
    }
}
