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

import com.trivir.idmunit.cli.converter.model.OperationConfigHeader;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

public class LintMessages {

    @Getter
    private final List<String> warnings = new ArrayList<>();
    private boolean includeCellValue = false;

    public LintMessages(boolean includeCellValue) {
        this.includeCellValue = includeCellValue;
    }

    public void clear() {
        warnings.clear();
    }

    public IdmUnitTestConverterException errorTooFewSectionDelimiterRows(int numDelimiterRows) {
        return new IdmUnitTestConverterException("IdmUnit Test sheets must contain at least 3 section delimiter rows, this sheet contains " + numDelimiterRows + ".");
    }

    public void warnTooManySectionDelimiterRows(int numDelimiterRows) {
        warnings.add("IdmUnit Test sheets should contain only 3 section delimiter rows, this sheet contains " + numDelimiterRows + ".");
    }

    public void warnCellWithValueOnSectionDelimiterRow(Cell cell) {
        warnings.add(String.format("Cell %s contains a value but is on a section delimiter row, it will not be included.", cellLogValue(cell)));
    }

    public IdmUnitTestConverterException errorNoRowsInConnectorsSection() {
        return new IdmUnitTestConverterException("No rows found in the Connectors Section (a.k.a. Attribute Stacker).");
    }

    public void warnNoRowsInTestDetailsSection() {
        warnings.add("No rows found in the Test Details Section.");
    }

    public void warnNoTitle() {
        warnings.add("No title for this test specified in cell A1.");
    }

    public void warnExtraCellWithValueInTestDetailsSection(Cell cell) {
        warnings.add(String.format("Only cells A1 and A2 should contain a value in the Test Details Section, but cell %s contains a value; it will not be included.", cellLogValue(cell)));
    }

    public void warnUnknownHeaderWithOperationConfigPrefix(Cell cell, String operationConfigPrefix) {
        warnings.add(String.format("Cell %s starts with the Operation Config prefix '%s' but is not a known Operation Config option; it will not be included.", cellLogValue(cell), operationConfigPrefix));
    }

    public IdmUnitTestConverterException errorNoTargetOperationConfigHeader() {
        return new IdmUnitTestConverterException("No Operation Config Header for '" + OperationConfigHeader.TARGET.getExcelHeader() + "' is defined.");
    }

    public void warnConnectorRowWithNoName(Row row) {
        warnings.add(String.format("Row %s does not define a name for the connector under the '%s' header; it will not be included.", row.getRowNum() + 1, OperationConfigHeader.TARGET.getExcelHeader()));
    }

    public void warnConnectorAttributeUnderOperationConfigHeader(Cell cell) {
        warnings.add(String.format("Cell %s defines a connector attribute but is under a Operation Config header; it will not be included.", cellLogValue(cell)));
    }

    public void warnConnectorRowWithSameName(Row row, String connectorName, int originalRowNum) {
        warnings.add(String.format("Row %s defines a connector with the name '%s' but row %s already defined a connector with the same name; only row %s will be included.", row.getRowNum() + 1, connectorName, originalRowNum + 1, row.getRowNum() + 1));
    }

    public IdmUnitTestConverterException errorBlankOperationRow(Row row) {
        return new IdmUnitTestConverterException(String.format("No blank rows are allowed in the Operations Section, but row %s is blank.", row.getRowNum() + 1));
    }

    public void warnCommentOperationWithNoCommentDefined(Row row) {
        warnings.add(String.format("Row %s is marked as a comment operation, but its comment cell is blank.", row.getRowNum() + 1));
    }

    public void warnNonBlankCellOnCommentOperationRow(Cell cell) {
        warnings.add(String.format("Cell %s contains a value but it is on a comment operation row; it will not be included.", cellLogValue(cell)));
    }

    public IdmUnitTestConverterException errorOperationRowWithNoTargetDefined(Row row) {
        return new IdmUnitTestConverterException(String.format("Row %s has no connector specified under the '%s' Operation Config header.", row.getRowNum() + 1, OperationConfigHeader.TARGET.getExcelHeader()));
    }

    public void warnOperationDataForDuplicateAttr(Cell cell, Cell originalCell, String attrName) {
        warnings.add(String.format("Row %s contains operation data in two cells under the same connector attr '%s' (cells %s and %s); only cell %s will be included.", cell.getRowIndex() + 1, attrName, cellLogValue(originalCell), cellLogValue(cell), cellLogValue(cell)));
    }

    public void warnNonBlankCellInColumnWithNoHeader(Cell cell) {
        warnings.add(String.format("Row %s defines a value at cell %s but there is no header in that column for its target", cell.getRowIndex() + 1, cellLogValue(cell)));
    }

    public void warnNonBlankCellInRowAfterOperationSection(Cell cell) {
        warnings.add(String.format("Row %s should be blank as it is after the final section delimiter, but cell %s is not blank; it will not be included", cell.getRowIndex() + 1, cellLogValue(cell)));
    }

    private String cellLogValue(Cell cell) {
        if (!includeCellValue) {
            return cell.getAddress().toString();
        } else {
            return String.format("%s['%s']", cell.getAddress(), ExcelUtils.parseCellAsString(cell));
        }
    }
}
