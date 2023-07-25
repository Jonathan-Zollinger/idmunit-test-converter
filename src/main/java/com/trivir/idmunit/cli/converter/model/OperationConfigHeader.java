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

package com.trivir.idmunit.cli.converter.model;

import lombok.AllArgsConstructor;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@AllArgsConstructor
public enum OperationConfigHeader {
    COMMENT("Comment"),
    OPERATION("Operation"),
    TARGET("Target"),
    WAIT_INTERVAL("WaitInterval"),
    RETRY_COUNT("RetryCount"),
    DISABLE_STEP("DisableStep"),
    EXPECT_FAILURE("ExpectFailure"),
    IS_CRITICAL("IsCritical"),
    REPEAT_OP_RANGE("RepeatOpRange");

    public static final String PREFIX = "//";
    private static List<String> knownExcelHeaders;
    static {
        knownExcelHeaders = Arrays.stream(OperationConfigHeader.values())
            .map(OperationConfigHeader::getExcelHeader)
            .collect(Collectors.toList());
    }

    private final String excelHeaderValue;

    public String getExcelHeader() {
        return PREFIX + excelHeaderValue;
    }

    public static boolean isKnownExcelOpConfigHeader(String cellValue) {
        return knownExcelHeaders.contains(cellValue);
    }

    public static int count() {
        return knownExcelHeaders.size();
    }
}
