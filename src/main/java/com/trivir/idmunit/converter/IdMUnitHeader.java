/*
 * Copyright (c) 2015-2020 TriVir LLC - All Rights Reserved
 *
 *  This software is proprietary and confidential.
 *  Unauthorized copying of this file, via any medium, is strictly prohibited.
 */

package com.trivir.idmunit.converter;

import java.util.Arrays;

public enum IdMUnitHeader {

    Comment("//Comment", "comment"),
    Operation("//Operation", "operation"),
    Target("//Target", "target"),
    WaitInterval("//WaitInterval", "waitInterval"),
    RetryCount("//RetryCount", "retryCount"),
    DisableStep("//DisableStep", "disabled"),
    ExpectFailure("//ExpectFailure", "expectFailure");

    private final String headerText;
    private final String jsonKey;

    IdMUnitHeader(String headerText, String jsonKey) {
        this.headerText = headerText;
        this.jsonKey = jsonKey;
    }

    public String getHeaderText() {
        return headerText;
    }

    public String getJsonKey() {
        return jsonKey;
    }

    public static IdMUnitHeader fromSheetHeader(String sheetHeader) {
        return Arrays.stream(IdMUnitHeader.values()).filter(h -> h.headerText.equals(sheetHeader)).findFirst().orElseThrow(() -> new IllegalArgumentException(String.format("Unknown sheet header %s", sheetHeader)));
    }

    public static IdMUnitHeader fromJsonKey(String jsonKeyInput) {
        return Arrays.stream(IdMUnitHeader.values()).filter(h -> h.jsonKey.equals(jsonKeyInput)).findFirst().orElseThrow(() -> new IllegalArgumentException(String.format("Unknown JSON key %s", jsonKeyInput)));
    }

    public static boolean isHeader(String columnHeaderText) {
        return Arrays.stream(IdMUnitHeader.values()).anyMatch(h -> h.headerText.equals(columnHeaderText));
    }
}
