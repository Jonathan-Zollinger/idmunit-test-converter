/*
 * Copyright (c) 2015-2020 TriVir LLC - All Rights Reserved
 *
 *  This software is proprietary and confidential.
 *  Unauthorized copying of this file, via any medium, is strictly prohibited.
 */

package com.trivir.idmunit.converter;

import com.fasterxml.jackson.core.util.DefaultIndenter;
import com.fasterxml.jackson.core.util.DefaultPrettyPrinter;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.ObjectWriter;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.Iterator;

import static org.junit.jupiter.api.Assertions.assertNotNull;

public class TestIdMUnitTestConverter {

    private static IdMUnitTestConverter idMUnitTestConverter;

    @BeforeAll
    static void setupClass() {
        idMUnitTestConverter = new IdMUnitTestConverter();
    }

    @Test
    void testLoadWorkbook() throws IOException {
        assertNotNull(idMUnitTestConverter.loadWorkbook(Paths.get("src/test/resources/ExampleTest.xls")));
    }

    @Test
    void testExample() throws IOException {
        Workbook workbook = idMUnitTestConverter.loadWorkbook(Paths.get("src/test/resources/ExampleTest.xls"));
        DefaultPrettyPrinter.Indenter unixIndenter = DefaultIndenter.SYSTEM_LINEFEED_INSTANCE.withLinefeed("\n");
        ObjectWriter writer = new ObjectMapper().writer(new DefaultPrettyPrinter().withObjectIndenter(unixIndenter));
        for (Iterator<Sheet> it = workbook.sheetIterator(); it.hasNext(); ) {
            Sheet s = it.next();
            ObjectNode node = idMUnitTestConverter.convertSheet(s);
            writer.writeValue(Files.newOutputStream(Paths.get("src/test/resources", s.getSheetName() + ".json")), node);
        }
    }

    @Test
    void testSOU() throws IOException {
        String fileName = "Sample SOU Test";
        Path outputDirectory = Paths.get("src/test/resources/", fileName);
        Workbook workbook = idMUnitTestConverter.loadWorkbook(Paths.get("src/test/resources/", String.format("%s.xls", fileName)));
        DefaultPrettyPrinter.Indenter unixIndenter = DefaultIndenter.SYSTEM_LINEFEED_INSTANCE.withLinefeed("\n");
        ObjectWriter writer = new ObjectMapper().writer(new DefaultPrettyPrinter().withObjectIndenter(unixIndenter));
        deleteDirectory(outputDirectory);
        Files.createDirectory(outputDirectory);
        for (Iterator<Sheet> it = workbook.sheetIterator(); it.hasNext(); ) {
            Sheet s = it.next();
            ObjectNode node = idMUnitTestConverter.convertSheet(s);
            writer.writeValue(Files.newOutputStream(outputDirectory.resolve(s.getSheetName() + ".json")), node);
        }
    }

    @Test
    void testDappsJBID() throws IOException {
        String fileName = "DappsJBIDProvisioning";
        Path outputDirectory = Paths.get("src/test/resources/", fileName);
        Workbook workbook = idMUnitTestConverter.loadWorkbook(Paths.get("src/test/resources/", String.format("%s.xls", fileName)));
        DefaultPrettyPrinter.Indenter unixIndenter = DefaultIndenter.SYSTEM_LINEFEED_INSTANCE.withLinefeed("\n");
        ObjectWriter writer = new ObjectMapper().writer(new DefaultPrettyPrinter().withObjectIndenter(unixIndenter));
        deleteDirectory(outputDirectory);
        Files.createDirectory(outputDirectory);
        for (Iterator<Sheet> it = workbook.sheetIterator(); it.hasNext(); ) {
            Sheet s = it.next();
            ObjectNode node = idMUnitTestConverter.convertSheet(s);
            writer.writeValue(Files.newOutputStream(outputDirectory.resolve(s.getSheetName() + ".json")), node);
        }
    }

    @SuppressWarnings("ResultOfMethodCallIgnored")
    static void deleteDirectory(Path directory) throws IOException {
        if (Files.exists(directory)) {
            Files.walk(directory)
                .sorted(Comparator.reverseOrder())
                .map(Path::toFile)
                .forEach(File::delete);
        }
    }
}
