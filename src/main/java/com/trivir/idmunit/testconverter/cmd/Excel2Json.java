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

package com.trivir.idmunit.testconverter.cmd;

import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.trivir.idmunit.testconverter.converter.ExcelParser;
import com.trivir.idmunit.testconverter.converter.ExcelUtils;
import com.trivir.idmunit.testconverter.converter.IdmUnitTestConverterException;
import com.trivir.idmunit.testconverter.converter.model.IdmUnitTest;
import com.trivir.idmunit.testconverter.util.FilesUtils;
import com.trivir.idmunit.testconverter.util.JsonUtils;
import com.trivir.idmunit.testconverter.util.PicoCliValidation;
import com.trivir.idmunit.testconverter.util.ProgressBar;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.fusesource.jansi.Ansi;
import org.fusesource.jansi.AnsiConsole;
import picocli.CommandLine;

import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Properties;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static picocli.CommandLine.*;

@Command(
    name = "excel2json",
    description = "Converts IdMUnit Excel Workbooks into JSON.",
    mixinStandardHelpOptions = true,
    versionProvider = Excel2Json.ManifestVersionProvider.class,
    showDefaultValues = true
)
public class Excel2Json implements Runnable {

    public static final String TEST_FOLDER_EXTENSION = ".idmunit";

    public static final String MANIFEST_FILE_NAME = "manifest.idmunit.json";
    public static final String SCHEMA_VERSION_KEY = "schemaVersion";
    public static final String ORIGINAL_FILE_EXTENSION_KEY = "workbookType";
    public static final String SHEET_ORDER_KEY = "sheets";

    @Spec
    Model.CommandSpec spec;

    @Option(
        names = "--test-dir",
        description = "The path to the directory containing the test workbooks.",
        defaultValue = "test/org/idmunit"
    )
    private Path testDirPath;

    @Option(
        names = "--log-file",
        description = "The path to write the output errors and warnings to.",
        defaultValue = "test/test-converter.log"
    )
    private Path logFilePath;

    @Option(
        names = {"-v", "--verbose"},
        description = "Enable verbose output. Any log messages referencing a cell will include the cell's contents in addition to its address."
    )
    private boolean verbose;

    @Option(
        names = "--suffix",
        description = "Set a suffix that will be appended to the filename of all converted test workbooks.",
        defaultValue = ""
    )
    private String suffix;

    @Option(
        names = {"--ow", "--overwrite"},
        description = "Overwrite output files even if they already exist."
    )
    private boolean overwrite;

    @Option(
        names = {"-l", "--lint", "--lint-only"},
        description = "Don't output any converted files, just run the converter to check for warnings and errors in the test workbooks."
    )
    private boolean lintOnly;

    private List<Path> filePaths;
    private boolean hasAnyErrors = false;

    public static void main(String[] args) {
        // To avoid warnings about Log42 not being in classpath
        // See https://poi.apache.org/components/logging.html for more information (Specifically the Log4J SimpleLogger section)
        Properties properties = System.getProperties();
        properties.setProperty("log4j2.loggerContextFactory", "org.apache.logging.log4j.simple.SimpleLoggerContextFactory");

        AnsiConsole.systemInstall();
        CommandLine cmd = new CommandLine(new Excel2Json());
        int exitCode = cmd.execute(args);
        AnsiConsole.systemUninstall();
        System.exit(exitCode);
    }

    @Override
    public void run() {
        validate();
        ExcelParser excelParser = new ExcelParser(verbose);
        try (PrintWriter logWriter = new PrintWriter(Files.newOutputStream(logFilePath))) {
            for (Path filePath : getFilePaths()) {
                convertWorkbook(excelParser, filePath, logWriter);
            }
            if (!hasAnyErrors) {
                logWriter.println("All tests converted with no warnings or errors.");
            } else {
                String errorMessage = String.format("\nAt least one of the workbooks contained problems. See the log file '%s' for more details.", logFilePath);
                spec.commandLine().getErr().println(Ansi.ansi().render("@|yellow " + errorMessage + "|@"));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void validate() {
        PicoCliValidation.directoryExistsAndIsReadable(spec, testDirPath);
        if (!lintOnly && !overwrite) {
            Path[] pathsToCreate = getFilePaths().stream().map(this::workbookPathToIdmUnitPath).toArray(Path[]::new);
            PicoCliValidation.fileDoeNotExistOrAskOverwrite(spec, pathsToCreate);
        }
    }

    private List<Path> getFilePaths() {
        if (filePaths == null) {
            try (Stream<Path> files = Files.walk(testDirPath, 1)) {
                filePaths = files
                    .filter(Files::isRegularFile)
                    .filter(path -> path.toString().endsWith(".xls") || path.toString().endsWith(".xlsx"))
                    .sorted(Comparator.reverseOrder())
                    .collect(Collectors.toList());
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        return filePaths;
    }

    private void convertWorkbook(ExcelParser parser, Path workbookPath, PrintWriter logWriter) {
        spec.commandLine().getErr().println(workbookPath.getFileName().toString());
        boolean workbookHasErrorsYet = false;
        List<IdmUnitTest> convertedTests = new ArrayList<>();
        try (Workbook workbook = ExcelUtils.loadWorkbook(workbookPath)) {
            ProgressBar progressBar = new ProgressBar(spec.commandLine().getErr(), workbook.getNumberOfSheets());
            int totalNumWarnings = 0;
            String lastSheetName = "";
            try {
                for (Sheet sheet : workbook) {
                    lastSheetName = sheet.getSheetName();
                    convertedTests.add(parser.parseSheet(sheet));
                    totalNumWarnings += parser.getLintMessages().getWarnings().size();
                    String progressBarSuffix = "";
                    if (totalNumWarnings == 1) {
                        progressBarSuffix = "1 warning.";
                    } else if (totalNumWarnings > 1) {
                        progressBarSuffix = totalNumWarnings + " warnings.";
                    }
                    progressBar.step(Ansi.ansi().render("@|yellow " + progressBarSuffix + "|@").toString());
                    if (parser.getLintMessages().getWarnings().size() > 0) {
                        if (!workbookHasErrorsYet) {
                            if (hasAnyErrors) {
                                logWriter.println();
                            }
                            hasAnyErrors = true;
                            workbookHasErrorsYet = true;
                            logWriter.println(workbookPath.getFileName().toString());
                        }
                        logWriter.println("|-- " + sheet.getSheetName());
                    }
                    for (String lintWarning : parser.getLintMessages().getWarnings()) {
                        logWriter.println("    |-- [WARN] " + lintWarning);
                    }
                }
            } catch (IdmUnitTestConverterException e) {
                if (hasAnyErrors) {
                    logWriter.println();
                }
                logWriter.println(workbookPath.getFileName().toString());
                logWriter.println("|-- " + lastSheetName);
                logWriter.println("    |-- [ERROR] " + e.getMessage());
                progressBar.finish(Ansi.ansi().render("@|red Failed. Error in workbook.|@").toString());
                return;
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        if (lintOnly) {
            return;
        }
        Path testDirectoryPath = workbookPathToIdmUnitPath(workbookPath);
        try {
            ObjectNode manifestNode = JsonUtils.getMapper().createObjectNode();
            manifestNode.put(SCHEMA_VERSION_KEY, "1.0");
            if (workbookPath.toString().endsWith(".xls")) {
                manifestNode.put(ORIGINAL_FILE_EXTENSION_KEY, "xls");
            } else {
                manifestNode.put(ORIGINAL_FILE_EXTENSION_KEY, "xlsx");
            }
            ArrayNode sheetOrderNode = JsonUtils.getMapper().createArrayNode();
            convertedTests.forEach(x -> sheetOrderNode.add(x.getName()));
            manifestNode.set(SHEET_ORDER_KEY, sheetOrderNode);
            FilesUtils.deleteDirectoryIfExists(testDirectoryPath);
            Files.createDirectory(testDirectoryPath);
            Files.write(testDirectoryPath.resolve(MANIFEST_FILE_NAME), JsonUtils.getWriter().writeValueAsBytes(manifestNode));
            for (IdmUnitTest test : convertedTests) {
                Files.write(testDirectoryPath.resolve(test.getName() + ".json"), JsonUtils.getWriter().writeValueAsBytes(test));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private Path workbookPathToIdmUnitPath(Path workbookPath) {
        String originalFileName = workbookPath.getFileName().toString();
        String nameWithoutExtension = originalFileName.substring(0, originalFileName.lastIndexOf("."));
        return workbookPath.resolveSibling(nameWithoutExtension + suffix + TEST_FOLDER_EXTENSION);
    }

    public static class ManifestVersionProvider implements CommandLine.IVersionProvider {
        public String[] getVersion() {
            return new String[] {Excel2Json.class.getPackage().getImplementationVersion()};
        }
    }
}
