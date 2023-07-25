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

package com.trivir.idmunit.cli.cmd;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.TextNode;
import com.trivir.idmunit.cli.converter.ExcelUtils;
import com.trivir.idmunit.cli.converter.ExcelWriter;
import com.trivir.idmunit.cli.converter.model.IdmUnitTest;
import com.trivir.idmunit.cli.util.JsonUtils;
import com.trivir.idmunit.cli.util.PicoCliValidation;
import com.trivir.idmunit.cli.util.ProgressBar;
import org.apache.poi.ss.usermodel.*;
import org.fusesource.jansi.AnsiConsole;
import picocli.CommandLine;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
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
    name = "json2excel",
    description = "Converts IdMUnit JSON tests to Excel Workbooks.",
    mixinStandardHelpOptions = true,
    versionProvider = Json2Excel.ManifestVersionProvider.class
)
public class Json2Excel implements Runnable {

    @Spec
    Model.CommandSpec spec;

    @Option(
        names = "--test-dir",
        description = "The path to the directory containing the test workbooks.",
        defaultValue = "test/org/idmunit"
    )
    private Path testDirPath;

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

    private List<Path> filePaths;

    public static void main(String[] args) {
        // To avoid warnings about Log42 not being in classpath
        // See https://poi.apache.org/components/logging.html for more information (Specifically the Log4J SimpleLogger section)
        Properties properties = System.getProperties();
        properties.setProperty("log4j2.loggerContextFactory", "org.apache.logging.log4j.simple.SimpleLoggerContextFactory");

        AnsiConsole.systemInstall();
        CommandLine cmd = new CommandLine(new Json2Excel());
        int exitCode = cmd.execute(args);
        AnsiConsole.systemUninstall();
        System.exit(exitCode);
    }

    @Override
    public void run() {
        validate();
        ProgressBar progressBar = new ProgressBar(spec.commandLine().getOut(), getFilePaths().size());
        for (Path idmunitDirPath : getFilePaths()) {
            Path workbookPath = idmunitDirPathToWorkbookPath(idmunitDirPath);
            progressBar.step(workbookPath.getFileName().toString());
            Path manifestPath = idmunitDirPath.resolve(Excel2Json.MANIFEST_FILE_NAME);
            try (Workbook workbook = ExcelUtils.createWorkbook(workbookPath);
                 InputStream inputStream = Files.newInputStream(manifestPath);
                 OutputStream outputStream = Files.newOutputStream(workbookPath)) {
                ExcelWriter writer = new ExcelWriter(workbook);
                ArrayNode sheetOrderNode = (ArrayNode) JsonUtils.getMapper().readTree(inputStream).get(Excel2Json.SHEET_ORDER_KEY);
                if (sheetOrderNode == null) {
                    throw new RuntimeException(String.format("Failed to read sheet order from '%s', for test '%s'.", Excel2Json.MANIFEST_FILE_NAME, workbookPath));
                }
                List<String> sheetNames = new ArrayList<>();
                for (JsonNode textNode : sheetOrderNode) {
                    sheetNames.add(textNode.asText());
                }
                sheetNames.stream()
                    .map(x -> {
                        Path sheetJsonPath = idmunitDirPath.resolve(x + ".json");
                        try (InputStream is = Files.newInputStream(sheetJsonPath)) {
                            return JsonUtils.getMapper().readValue(is, IdmUnitTest.class);
                        } catch (IOException e) {
                            throw new RuntimeException(e);
                        }
                    })
                    .forEach(writer::writeTest);
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                for (Sheet sheet : workbook) {
                    for (Row r : sheet) {
                        for (Cell c : r) {
                            if (c.getCellType() == CellType.FORMULA) {
                                evaluator.evaluateFormulaCell(c);
                            }
                        }
                    }
                }
                workbook.write(outputStream);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private void validate() {
        PicoCliValidation.directoryExistsAndIsReadable(spec, testDirPath);
        if (!overwrite) {
            Path[] pathsToCreate = getFilePaths().stream().map(this::idmunitDirPathToWorkbookPath).toArray(Path[]::new);
            PicoCliValidation.fileDoeNotExistOrAskOverwrite(spec, pathsToCreate);
        }
    }

    private List<Path> getFilePaths() {
        if (filePaths == null) {
            try (Stream<Path> files = Files.walk(testDirPath, 1)) {
                filePaths = files
                    .filter(Files::isDirectory)
                    .filter(path -> path.toString().endsWith(Excel2Json.TEST_FOLDER_EXTENSION))
                    .sorted(Comparator.reverseOrder())
                    .collect(Collectors.toList());
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
        return filePaths;
    }

    private Path idmunitDirPathToWorkbookPath(Path idmunitDirPath) {
        String originalWorkbookType = getOriginalWorkbookType(idmunitDirPath);
        String originalDirName = idmunitDirPath.getFileName().toString();
        String nameWithoutExtension = originalDirName.substring(0, originalDirName.lastIndexOf("."));
        return idmunitDirPath.resolveSibling(nameWithoutExtension + suffix + "." + originalWorkbookType);
    }

    private String getOriginalWorkbookType(Path idmunitDirPath) {
        Path manifestFilePath = idmunitDirPath.resolve(Excel2Json.MANIFEST_FILE_NAME);
        try {
            byte[] manifestContents = Files.readAllBytes(manifestFilePath);
            TextNode workbookTypeNode = (TextNode) JsonUtils.getMapper().readTree(manifestContents).get(Excel2Json.ORIGINAL_FILE_EXTENSION_KEY);
            if (workbookTypeNode == null) {
                throw new RuntimeException(String.format("Failed to read original workbook type from '%s' for test '%s'.", Excel2Json.MANIFEST_FILE_NAME, idmunitDirPath));
            }
            return workbookTypeNode.asText();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static class ManifestVersionProvider implements CommandLine.IVersionProvider {
        public String[] getVersion() {
            return new String[] {Excel2Json.class.getPackage().getImplementationVersion()};
        }
    }
}
