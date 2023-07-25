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

package com.trivir.idmunit.cli.util;

import picocli.CommandLine;
import picocli.CommandLine.Model.CommandSpec;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class PicoCliValidation {

    public static void fileExistsAndIsReadable(CommandSpec spec, Path... paths) {
        for (Path path : paths) {
            if (!Files.exists(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid file '%s': file does not exist.", path));
            }
            if (!Files.isRegularFile(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid file '%s': file is not a regular file.", path));
            }
            if (!Files.isReadable(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid file '%s': file cannot be read.", path));
            }
        }
    }

    public static void directoryExistsAndIsReadable(CommandSpec spec, Path... paths) {
        for (Path path : paths) {
            if (!Files.exists(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid directory '%s': directory does not exist.", path));
            }
            if (!Files.isDirectory(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid directory '%s': is not a directory.", path));
            }
            if (!Files.isReadable(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid directory '%s': directory cannot be read.", path));
            }
        }
    }

    public static void fileDoesNotExist(CommandSpec spec, Path... paths) {
        for (Path path : paths) {
            if (Files.exists(path)) {
                throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid file '%s': file already exists.", path));
            }
        }
    }

    public static void fileDoeNotExistOrAskOverwrite(CommandSpec spec, Path... paths) {
        List<Path> pathsToOverwrite = Arrays.stream(paths)
            .filter(Files::exists)
            .collect(Collectors.toList());
        if (pathsToOverwrite.isEmpty()) {
            return;
        }
        String input = readInput("The following file(s) already exist: '%s'.\n    Overwrite these files? y/n: ", pathsToOverwrite);
        if (!"y".equalsIgnoreCase(input)) {
            throw new CommandLine.ParameterException(spec.commandLine(), String.format("Invalid file(s) '%s': file(s) already exists and overwrite not specified.", pathsToOverwrite));
        }
    }

    private static String readInput(String fmt, Object... objects) {
        // IntelliJ's runner has no console.
        if (System.console() == null) {
            return null;
        }
        return System.console().readLine(fmt, objects);
    }
}
