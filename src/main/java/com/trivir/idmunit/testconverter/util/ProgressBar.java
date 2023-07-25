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

package com.trivir.idmunit.testconverter.util;

import java.io.PrintWriter;

public class ProgressBar {

    private static final int TOTAL_LENGTH = 30;

    private final int max;
    private final PrintWriter writer;
    private int current;

    public ProgressBar(PrintWriter writer, int max) {
        this.writer = writer;
        this.max = max;
        outputCurrent("");
    }

    public void step(String suffix) {
        current++;
        if (current == max) {
            finish(suffix);
        } else {
            outputCurrent(suffix);
        }
    }

    private void outputCurrent(String suffix) {
        double currentPercentage = (double) current / max;
        int numCharacters = (int) Math.floor(currentPercentage * TOTAL_LENGTH);
        String currentProgress = repeatChar("=", numCharacters - 1) + ">";
        String remaining = repeatChar(" ", TOTAL_LENGTH - Math.max(numCharacters, 1));
        writer.printf("\r[%s%s] %s", currentProgress, remaining, suffix);
    }

    public void finish(String suffix) {
        String finalOutput = repeatChar("=", TOTAL_LENGTH - 1) + ">";
        writer.printf("\r[%s] %s\n", finalOutput, suffix);
    }

    private String repeatChar(String character, int numTimes) {
        if (numTimes <= 0) {
            return "";
        }
        return new String(new char[numTimes]).replace("\0", character);
    }
}
