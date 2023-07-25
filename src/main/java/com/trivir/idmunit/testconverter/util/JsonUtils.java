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

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.util.DefaultIndenter;
import com.fasterxml.jackson.core.util.DefaultPrettyPrinter;
import com.fasterxml.jackson.core.util.Separators;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.ObjectWriter;

import java.io.IOException;

public class JsonUtils {

    private static final ObjectMapper OBJECT_MAPPER;
    private static final ObjectWriter OBJECT_WRITER;

    static {
        OBJECT_MAPPER = new ObjectMapper();
        OBJECT_MAPPER.setSerializationInclusion(JsonInclude.Include.NON_NULL);
        OBJECT_WRITER = OBJECT_MAPPER.writer(new NormalPrettyPrinter(4));
    }

    public static ObjectMapper getMapper() {
        return OBJECT_MAPPER;
    }

    public static ObjectWriter getWriter() {
        return OBJECT_WRITER;
    }


    public static class NormalPrettyPrinter extends DefaultPrettyPrinter {

        public NormalPrettyPrinter() {
            _arrayIndenter = DefaultIndenter.SYSTEM_LINEFEED_INSTANCE;
            _objectIndenter = DefaultIndenter.SYSTEM_LINEFEED_INSTANCE;
        }

        public NormalPrettyPrinter(int indentNumSpaces) {
            String indent = new String(new char[indentNumSpaces]).replace("\0", " ");
            Indenter indenter = new DefaultIndenter(indent, "\n");
            _arrayIndenter = indenter;
            _objectIndenter = indenter;
        }

        public NormalPrettyPrinter(DefaultPrettyPrinter base) {
            super(base);
        }

        @Override
        public NormalPrettyPrinter createInstance() {
            if (getClass() != NormalPrettyPrinter.class) {
                throw new IllegalStateException("Failed `createInstance()`: " + getClass().getName() + " does not override method; it has to");
            }
            return new NormalPrettyPrinter(this);
        }

        @Override
        public NormalPrettyPrinter withSeparators(Separators separators) {
            this._separators = separators;
            this._objectFieldValueSeparatorWithSpaces = separators.getObjectFieldValueSeparator() + " ";
            return this;
        }

        @Override
        public void writeEndArray(JsonGenerator g, int nrOfValues) throws IOException {
            if (!_arrayIndenter.isInline()) {
                --_nesting;
            }
            if (nrOfValues > 0) {
                _arrayIndenter.writeIndentation(g, _nesting);
            }
            g.writeRaw(']');
        }

        @Override
        public void writeEndObject(JsonGenerator g, int nrOfEntries) throws IOException {
            if (!_objectIndenter.isInline()) {
                --_nesting;
            }
            if (nrOfEntries > 0) {
                _objectIndenter.writeIndentation(g, _nesting);
            }
            g.writeRaw('}');
        }
    }
}
