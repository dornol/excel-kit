package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jspecify.annotations.Nullable;

/** Encodes SpreadsheetML text consistently across the supported POI versions. */
final class ExcelCellText {
    private static final String HEX = "0123456789ABCDEF";

    private ExcelCellText() {}

    static void write(Cell cell, @Nullable String text) {
        if (text == null) {
            cell.setBlank();
            return;
        }
        String encoded = encode(text);
        if (encoded == text) {
            cell.setCellValue(text);
        } else if (cell instanceof SXSSFCell) {
            var stored = new StreamingText(text, encoded);
            stored.setString(encoded);
            writeStreaming(cell, stored);
        } else {
            cell.setCellValue(new XSSFRichTextString(encoded));
        }
    }

    /**
     * The text of a value that is written as text: {@code toString()}, except that a {@code char[]} is its characters
     * (its {@code toString()} is an object identity such as {@code [C@1b6d3586}).
     */
    static String textOf(@Nullable Object value) {
        return value instanceof char[] chars ? new String(chars) : String.valueOf(value);
    }

    static void writeStreaming(Cell cell, XSSFRichTextString text) {
        cell.setCellValue(text);
        if (text instanceof StreamingText encoded) encoded.attached = true;
    }

    static String encode(String text) {
        StringBuilder result = null;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (needsEscape(text, i)) {
                if (result == null) {
                    result = new StringBuilder(text.length() + 6);
                    result.append(text, 0, i);
                }
                result.append("_x").append(HEX.charAt(c >>> 12)).append(HEX.charAt(c >>> 8 & 15))
                        .append(HEX.charAt(c >>> 4 & 15)).append(HEX.charAt(c & 15)).append('_');
            } else if (result != null) {
                result.append(c);
            }
        }
        return result == null ? text : result.toString();
    }

    static boolean needsEscape(String text, int index) {
        if (requiresXmlEscape(text, index)) return true;
        if (text.charAt(index) != '_' || index + 6 >= text.length() || text.charAt(index + 1) != 'x') {
            return false;
        }
        for (int i = index + 2; i < index + 6; i++) {
            if (HEX.indexOf(Character.toUpperCase(text.charAt(i))) < 0) return false;
        }
        // The next escaped character's leading underscore can also complete a literal escape.
        return text.charAt(index + 6) == '_' || requiresXmlEscape(text, index + 6);
    }

    private static boolean requiresXmlEscape(String text, int index) {
        char c = text.charAt(index);
        return (c < ' ' && c != '\t' && c != '\n') || c == '\uFFFE' || c == '\uFFFF'
                || (Character.isHighSurrogate(c)
                    && (index + 1 == text.length() || !Character.isLowSurrogate(text.charAt(index + 1))))
                || (Character.isLowSurrogate(c)
                    && (index == 0 || !Character.isHighSurrogate(text.charAt(index - 1))));
    }

    /** POI validates logical length on assignment, then uses the stored text for inline serialization. */
    static final class StreamingText extends XSSFRichTextString {
        private final String logical;
        private final String encoded;
        private boolean attached;

        StreamingText(String logical, String encoded) {
            this.logical = logical;
            this.encoded = encoded;
        }

        @Override
        public String getString() {
            return attached ? encoded : logical;
        }
    }
}
