package io.github.dornol.excelkit.csv;

import com.opencsv.CSVParser;
import org.jspecify.annotations.Nullable;

import java.io.Closeable;
import java.io.IOException;
import java.io.Reader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/** Feeds physical lines to OpenCSV and preserves line endings inside quoted fields. */
final class CsvRecordReader implements Closeable {
    private final Reader input;
    private final CSVParser parser;
    private final char[] buffer = new char[8192];
    private int position;
    private int end;
    private long nextPhysicalLine = 1;
    private long recordStartLine;

    CsvRecordReader(Reader input, CSVParser parser) {
        this.input = input;
        this.parser = parser;
    }

    @Nullable String[] readNext() throws IOException {
        recordStartLine = nextPhysicalLine;
        PhysicalLine line = readLine();
        if (line == null) return null;
        String[] cells = parser.parseLineMulti(line.text());
        if (!parser.isPending()) return cells;

        var fields = new ArrayList<>(Arrays.asList(cells));
        var endings = new ArrayList<EmbeddedEnding>();
        while (parser.isPending()) {
            if (!line.ending().equals("\n")) {
                // OpenCSV has just added an LF to the pending field.
                endings.add(new EmbeddedEnding(parser.getPendingText().length() - 1, line.ending()));
            }
            line = readLine();
            if (line == null) {
                throw new CsvFormatException("Unterminated quoted CSV field in the record starting at line " + recordStartLine
                        + ". A quote inside a value must be written as \"\" within a quoted field "
                        + "(for example \"12\"\" pipe\"); a bare quote such as 12\" pipe opens a quoted field.");
            }
            cells = parser.parseLineMulti(line.text());
            if (cells.length > 0) {
                if (!endings.isEmpty()) {
                    cells[0] = restoreEndings(cells[0], endings);
                    endings.clear();
                }
                fields.addAll(Arrays.asList(cells));
            }
        }
        return fields.toArray(String[]::new);
    }

    long recordStartLine() {
        return recordStartLine;
    }

    private static String restoreEndings(String field, List<EmbeddedEnding> endings) {
        var restored = new StringBuilder(field.length());
        int start = 0;
        for (EmbeddedEnding ending : endings) {
            restored.append(field, start, ending.offset()).append(ending.text());
            start = ending.offset() + 1;
        }
        return restored.append(field, start, field.length()).toString();
    }

    private @Nullable PhysicalLine readLine() throws IOException {
        StringBuilder fragments = null;
        while (peek() >= 0) {
            int start = position;
            while (position < end && buffer[position] != '\n' && buffer[position] != '\r') {
                if (buffer[position] == '\0') throw new CsvFormatException("CSV input contains a NUL character");
                position++;
            }
            int length = position - start;
            if (position < end) {
                String text = fragments == null ? new String(buffer, start, length)
                        : fragments.append(buffer, start, length).toString();
                char terminator = buffer[position++];
                nextPhysicalLine++;
                if (terminator == '\n') return new PhysicalLine(text, "\n");
                if (peek() != '\n') return new PhysicalLine(text, "\r");
                position++;
                return new PhysicalLine(text, "\r\n");
            }
            if (fragments == null) fragments = new StringBuilder(length);
            fragments.append(buffer, start, length);
        }
        return fragments == null ? null : new PhysicalLine(fragments.toString(), "");
    }

    private int peek() throws IOException {
        while (position == end) {
            end = input.read(buffer);
            position = 0;
        }
        return end < 0 ? -1 : buffer[position];
    }

    @Override
    public void close() throws IOException {
        input.close();
    }

    private record PhysicalLine(String text, String ending) {}

    private record EmbeddedEnding(int offset, String text) {}
}
