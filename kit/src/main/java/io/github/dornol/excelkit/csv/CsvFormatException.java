package io.github.dornol.excelkit.csv;

import java.io.IOException;

/** The input is not CSV the reader can take: a NUL character, or a quoted field that never ends. */
final class CsvFormatException extends IOException {
    private static final long serialVersionUID = 1L;

    CsvFormatException(String message) {
        super(message);
    }
}
