package io.github.dornol.excelkit.excel;

/**
 * Package-private utility methods shared by {@link ExcelReadHandler}, {@link ExcelMapReader},
 * and {@link ExcelReader} to eliminate duplicate read logic.
 *
 * @author dhkim
 */
class ExcelReadSupport {

    /** Excel maximum column count (XFD = 16,384). */
    static final int EXCEL_MAX_COLUMNS = 16_384;

    private ExcelReadSupport() {
    }

    /**
     * Converts an Excel cell reference (e.g., "C5", "AA12") to a zero-based column index.
     *
     * @param cellReference The Excel cell reference (e.g., "C5", "AA10")
     * @return The zero-based column index
     * @throws ExcelReadException if the column index exceeds the Excel maximum (XFD = 16,384)
     */
    static int getColumnIndex(String cellReference) {
        int colIdx = 0;
        for (char c : cellReference.toCharArray()) {
            if (!Character.isLetter(c)) break;
            colIdx = colIdx * 26 + (Character.toUpperCase(c) - 'A' + 1);
            if (colIdx > EXCEL_MAX_COLUMNS) {
                throw new ExcelReadException("Column index exceeds Excel maximum (XFD): " + cellReference);
            }
        }
        return colIdx - 1;
    }
}
