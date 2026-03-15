package io.github.dornol.excelkit.excel;

/**
 * Holds metadata about a sheet in an Excel file.
 *
 * @param index the 0-based sheet index
 * @param name  the sheet name
 * @author dhkim
 * @since 0.6.0
 */
public record ExcelSheetInfo(int index, String name) {
}
