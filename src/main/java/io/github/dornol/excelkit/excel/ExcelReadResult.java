package io.github.dornol.excelkit.excel;

import java.util.List;

/**
 * Represents the result of reading a single row from an Excel file.
 *
 * @param <T>      The type of the parsed row data
 * @param data     The actual parsed object
 * @param success  Indicates whether the row was successfully parsed and validated
 * @param messages Any validation or processing messages (e.g. errors or warnings)
 *
 * @author dhkim
 * @since 2025-07-19
 */
public record ExcelReadResult<T>(
        T data,
        boolean success,
        List<String> messages
) {
}
