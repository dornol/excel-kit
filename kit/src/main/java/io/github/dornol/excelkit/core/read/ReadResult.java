package io.github.dornol.excelkit.core.read;

import org.jspecify.annotations.Nullable;

import java.util.List;

/**
 * Represents the result of reading a single row from an Excel file.
 *
 * @param <T>      The type of the parsed row data
 * @param data     The actual parsed object (null on failure)
 * @param success  Indicates whether the row was successfully parsed and validated
 * @param errorType the category of failure, or {@code null} for successful rows
 * @param messages Any validation or processing messages (e.g. errors or warnings),
 *                 always immutable and never {@code null}
 * @param cause    The underlying exception when one is available; setter mapping
 *                 failures may have a {@code null} cause
 * @param fileRowNum 1-based physical row number in the source file, or {@code -1} if unknown
 * @param cellErrors structured cell-level errors, if available
 * @param rawValues source row values, including the stored text of numeric XLSX cells
 *
 * @author dhkim
 * @since 0.1.2
 */
public record ReadResult<T>(
        @Nullable T data,
        boolean success,
        RowError.@Nullable Type errorType,
        List<String> messages,
        @Nullable Throwable cause,
        long fileRowNum,
        List<CellError> cellErrors,
        List<String> rawValues
) {
    public ReadResult {
        errorType = success ? null : errorType == null
                ? (cause == null ? RowError.Type.VALIDATION : RowError.Type.MAPPING) : errorType;
        messages = messages == null ? List.of() : List.copyOf(messages);
        cellErrors = cellErrors == null ? List.of() : List.copyOf(cellErrors);
        rawValues = rawValues == null ? List.of() : List.copyOf(rawValues);
    }

    /**
     * Creates a successful result for a row whose position in the file is not known.
     *
     * @param data the mapped row
     * @param <T>  the row type
     */
    public static <T> ReadResult<T> success(@Nullable T data) {
        return success(data, -1, List.of(), List.of());
    }

    /**
     * Creates a successful result.
     *
     * @param data       the mapped row
     * @param fileRowNum the physical row number in the file
     * @param cellErrors cell-level errors, normally empty for a success
     * @param rawValues the row's raw cell values
     * @param <T>       the row type
     */
    public static <T> ReadResult<T> success(@Nullable T data, long fileRowNum,
                                            List<CellError> cellErrors, List<String> rawValues) {
        return new ReadResult<>(data, true, null, List.of(), null, fileRowNum, cellErrors, rawValues);
    }

    /**
     * Creates a failed result of the given type.
     *
     * @param data       the partly mapped row, if any
     * @param errorType  the kind of failure
     * @param messages   what went wrong; {@code null} is treated as empty
     * @param cause      the exception behind the failure, if any
     * @param fileRowNum the physical row number in the file
     * @param cellErrors the cells that failed
     * @param rawValues  the row's raw cell values
     * @param <T>        the row type
     */
    public static <T> ReadResult<T> failure(@Nullable T data, RowError.Type errorType,
                                            @Nullable List<String> messages,
                                            @Nullable Throwable cause, long fileRowNum,
                                            List<CellError> cellErrors, List<String> rawValues) {
        if (errorType == null) throw new NullPointerException("errorType cannot be null");
        return new ReadResult<>(data, false, errorType, messages, cause, fileRowNum, cellErrors, rawValues);
    }

    /**
     * Creates a {@link RowError.Type#MAPPING} failure.
     *
     * @param data       the partly mapped row, if any
     * @param messages   what went wrong
     * @param cause      the exception thrown while mapping, if any
     * @param fileRowNum the physical row number in the file
     * @param cellErrors the cells that failed
     * @param rawValues  the row's raw cell values
     * @param <T>        the row type
     */
    public static <T> ReadResult<T> mappingFailure(@Nullable T data, List<String> messages,
                                                   @Nullable Throwable cause, long fileRowNum,
                                                   List<CellError> cellErrors, List<String> rawValues) {
        return failure(data, RowError.Type.MAPPING, messages, cause, fileRowNum, cellErrors, rawValues);
    }

    /**
     * Creates a {@link RowError.Type#VALIDATION} failure, which has no cause.
     *
     * @param data       the mapped row that failed validation
     * @param messages   the validation messages
     * @param fileRowNum the physical row number in the file
     * @param cellErrors the cells that failed
     * @param rawValues  the row's raw cell values
     * @param <T>        the row type
     */
    public static <T> ReadResult<T> validationFailure(@Nullable T data, List<String> messages,
                                                      long fileRowNum, List<CellError> cellErrors,
                                                      List<String> rawValues) {
        return failure(data, RowError.Type.VALIDATION, messages, null, fileRowNum, cellErrors, rawValues);
    }
}
