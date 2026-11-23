package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellError;
import io.github.dornol.excelkit.core.read.ReadReport;
import io.github.dornol.excelkit.core.read.RowError;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Stream;

/**
 * Writes the rows of an upload that failed to read into an Excel file for the person who uploaded it: the
 * original row number, the values as they were uploaded, the cells that failed shaded red with their message in
 * a cell comment, and the messages of the row in a last column.
 *
 * <pre>{@code
 * var errors = new ArrayList<RowError>();
 * ExcelReader.<Row>mapping(Row::new)
 *         .column("Name", Row::setName)
 *         .build(input)
 *         .read(row -> save(row.data()), errors::add);
 *
 * ExcelHandler report = ExcelErrorReport.create(List.of("Name", "Age")).write(errors);
 * }</pre>
 *
 * The report is built from the {@link RowError}s alone, so it holds the failed rows only, with the values the
 * reader kept ({@link RowError#rawValues()}); the original file is not read again. The headers are those of the
 * uploaded file, starting at its first column, as {@link CellError#columnIndex()} counts them (for example from
 * {@code ExcelWorkbookInspector.headers(...)}). A blank or repeated header is renamed, since a sheet's column
 * names must be unique.
 *
 * @author dhkim
 * @since 0.29.0
 */
public final class ExcelErrorReport {

    private static final String DEFAULT_ROW_COLUMN = "Row";
    private static final String DEFAULT_ERROR_COLUMN = "Error";
    private static final String DEFAULT_SHEET_NAME = "Errors";

    private final List<String> headers;
    private String rowColumnName = DEFAULT_ROW_COLUMN;
    private String errorColumnName = DEFAULT_ERROR_COLUMN;
    private String sheetName = DEFAULT_SHEET_NAME;
    private ExcelColor errorColor = ExcelColor.LIGHT_RED;

    private ExcelErrorReport(List<String> headers) {
        this.headers = headers;
    }

    /**
     * Creates a report for an upload with the given column headers.
     *
     * @param headers the header names of the uploaded file in column order; at least one
     * @return a new report
     * @throws IllegalArgumentException if there is no header
     */
    public static ExcelErrorReport create(List<String> headers) {
        Objects.requireNonNull(headers, "headers cannot be null");
        if (headers.isEmpty()) throw new IllegalArgumentException("headers must not be empty");
        var copy = new ArrayList<String>(headers.size());
        for (String header : headers) copy.add(header == null ? "" : header);
        return new ExcelErrorReport(List.copyOf(copy));
    }

    /**
     * Sets the header of the column holding the row number in the uploaded file. Default {@code "Row"}.
     *
     * @param name the header, not blank
     * @return this report
     */
    public ExcelErrorReport rowColumnName(String name) {
        this.rowColumnName = requireName(name, "rowColumnName");
        return this;
    }

    /**
     * Sets the header of the last column, which holds the messages of the row. Default {@code "Error"}.
     *
     * @param name the header, not blank
     * @return this report
     */
    public ExcelErrorReport errorColumnName(String name) {
        this.errorColumnName = requireName(name, "errorColumnName");
        return this;
    }

    /**
     * Sets the sheet name. Default {@code "Errors"}.
     *
     * @param name the sheet name, not blank
     * @return this report
     */
    public ExcelErrorReport sheetName(String name) {
        this.sheetName = requireName(name, "sheetName");
        return this;
    }

    /**
     * Sets the shading of the cells that failed and of the message cell. Default {@link ExcelColor#LIGHT_RED}.
     *
     * @param color the color
     * @return this report
     */
    public ExcelErrorReport errorColor(ExcelColor color) {
        this.errorColor = Objects.requireNonNull(color, "color cannot be null");
        return this;
    }

    /**
     * Writes the errors of a read report. The report keeps a bounded number of errors; when
     * {@link ReadReport#errorsTruncated()} is set the file holds the ones it kept.
     *
     * @param report the read report
     * @return the handler to stream the workbook with
     */
    public ExcelHandler write(ReadReport report) {
        Objects.requireNonNull(report, "report cannot be null");
        return write(report.errors());
    }

    /**
     * Writes the given row errors, in the order given.
     *
     * @param errors the failed rows
     * @return the handler to stream the workbook with
     */
    public ExcelHandler write(Iterable<RowError> errors) {
        Objects.requireNonNull(errors, "errors cannot be null");
        return writer().write(errors);
    }

    /**
     * Writes the given row errors, in the order given. The stream is read but not closed.
     *
     * @param errors the failed rows
     * @return the handler to stream the workbook with
     */
    public ExcelHandler write(Stream<RowError> errors) {
        Objects.requireNonNull(errors, "errors cannot be null");
        return writer().write(errors);
    }

    private ExcelWriter<RowError> writer() {
        List<String> names = uniqueNames();
        ExcelWriter<RowError> writer = ExcelWriter.<RowError>create().sheetName(sheetName).freezeRows(1);
        writer.column(names.get(0), (RowError e) -> e.fileRowNum() >= 0 ? e.fileRowNum() : null,
                c -> c.type(ExcelDataType.LONG));
        for (int i = 0; i < headers.size(); i++) {
            int index = i;
            writer.column(names.get(i + 1), (RowError e) -> valueAt(e, index), c -> c
                    .type(ExcelDataType.STRING)
                    .cellColor((value, e) -> cellError(e, index) != null ? errorColor : null)
                    .comment(e -> {
                        CellError cellError = cellError(e, index);
                        return cellError == null ? null : cellError.message();
                    }));
        }
        writer.column(names.get(names.size() - 1), (RowError e) -> messages(e), c -> c
                .type(ExcelDataType.STRING)
                .wrapText()
                .cellColor((value, e) -> errorColor));
        return writer;
    }

    /** The names of the columns: the row column, the headers, the error column; renamed where they repeat. */
    private List<String> uniqueNames() {
        var all = new ArrayList<String>(headers.size() + 2);
        all.add(rowColumnName);
        for (int i = 0; i < headers.size(); i++) {
            String header = headers.get(i);
            all.add(header.isBlank() ? "Column " + (i + 1) : header);
        }
        all.add(errorColumnName);
        Set<String> used = new HashSet<>();
        var unique = new ArrayList<String>(all.size());
        for (String name : all) {
            String candidate = name;
            for (int n = 2; !used.add(candidate.toLowerCase(java.util.Locale.ROOT)); n++) {
                candidate = name + " (" + n + ")";
            }
            unique.add(candidate);
        }
        return unique;
    }

    private static @Nullable String valueAt(RowError error, int index) {
        List<String> raw = error.rawValues();
        return index < raw.size() ? raw.get(index) : null;
    }

    private static @Nullable CellError cellError(RowError error, int index) {
        for (CellError cellError : error.cellErrors()) {
            if (cellError.columnIndex() == index) return cellError;
        }
        return null;
    }

    /** The messages of the row; when it has none of its own, those of its cells with the header they belong to. */
    private static String messages(RowError error) {
        if (!error.messages().isEmpty()) return String.join("\n", error.messages());
        var lines = new ArrayList<String>();
        for (CellError cellError : error.cellErrors()) {
            String header = cellError.headerName();
            lines.add(header == null || header.isBlank() ? cellError.message() : header + ": " + cellError.message());
        }
        return String.join("\n", lines);
    }

    private static String requireName(String name, String parameter) {
        Objects.requireNonNull(name, parameter + " cannot be null");
        if (name.isBlank()) throw new IllegalArgumentException(parameter + " must not be blank");
        return name;
    }
}
