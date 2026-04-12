package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.OutputStream;

/**
 * Common contract for file handlers produced by writer entry points.
 * <p>
 * A {@code FileHandler} wraps a generated file payload (an SXSSF workbook, a staged
 * CSV temp file, etc.) and writes it to a target {@link OutputStream} exactly once.
 * This lets callers treat Excel and CSV outputs polymorphically — for example,
 * a single Spring controller method can return either by wiring the handler into
 * a {@code StreamingResponseBody}.
 *
 * <h2>One-shot contract</h2>
 * Each implementation must refuse a second {@link #write(OutputStream)} call and
 * must release any temporary resources (workbooks, staging files) inside the call,
 * whether it succeeds or throws.
 *
 * <h2>Closed hierarchy</h2>
 * The shipped implementations ({@code ExcelHandler}, {@code CsvHandler}) are {@code final}.
 * {@code excel-kit} ships as an automatic module (no {@code module-info.java}), so this
 * interface is not declared {@code sealed} — but the library does not support pluggable
 * output formats beyond the ones it ships. Third-party implementations are unsupported.
 *
 * @author dhkim
 * @since 0.11.0
 */
public interface FileHandler {

    /**
     * Writes the generated file content to the given output stream.
     * <p>
     * Can only be called once per handler instance. The handler's backing resources
     * (workbook, staging file) are released before this method returns.
     *
     * @param out the destination stream
     * @throws IOException if an I/O error occurs while writing
     */
    void write(OutputStream out) throws IOException;
}
