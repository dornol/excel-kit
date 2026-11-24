package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;

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
 * Each implementation must refuse a second {@link #writeTo(OutputStream)} call and
 * must release any temporary resources (workbooks, staging files) inside the call,
 * whether it succeeds or throws.
 *
 * <h2>Releasing a handler that will not be written</h2>
 * A handler is deliberately not {@link AutoCloseable}. It is usually returned, or captured by a lambda, and written
 * later, after the method that made it has returned (a {@code StreamingResponseBody} runs after the controller).
 * With {@code try (handler)} the block would end, and {@link #discard()} would release the file, before that
 * happens. Call {@link #discard()} yourself where the handler is dropped without being written.
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
     * <p>
     * I/O errors are wrapped as unchecked exceptions (e.g., {@code ExcelWriteException},
     * {@code CsvWriteException}) so callers do not need to handle checked exceptions.
     *
     * @param out the destination stream
     */
    void writeTo(OutputStream out);

    /**
     * Releases the handler's backing resources (workbook, staging file) without writing it.
     * <p>
     * Use this when the handler will not be written, for example when the export is cancelled or the handler is
     * dropped after an error. It does nothing if the handler was already written or discarded; afterwards
     * {@code writeTo} throws. The default does nothing, for a handler that holds no resources.
     *
     * @since 0.29.0
     */
    default void discard() {
    }

    /**
     * Writes the generated file content directly to a file path.
     * <p>
     * Convenience overload that opens a buffered {@link OutputStream} to the given path,
     * delegates to {@link #writeTo(OutputStream)}, and closes the stream.
     * The one-shot contract applies — this method can only be called once.
     *
     * <pre>{@code
     * ExcelWriter.<User>create()
     *     .column("Name", User::getName)
     *     .write(stream)
     *     .writeTo(Path.of("users.xlsx"));
     * }</pre>
     *
     * @param path the destination file path
     * @since 0.14.0
     */
    default void writeTo(Path path) {
        try {
            AtomicFileWriter.write(path, this::writeTo);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to write file: " + path, e);
        }
    }
}
