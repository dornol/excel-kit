package io.github.dornol.excelkit.core.io;

import java.io.IOException;
import java.io.InputStream;

/**
 * Opens an input stream whose lifecycle is owned by the reader.
 *
 * @since 0.20.0
 */
@FunctionalInterface
public interface InputStreamSource {
    InputStream openStream() throws IOException;
}
