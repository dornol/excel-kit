package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.InputStream;

/** Opens an input stream whose lifecycle is owned by the reader. */
@FunctionalInterface
public interface InputStreamSource {
    InputStream openStream() throws IOException;
}
