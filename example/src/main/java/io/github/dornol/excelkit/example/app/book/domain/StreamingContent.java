package io.github.dornol.excelkit.example.app.book.domain;

import java.io.IOException;
import java.io.OutputStream;

/**
 * Abstraction for streaming file content to an OutputStream.
 * Decouples the domain/application layer from specific file format libraries.
 */
@FunctionalInterface
public interface StreamingContent {

    void writeTo(OutputStream out) throws IOException;

}
