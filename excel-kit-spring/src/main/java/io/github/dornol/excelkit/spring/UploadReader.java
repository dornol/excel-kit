package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.read.ReadResult;

import java.io.InputStream;
import java.util.function.Consumer;

/** Executes a configured reader against an upload stream. */
@FunctionalInterface
public interface UploadReader<T> {
    void read(InputStream inputStream, Consumer<ReadResult<T>> consumer);
}
