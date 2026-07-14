package io.github.dornol.excelkit.excel;

import java.util.Objects;

/** Shared mutable state for SXSSF construction-time options. */
final class StreamingInitOptionsState {
    private StreamingOptions options = StreamingOptions.DEFAULT;

    void rowAccessWindowSize(int size) {
        if (size <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
        options = new StreamingOptions(size, options.compressTempFiles(), options.useSharedStrings());
    }

    void compressTempFiles(boolean enabled) {
        options = new StreamingOptions(options.rowAccessWindowSize(), enabled, options.useSharedStrings());
    }

    void useSharedStrings(boolean enabled) {
        options = new StreamingOptions(options.rowAccessWindowSize(), options.compressTempFiles(), enabled);
    }

    void streaming(StreamingOptions streamingOptions) {
        options = Objects.requireNonNull(streamingOptions, "options cannot be null");
    }

    StreamingOptions options() {
        return options;
    }
}
