package io.github.dornol.excelkit.excel;

import java.util.Objects;

/** Shared implementation for SXSSF construction-time options. */
abstract class AbstractStreamingInitOptions<SELF extends AbstractStreamingInitOptions<SELF>> {
    private StreamingOptions streamingOptions = StreamingOptions.DEFAULT;

    protected abstract SELF self();

    public final SELF rowAccessWindowSize(int size) {
        if (size <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
        streamingOptions = new StreamingOptions(size, streamingOptions.compressTempFiles(),
                streamingOptions.useSharedStrings());
        return self();
    }

    public final SELF compressTempFiles(boolean enabled) {
        streamingOptions = new StreamingOptions(streamingOptions.rowAccessWindowSize(), enabled,
                streamingOptions.useSharedStrings());
        return self();
    }

    public final SELF useSharedStrings(boolean enabled) {
        streamingOptions = new StreamingOptions(streamingOptions.rowAccessWindowSize(),
                streamingOptions.compressTempFiles(), enabled);
        return self();
    }

    public final SELF streaming(StreamingOptions options) {
        streamingOptions = Objects.requireNonNull(options, "options cannot be null");
        return self();
    }

    final StreamingOptions streamingOptions() {
        return streamingOptions;
    }
}
