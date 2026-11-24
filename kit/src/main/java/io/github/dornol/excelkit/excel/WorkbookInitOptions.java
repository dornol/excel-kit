package io.github.dornol.excelkit.excel;

import java.util.Objects;

/**
 * What can only be decided when the workbook is created: how the underlying streaming workbook (Apache POI's
 * {@code SXSSFWorkbook}) holds rows and temp files. Given to the configurer of {@link ExcelWriter#create(java.util.function.Consumer)},
 * {@link ExcelWriter#forMap(java.util.function.Consumer, String...)} and
 * {@link ExcelWorkbook#create(java.util.function.Consumer)}; everything else is a fluent method of the writer
 * ({@link WorkbookOptions}, {@link SheetWriter}).
 *
 * <pre>{@code
 * ExcelWriter.<Row>create(opts -> opts.rowAccessWindowSize(500).compressTempFiles(true));
 * ExcelWorkbook.create(opts -> opts.streaming(StreamingOptions.DEFAULT));
 * }</pre>
 *
 * @since 0.29.0 (as the {@code InitOptions} of {@code ExcelWriter} and of {@code ExcelWorkbook} since 0.17.0)
 */
public final class WorkbookInitOptions {
    private StreamingOptions options = StreamingOptions.DEFAULT;

    WorkbookInitOptions() { }

    /**
     * Sets how many rows the streaming workbook keeps in memory before it flushes the oldest to its temp file.
     *
     * @param size the number of rows (positive); the default is 1,000
     * @return this for chaining
     */
    public WorkbookInitOptions rowAccessWindowSize(int size) {
        if (size <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
        options = new StreamingOptions(size, options.compressTempFiles(), options.useSharedStrings());
        return this;
    }

    /**
     * Sets whether the temp files of the streaming workbook are compressed (smaller on disk, slower to write).
     *
     * @param enabled whether to compress them; the default is {@code false}
     * @return this for chaining
     */
    public WorkbookInitOptions compressTempFiles(boolean enabled) {
        options = new StreamingOptions(options.rowAccessWindowSize(), enabled, options.useSharedStrings());
        return this;
    }

    /**
     * Sets whether text cells go through a shared strings table (a smaller file when values repeat, but the table is
     * kept in memory, which costs heap with millions of distinct strings).
     *
     * @param enabled whether to use it; the default is {@code false}
     * @return this for chaining
     */
    public WorkbookInitOptions useSharedStrings(boolean enabled) {
        options = new StreamingOptions(options.rowAccessWindowSize(), options.compressTempFiles(), enabled);
        return this;
    }

    /**
     * Sets all of them at once.
     *
     * @param streamingOptions the options
     * @return this for chaining
     */
    public WorkbookInitOptions streaming(StreamingOptions streamingOptions) {
        options = Objects.requireNonNull(streamingOptions, "options cannot be null");
        return this;
    }

    StreamingOptions options() {
        return options;
    }
}
