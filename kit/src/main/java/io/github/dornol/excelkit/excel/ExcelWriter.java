package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.core.read.RowFunction;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;
import java.util.Map;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * ExcelWriter is a utility class for generating large Excel files using Apache POI's SXSSFWorkbook.
 * Supports streaming writes, column configuration, style customization, and sheet auto-splitting.
 *
 * @author dhkim
 * @param <T> The data type of each row to be written into the Excel file
 * @since 0.1.2
 */
public class ExcelWriter<T> extends SheetWriter<T, ExcelWriter<T>> implements WorkbookOptions<ExcelWriter<T>> {
    private static final Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final SXSSFWorkbook wb;
    private final WorkbookSettings settings;
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private @Nullable AfterDataWriter afterAllWriter;
    private final Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache = new HashMap<>();
    private final Map<String, CellStyle> headerStyleCache = new HashMap<>();
    private boolean written;



    /**
     * Creates a new ExcelWriter with default initialization (white header, 1,000,000 max rows, 1000 row window).
     *
     * @param <T> the row data type
     * @return a new ExcelWriter instance
     */
    public static <T> ExcelWriter<T> create() {
        return create(opts -> {});
    }

    /**
     * Creates a new ExcelWriter with initialization options.
     * <p>
     * The {@link WorkbookInitOptions} passed to the configurer contains settings that must be
     * fixed at workbook creation time — currently only {@code rowAccessWindowSize}
     * (SXSSF's in-memory row window, which cannot be changed after the workbook exists).
     * All other configuration (header color, max rows, columns, filters, callbacks, etc.)
     * is set via fluent methods on the returned writer.
     *
     * <pre>{@code
     * ExcelWriter<User> writer = ExcelWriter.<User>create(opts -> opts
     *     .rowAccessWindowSize(500));
     * }</pre>
     *
     * @param configurer consumer that configures {@link WorkbookInitOptions}
     * @param <T>        the row data type
     * @return a new ExcelWriter instance
     */
    public static <T> ExcelWriter<T> create(Consumer<WorkbookInitOptions> configurer) {
        WorkbookInitOptions opts = new WorkbookInitOptions();
        configurer.accept(opts);
        return new ExcelWriter<>(opts);
    }

    /**
     * Creates an ExcelWriter pre-configured to write rows of {@code Map<String, Object>},
     * with one column per given column name. Each column reads its value from the map
     * by using the column name as the key.
     * <p>
     * Use this when your data is already in map form and you don't need per-column
     * customization beyond the header labels. The returned writer is a regular
     * {@link ExcelWriter}, so all of its fluent configuration methods (row height,
     * auto filter, freeze pane, sheet name, password, etc.) are available.
     *
     * <pre>{@code
     * ExcelWriter.forMap("Name", "Age", "Email")
     *     .rowHeight(22)
     *     .autoFilter(true)
     *     .write(stream)
     *     .write(out);
     * }</pre>
     *
     * @param columnNames the column names (used as both header labels and map keys)
     * @return a new ExcelWriter with the columns registered
     * @since 0.11.0
     */
    public static ExcelWriter<Map<String, Object>> forMap(String... columnNames) {
        return forMap(opts -> {}, columnNames);
    }

    /**
     * Creates a map-valued ExcelWriter with initialization options
     * (currently only {@code rowAccessWindowSize}). Header color and max rows are set
     * via fluent methods on the returned writer.
     *
     * <pre>{@code
     * ExcelWriter.forMap(
     *         opts -> opts.rowAccessWindowSize(500),
     *         "Name", "Age", "City")
     *     .headerColor(ExcelColor.STEEL_BLUE)
     *     .maxRows(500_000)
     *     .autoFilter(true)
     *     .write(stream)
     *     .write(out);
     * }</pre>
     *
     * @param configurer  consumer that configures {@link WorkbookInitOptions}
     * @param columnNames the column names (used as both header labels and map keys)
     * @return a new ExcelWriter with the columns registered
     * @since 0.13.0
     */
    public static ExcelWriter<Map<String, Object>> forMap(Consumer<WorkbookInitOptions> configurer, String... columnNames) {
        ExcelWriter<Map<String, Object>> writer = create(configurer);
        for (String name : columnNames) {
            writer.column(name, map -> map.get(name));
        }
        return writer;
    }

    /**
     * Creates an ExcelWriter pre-configured for {@code Map<String, Object>} rows, with
     * per-column configurers that can adjust type, format, styling, etc.
     * <p>
     * Each configurer applies to the column at the matching index. Extra column names
     * beyond the {@code configurers} array get no configurer (plain column).
     *
     * <pre>{@code
     * ExcelWriter.forMap(
     *         new String[]{"Name", "Price", "Date"},
     *         cfg -> cfg.bold(true),
     *         cfg -> cfg.type(ExcelDataType.INTEGER),
     *         cfg -> cfg.type(ExcelDataType.DATE))
     *     .write(stream)
     *     .write(out);
     * }</pre>
     *
     * @param columnNames the column names
     * @param configurers per-column configurers (length must not exceed {@code columnNames.length})
     * @return a new ExcelWriter with the columns registered
     * @throws IllegalArgumentException if {@code configurers.length > columnNames.length}
     * @since 0.11.0
     */
    @SafeVarargs
    public static ExcelWriter<Map<String, Object>> forMap(
            String[] columnNames,
            Consumer<ColumnConfig<Map<String, Object>>>... configurers) {
        if (configurers.length > columnNames.length) {
            throw new IllegalArgumentException(
                    "configurers length (" + configurers.length
                            + ") exceeds columnNames length (" + columnNames.length + ")");
        }
        ExcelWriter<Map<String, Object>> writer = ExcelWriter.create();
        for (int i = 0; i < columnNames.length; i++) {
            String name = columnNames[i];
            Consumer<ColumnConfig<Map<String, Object>>> cfg =
                    (i < configurers.length) ? configurers[i] : null;
            writer.column(name, map -> map.get(name), cfg);
        }
        return writer;
    }

    private ExcelWriter(WorkbookInitOptions opts) {
        StreamingOptions streaming = opts.options();
        this.wb = ExcelWorkbookSupport.streamingWorkbook(streaming);
        this.settings = new WorkbookSettings(wb);
    }

    @Override
    public ExcelWriter<T> documentProperty(String key, String value) {
        settings.documentProperty(key, value);
        return this;
    }
    @Override
    public ExcelWriter<T> headerColor(ExcelColor color) {
        settings.headerColor(color);
        return this;
    }

    /**
     * Registers a callback that writes custom content once on the last sheet after all data.
     * <p>
     * Called only once, on the last sheet, after {@code afterData} (if set).
     * Useful for writing grand totals or summary rows.
     *
     * @param afterAllWriter the callback to invoke after all data is written
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> afterAll(AfterDataWriter afterAllWriter) {
        this.afterAllWriter = Objects.requireNonNull(afterAllWriter, "afterAllWriter cannot be null");
        return this;
    }

    /**
     * Sets a fixed sheet name. When sheets roll over, subsequent sheets are named
     * "{name} (2)", "{name} (3)", etc.
     *
     * @param name the base sheet name
     * @return Current ExcelWriter instance for chaining
     * @throws IllegalArgumentException if the name is empty, contains one of {@code \ / ? * [ ] :}, or
     *                                  begins or ends with an apostrophe
     */
    public ExcelWriter<T> sheetName(String name) {
        Objects.requireNonNull(name, "name cannot be null");
        ExcelWriteSupport.requireValidSheetName(name);
        this.cfg.sheetNameFunction = index -> ExcelWriteSupport.rolloverSheetName(name, index);
        return this;
    }

    @Override
    public ExcelWriter<T> protectWorkbook(String password) {
        settings.protectWorkbook(password);
        return this;
    }

    @Override
    public ExcelWriter<T> password(String password) {
        settings.password(password);
        return this;
    }

    @Override
    public ExcelWriter<T> headerFontName(String fontName) {
        settings.headerFontName(fontName);
        return this;
    }

    @Override
    public ExcelWriter<T> headerFontSize(int fontSize) {
        settings.headerFontSize(fontSize);
        return this;
    }

    @Override
    public ExcelWriter<T> headerStyle(Consumer<HeaderStyleConfig> configurer) {
        settings.headerStyle(configurer);
        return this;
    }

    /**
     * Adds an already-built column to the column list.
     *
     * @param column The ExcelColumn to add
     */
    @Override
    void addColumn(ExcelColumn<T> column) {
        this.columns.add(column);
    }

    @Override
    ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function, @Nullable ColumnConfig<T> config) {
        return ExcelWriteSupport.buildColumn(wb, cellStyleCache, cfg.defaultStyleConfig, name, function, config);
    }

    /**
     * Writes the stream of row data into an Excel file using custom row-level callback.
     *
     * <p>
     * The stream is read but <b>not closed</b>: whoever opened it (a JPA or JDBC stream, for example)
     * closes it, typically with try-with-resources around this call.
     * <p>
     * Rows are written one at a time on the calling thread, in the stream's encounter order: a parallel
     * stream is read sequentially, so {@code parallel()} brings no speed-up here.
     *
     * @param stream   The data stream
     * @param consumer Custom consumer for post-processing row with cursor
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream, WriteRowCallback<T> consumer) {
        Objects.requireNonNull(stream, "stream cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        if (written) {
            throw new ExcelWriteException("write() has already been called on this writer; "
                    + "a writer builds one workbook, so create a new ExcelWriter for another file");
        }
        List<ExcelColumn<T>> writtenColumns = List.copyOf(columns);
        SheetConfig<T> config = new SheetConfig<>(cfg);
        SheetWriteEngine.validate(writtenColumns, config);
        // From here the workbook (and its temp files) is in use; columns also keep per-workbook state.
        written = true;
        SheetWriteEngine<T> engine = new SheetWriteEngine<>(wb, writtenColumns, config, settings.headerStyle(),
                headerStyleCache, rowStyleCache, null, index -> createNamedSheet(config, index));
        try {
            engine.start(createNamedSheet(config, 0));
        } catch (RuntimeException e) {
            // The first sheet may already own a temp file; release it, but keep the original exception.
            closeWorkbookQuietly();
            throw e;
        }

        try {
            stream.sequential().forEach(rowData -> {
                engine.row(rowData);
                consumer.accept(rowData, engine.cursor());
            });
            int nextRow = engine.finishData();
            if (this.afterAllWriter != null) {
                this.afterAllWriter.write(new SheetContext(engine.sheet(), wb, nextRow, writtenColumns, engine.headerRow()));
            }
            engine.finishSheets();
            ExcelWorkbookSupport.applyProtection(wb, settings.workbookPassword());
            return new ExcelHandler(this.wb, settings.password());
        } catch (ExcelWriteException e) {
            closeWorkbookQuietly();
            throw e;
        } catch (RuntimeException e) {
            closeWorkbookQuietly();
            throw new ExcelWriteException("Failed to write excel", e);
        } catch (Error e) {
            // An OutOfMemoryError or the like still must not leave the workbook's temp files behind.
            closeWorkbookQuietly();
            throw e;
        }
    }

    /** Creates a sheet of the list: the name function of the writer names it by its index, else POI does. */
    private SXSSFSheet createNamedSheet(SheetConfig<T> config, int index) {
        if (config.sheetNameFunction != null) {
            return wb.createSheet(config.sheetNameFunction.apply(index));
        }
        return wb.createSheet();
    }

    /**
     * Writes the stream of row data into Excel file without row-level callback.
     *
     * <p>
     * The stream is read but <b>not closed</b>: whoever opened it (a JPA or JDBC stream, for example)
     * closes it, typically with try-with-resources around this call.
     * <p>
     * Rows are written one at a time on the calling thread, in the stream's encounter order: a parallel
     * stream is read sequentially, so {@code parallel()} brings no speed-up here.
     *
     * @param stream The data stream
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream) {
        return this.write(stream, (rowData, consumer) -> {});
    }

    /** Writes rows from an Iterable without copying them. */
    public ExcelHandler write(Iterable<T> rows) {
        Objects.requireNonNull(rows, "rows cannot be null");
        return write(StreamSupport.stream(rows.spliterator(), false));
    }

    private void closeWorkbookQuietly() {
        try {
            ExcelWorkbookSupport.close(wb);
        } catch (RuntimeException | IOException e) {
            log.warn("Failed to close workbook after error", e);
        }
    }

    /**
     * Returns the underlying streaming workbook instance.
     *
     * @return SXSSFWorkbook instance
     */
    SXSSFWorkbook getWb() {
        return wb;
    }

    Map<String, CellStyle> getCellStyleCache() {
        return cellStyleCache;
    }

    ColumnStyleConfig.@Nullable DefaultStyleConfig<T> getDefaultStyleConfig() {
        return cfg.defaultStyleConfig;
    }

}
