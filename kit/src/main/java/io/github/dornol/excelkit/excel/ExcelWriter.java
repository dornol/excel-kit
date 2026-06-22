package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * ExcelWriter is a utility class for generating large Excel files using Apache POI's SXSSFWorkbook.
 * Supports streaming writes, column configuration, style customization, and sheet auto-splitting.
 *
 * @author dhkim
 * @param <T> The data type of each row to be written into the Excel file
 * @since 2025-07-19
 */
public class ExcelWriter<T> extends AbstractSheetWriter<T, ExcelWriter<T>> {
    private static final Logger log = LoggerFactory.getLogger(ExcelWriter.class);
    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final SXSSFWorkbook wb;
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private int maxRows = DEFAULT_MAX_ROWS;
    private CellStyle headerStyle;
    private XSSFColor headerColor;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private @Nullable AfterDataWriter afterAllWriter;
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private final Map<String, CellStyle> headerStyleCache = new HashMap<>();
    private int headerRowIndex;
    private char @Nullable [] password;
    private @Nullable String workbookPassword;
    private @Nullable String headerFontName;
    private @Nullable Integer headerFontSize;
    private @Nullable HeaderStyleConfig headerStyleConfig;
    private @Nullable SXSSFSheet sheet;
    private @Nullable Cursor cursor;


    private static final int DEFAULT_MAX_ROWS = 1_000_000;

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
     * The {@link InitOptions} passed to the configurer contains settings that must be
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
     * @param configurer consumer that configures {@link InitOptions}
     * @param <T>        the row data type
     * @return a new ExcelWriter instance
     */
    public static <T> ExcelWriter<T> create(Consumer<InitOptions> configurer) {
        InitOptions opts = new InitOptions();
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
     * @param configurer  consumer that configures {@link InitOptions}
     * @param columnNames the column names (used as both header labels and map keys)
     * @return a new ExcelWriter with the columns registered
     * @since 0.13.0
     */
    public static ExcelWriter<Map<String, Object>> forMap(Consumer<InitOptions> configurer, String... columnNames) {
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
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>... configurers) {
        if (configurers.length > columnNames.length) {
            throw new IllegalArgumentException(
                    "configurers length (" + configurers.length
                            + ") exceeds columnNames length (" + columnNames.length + ")");
        }
        ExcelWriter<Map<String, Object>> writer = ExcelWriter.create();
        for (int i = 0; i < columnNames.length; i++) {
            String name = columnNames[i];
            Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>> cfg =
                    (i < configurers.length) ? configurers[i] : null;
            writer.column(name, map -> map.get(name), cfg);
        }
        return writer;
    }

    private ExcelWriter(InitOptions opts) {
        this.wb = new SXSSFWorkbook(opts.rowAccessWindowSize);
        ExcelColor defaultColor = ExcelColor.WHITE;
        this.headerColor = new XSSFColor(new byte[]{
                (byte) defaultColor.getR(),
                (byte) defaultColor.getG(),
                (byte) defaultColor.getB()
        });
        this.headerStyle = rebuildHeaderStyle();
    }

    private CellStyle rebuildHeaderStyle() {
        return ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize, headerStyleConfig);
    }

    /**
     * Initialization options for {@link ExcelWriter}. Passed to the configurer given to
     * {@link ExcelWriter#create(Consumer)}.
     * <p>
     * These options are restricted to settings that cannot be changed after the underlying
     * {@link SXSSFWorkbook} is constructed (currently just {@code rowAccessWindowSize}).
     * All other configuration is available as fluent methods on {@link ExcelWriter}.
     *
     * @since 0.17.0
     */
    public static final class InitOptions {
        private int rowAccessWindowSize = DEFAULT_ROW_ACCESS_WINDOW_SIZE;

        private InitOptions() {
        }

        /**
         * Sets the number of rows kept in memory by the underlying SXSSFWorkbook.
         * Higher values use more memory but reduce disk I/O; lower values are the inverse.
         * Defaults to 1000.
         * <p>
         * This must be set at construction time because POI's SXSSFWorkbook takes it as
         * a constructor argument and does not support changing it afterwards.
         *
         * @param size row access window size (must be positive)
         * @return this options object for chaining
         */
        public InitOptions rowAccessWindowSize(int size) {
            if (size <= 0) {
                throw new IllegalArgumentException("rowAccessWindowSize must be positive");
            }
            this.rowAccessWindowSize = size;
            return this;
        }
    }

    /**
     * Sets a document property on the generated Excel file.
     * <p>
     * Standard keys ({@code title}, {@code subject}, {@code author}, {@code keywords},
     * {@code description}, {@code category}) are mapped to Excel core properties.
     * Other keys are stored as custom properties.
     *
     * @param key   the property name (case-insensitive for standard keys)
     * @param value the property value
     * @return Current ExcelWriter instance for chaining
     * @since 0.17.0
     */
    public ExcelWriter<T> documentProperty(String key, String value) {
        ExcelWriteSupport.applyDocumentProperty(wb, key, value);
        return this;
    }

    /**
     * Sets the header background color. Must be called before {@link #write(Stream)}.
     * <p>
     * Use presets like {@link ExcelColor#STEEL_BLUE} or custom via {@link ExcelColor#of(int, int, int)}.
     * Defaults to {@link ExcelColor#WHITE}.
     *
     * @param color header color (must not be null)
     * @return Current ExcelWriter instance for chaining
     * @since 0.17.0
     */
    public ExcelWriter<T> headerColor(ExcelColor color) {
        if (color == null) {
            throw new IllegalArgumentException("color must not be null");
        }
        this.headerColor = new XSSFColor(new byte[]{
                (byte) color.getR(),
                (byte) color.getG(),
                (byte) color.getB()
        });
        this.headerStyle = rebuildHeaderStyle();
        return this;
    }

    /**
     * Sets the maximum number of rows per sheet before a new sheet is created.
     * Must be called before {@link #write(Stream)}. Defaults to 1,000,000.
     *
     * @param maxRows maximum rows per sheet (must be positive)
     * @return Current ExcelWriter instance for chaining
     * @since 0.17.0
     */
    public ExcelWriter<T> maxRows(int maxRows) {
        if (maxRows <= 0) {
            throw new IllegalArgumentException("maxRows must be positive");
        }
        this.maxRows = maxRows;
        return this;
    }

    /**
     * Adds a 1-based sequential row-number column (spans all sheets when auto-rolled).
     *
     * @param name the column header name (e.g. {@code "No."})
     * @return this writer for chaining
     * @since 0.16.11
     */
    public ExcelWriter<T> rowNumberColumn(String name) {
        column(name, (row, cursor) -> cursor.getCurrentTotal(),
                c -> c.type(ExcelDataType.LONG));
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
        this.afterAllWriter = afterAllWriter;
        return this;
    }

    /**
     * Sets a fixed sheet name. When sheets roll over, subsequent sheets are named
     * "{name} (2)", "{name} (3)", etc.
     *
     * @param name the base sheet name
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> sheetName(String name) {
        this.cfg.sheetNameFunction = index -> index == 0 ? name : name + " (" + (index + 1) + ")";
        return this;
    }

    /**
     * Protects the workbook structure with the given password.
     * <p>
     * When enabled, users cannot add, delete, rename, or reorder sheets.
     *
     * @param password the protection password
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> protectWorkbook(String password) {
        this.workbookPassword = password;
        return this;
    }

    /**
     * Sets the file encryption password.
     * <p>
     * When set, the resulting Excel file will be encrypted using the "agile" encryption mode,
     * and {@link ExcelHandler#writeTo(java.io.OutputStream)} will automatically
     * apply encryption — no need to pass the password to {@link ExcelHandler#writeTo(java.io.OutputStream, String)}.
     *
     * @param password the encryption password (must not be null or blank)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> password(String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.password = password.toCharArray();
        return this;
    }

    /**
     * Sets the file encryption password using a char array.
     * <p>
     * The array is copied internally; the caller may zero the original after this call.
     * Prefer this over {@link #password(String)} when the password should not linger
     * in the heap as an immutable String.
     *
     * @param password the encryption password (copied internally; must not be null or blank)
     * @return Current ExcelWriter instance for chaining
     * @since 0.17.0
     */
    public ExcelWriter<T> password(char[] password) {
        if (password == null || password.length == 0 || isBlank(password)) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.password = password.clone();
        return this;
    }

    private static boolean isBlank(char[] chars) {
        for (char c : chars) {
            if (!Character.isWhitespace(c)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Sets the header font name.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> headerFontName(String fontName) {
        this.headerFontName = fontName;
        return this;
    }

    /**
     * Sets the header font size in points.
     *
     * @param fontSize font size in points (must be positive)
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> headerFontSize(int fontSize) {
        if (fontSize <= 0) {
            throw new IllegalArgumentException("fontSize must be positive");
        }
        this.headerFontSize = fontSize;
        return this;
    }

    /**
     * Configures the global header cell style (alignment, bold, border, wrap text).
     * <p>
     * Font name, font size, and background color are set separately via
     * {@link #headerFontName(String)}, {@link #headerFontSize(int)}, and {@link #headerColor(ExcelColor)}.
     *
     * <pre>{@code
     * writer.headerStyle(h -> h
     *     .bold(true)
     *     .alignment(HorizontalAlignment.LEFT)
     *     .border(ExcelBorderStyle.MEDIUM)
     *     .wrapText(true));
     * }</pre>
     *
     * @param configurer consumer to configure header style properties
     * @return Current ExcelWriter instance for chaining
     * @since 0.17.0
     */
    public ExcelWriter<T> headerStyle(Consumer<HeaderStyleConfig> configurer) {
        HeaderStyleConfig config = new HeaderStyleConfig();
        configurer.accept(config);
        this.headerStyleConfig = config;
        this.headerStyle = rebuildHeaderStyle();
        return this;
    }

    /**
     * Adds an already-built column to the column list.
     *
     * @param column The ExcelColumn to add
     */
    void addColumn(ExcelColumn<T> column) {
        this.columns.add(column);
    }

    /**
     * Adds a column with default STRING type using a simple Function.
     * Useful for schema-based column registration.
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> column(String name, Function<T, @Nullable Object> function) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with additional configuration using a configurer consumer.
     * <p>
     * The configurer receives an {@link ExcelColumn.ExcelColumnBuilder} to set
     * column properties such as type, format, alignment, width, etc.
     *
     * <pre>{@code
     * writer.column("Price", Book::getPrice, c -> c.type(ExcelDataType.INTEGER).format("#,##0"));
     * }</pre>
     *
     * @param name        Column header name
     * @param function    Function to extract cell value from row
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> column(String name, Function<T, @Nullable Object> function,
                                  Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, (r, c) -> function.apply(r));
        if (configurer != null) {
            configurer.accept(builder);
        }
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with cursor access using an RowFunction.
     * Useful when the column value depends on row position (e.g., row number).
     *
     * @param name     Column header name
     * @param function Function to extract cell value from row with cursor access
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> column(String name, RowFunction<T, @Nullable Object> function) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Adds a column with cursor access and additional configuration.
     *
     * @param name        Column header name
     * @param function    Function to extract cell value from row with cursor access
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> column(String name, RowFunction<T, @Nullable Object> function,
                                  Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        ExcelColumn.ExcelColumnBuilder<T> builder =
                new ExcelColumn.ExcelColumnBuilder<>(this, name, function);
        if (configurer != null) {
            configurer.accept(builder);
        }
        this.columns.add(builder.build());
        return this;
    }

    /**
     * Conditionally adds a column with default STRING type using a simple Function.
     * If condition is false, the column is not added.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param function  Function to extract cell value from row
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return this;
    }

    /**
     * Conditionally adds a column with additional configuration.
     * If condition is false, the column is not added.
     *
     * @param name        Column header name
     * @param condition   Whether to include this column
     * @param function    Function to extract cell value from row
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function,
                                    Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        if (condition) {
            column(name, function, configurer);
        }
        return this;
    }

    /**
     * Conditionally adds a column with cursor access using an RowFunction.
     * If condition is false, the column is not added.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param function  Function to extract cell value from row with cursor access
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> columnIf(String name, boolean condition, RowFunction<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return this;
    }

    /**
     * Conditionally adds a column with cursor access and additional configuration.
     * If condition is false, the column is not added.
     *
     * @param name        Column header name
     * @param condition   Whether to include this column
     * @param function    Function to extract cell value from row with cursor access
     * @param configurer  Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> columnIf(String name, boolean condition, RowFunction<T, @Nullable Object> function,
                                    Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        if (condition) {
            column(name, function, configurer);
        }
        return this;
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  Column header name
     * @param value Constant value to be used in all rows
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> constColumn(String name, @Nullable Object value) {
        return column(name, (RowFunction<T, Object>) (r, c) -> value);
    }

    /**
     * Adds a column with a constant value for all rows, with additional configuration.
     *
     * @param name       Column header name
     * @param value      Constant value to be used in all rows
     * @param configurer Consumer to configure column properties
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> constColumn(String name, @Nullable Object value,
                                       Consumer<ExcelColumn.ExcelColumnBuilder<T>> configurer) {
        return column(name, (RowFunction<T, Object>) (r, c) -> value, configurer);
    }

    /**
     * Conditionally adds a column with a constant value for all rows.
     * If condition is false, the column is not added.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param value     Constant value to be used in all rows
     * @return Current ExcelWriter instance for chaining
     */
    public ExcelWriter<T> constColumnIf(String name, boolean condition, @Nullable Object value) {
        if (condition) {
            constColumn(name, value);
        }
        return this;
    }

    /**
     * Writes the stream of row data into an Excel file using custom row-level callback.
     *
     * @param stream   The data stream
     * @param consumer Custom consumer for post-processing row with cursor
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream, WriteRowCallback<T> consumer) {
        if (this.columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        if (headerFontName != null || headerFontSize != null) {
            this.headerStyle = rebuildHeaderStyle();
        }

        this.sheet = createNamedSheet();
        int headerStartRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        this.cursor = new Cursor(headerStartRow);
        this.headerRowIndex = headerStartRow;

        ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
        applySheetOptions();

        try {
            try (stream) {
                stream.forEach(rowData -> {
                    this.handleRowData(rowData);
                    consumer.accept(rowData, cursor);
                });
            }

            int nextRow = ExcelWriteSupport.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);
            if (this.afterAllWriter != null) {
                this.afterAllWriter.write(new SheetContext(sheet, wb, nextRow, columns, headerRowIndex));
            }

            applyPostProcessingAllSheets();
            ExcelWriteSupport.applyWorkbookProtection(wb, workbookPassword);

            // Apply chart on last sheet
            if (cfg.chartConfig != null) {
                ExcelWriteSupport.applyChart(sheet, cfg.chartConfig, headerRowIndex, cursor.getRowOfSheet() - 1);
            }

            return new ExcelHandler(this.wb, this.password);
        } catch (ExcelWriteException e) {
            closeWorkbookQuietly();
            throw e;
        } catch (Exception e) {
            closeWorkbookQuietly();
            throw new ExcelWriteException("Failed to write excel", e);
        }
    }

    /**
     * Writes the stream of row data into Excel file without row-level callback.
     *
     * @param stream The data stream
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler write(Stream<T> stream) {
        return this.write(stream, (rowData, consumer) -> {});
    }

    /**
     * Applies optional sheet-level settings such as auto-filter and freeze panes.
     */
    private void applySheetOptions() {
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, cfg.autoFilter, cfg.freezePaneCols, cfg.freezePaneRows, columns.size());
    }

    /**
     * Handles the logic of writing a single row to the sheet, including value mapping and style.
     *
     * @param rowData A row of data
     */
    void handleRowData(T rowData) {
        cursor.plusTotal();
        if (isOverMaxRows()) {
            ExcelWriteSupport.writeAfterDataAndSummary(sheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);
            turnOverSheet();
            int preambleRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
            cursor.setRowOfSheet(preambleRow);
            headerRowIndex = preambleRow;
            ExcelWriteSupport.writeColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
            applySheetOptions();
        }
        ExcelWriteSupport.writeRowCells(sheet, cursor, rowData, columns, cfg, rowStyleCache, wb);
        ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
    }

    /**
     * Creates a new sheet with a name determined by the sheet name function (if set).
     *
     * @return the newly created sheet
     */
    private SXSSFSheet createNamedSheet() {
        int index = wb.getNumberOfSheets();
        if (cfg.sheetNameFunction != null) {
            return wb.createSheet(cfg.sheetNameFunction.apply(index));
        }
        return wb.createSheet();
    }

    /**
     * Creates a new sheet and resets the row index when the current sheet exceeds row limit.
     */
    private void turnOverSheet() {
        this.sheet = createNamedSheet();
        this.cursor.initRow();
    }

    /**
     * Checks whether the current sheet has exceeded its max row limit.
     *
     * @return true if a sheet needs to turn over; otherwise false
     */
    private boolean isOverMaxRows() {
        return cursor.getCurrentTotal() >= maxRows && cursor.getCurrentTotal() % maxRows == 1;
    }

    /**
     * Applies all post-processing steps (column widths, validations, outlines, hiding,
     * protection, conditional formatting, print setup, tab color) to every sheet.
     */
    private void applyPostProcessingAllSheets() {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            ExcelWriteSupport.applyPostProcessing(wb.getSheetAt(i), columns, headerRowIndex, cfg);
        }
    }

    /**
     * Returns the underlying streaming workbook instance.
     *
     * @return SXSSFWorkbook instance
     */
    private void closeWorkbookQuietly() {
        try {
            wb.close();
        } catch (Exception e) {
            log.warn("Failed to close workbook after error", e);
        }
    }

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
