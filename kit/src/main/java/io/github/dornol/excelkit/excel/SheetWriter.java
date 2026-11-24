package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.core.ProgressCallback;
import org.jspecify.annotations.Nullable;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Objects;
import java.util.function.BiPredicate;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * What {@link ExcelWriter} (one sheet, rolled over to more when it is full) and {@link ExcelSheetWriter} (a sheet of an
 * {@link ExcelWorkbook}) have in common: the columns and everything that configures a sheet (layout, callbacks, styling,
 * protection, charts, print setup, subtotals...). Only writing differs between the two.
 * <p>
 * The type is public so that a setting can be written once for both, for example a standard look for the reports of a
 * company:
 *
 * <pre>{@code
 * static <T, W extends SheetWriter<T, W>> W standardLook(W writer) {
 *     return writer.stripedRows().autoFilter(true).freezePane(0, 1);
 * }
 *
 * standardLook(ExcelWriter.<Order>create()).column("Id", Order::id);
 * standardLook(workbook.<Order>sheet("Orders")).column("Id", Order::id);
 * }</pre>
 *
 * A column setting written as a {@code Consumer<ColumnConfig<T>>} takes both as well. It cannot be extended outside
 * this package: use {@link ExcelWriter} or {@link ExcelSheetWriter}.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete writer type, for fluent method chaining
 * @author dhkim
 * @since 0.17.0
 */
@SuppressWarnings("unchecked")
public abstract class SheetWriter<T, SELF extends SheetWriter<T, SELF>> extends ColumnsWriter<T, SELF> {

    SheetWriter() { }

    final SheetConfig<T> cfg = new SheetConfig<>();

    // ── Rows per sheet and structured table ──

    /**
     * Sets the maximum number of data rows per sheet before a new sheet is created. Must be called before
     * writing. Defaults to 1,000,000. Rows written before the data (such as {@code beforeHeader} and the
     * header) and after it ({@code afterData}, summary) count towards Excel's limit of 1,048,576 rows per
     * sheet, so leave room for them. Use {@link #sheetName(Function)} to control rollover sheet naming.
     *
     * @param maxRows maximum data rows per sheet (positive, at most 1,048,575)
     * @return this writer for chaining
     * @since 0.17.0
     */
    public SELF maxRows(int maxRows) {
        cfg.maxRows = ExcelLimits.requireMaxRows(maxRows);
        return self();
    }

    /**
     * Turns the written range (header and data) into an Excel structured table with default options.
     *
     * @param name the table name (letters, digits and underscores; must not look like a cell reference)
     * @return this writer for chaining
     * @see TableOptions
     */
    public SELF table(String name) {
        return table(TableOptions.defaults(name));
    }

    /**
     * Turns the written range (header and data) into an Excel structured table.
     *
     * @param options the table name and style
     * @return this writer for chaining
     * @see TableOptions
     */
    public SELF table(TableOptions options) {
        Objects.requireNonNull(options, "options cannot be null");
        StructuredTableWriter.validateName(options.name());
        cfg.tableOptions = options;
        return self();
    }

    // ── Layout ──

    /**
     * Sets the row height for data rows in points. Defaults to 20.
     */
    public SELF rowHeight(float rowHeightInPoints) {
        if (rowHeightInPoints <= 0) {
            throw new IllegalArgumentException("rowHeightInPoints must be positive");
        }
        cfg.rowHeightInPoints = rowHeightInPoints;
        return self();
    }

    /**
     * Enables or disables auto-filter on the header row.
     */
    public SELF autoFilter(boolean autoFilter) {
        cfg.autoFilter = autoFilter;
        return self();
    }

    /**
     * Freezes the given number of rows below the header row.
     */
    public SELF freezeRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("freezePaneRows must be non-negative");
        }
        cfg.freezePaneRows = rows;
        return self();
    }

    /**
     * Freezes the given number of columns from the left edge.
     */
    public SELF freezeCols(int cols) {
        if (cols < 0) {
            throw new IllegalArgumentException("freezePaneCols must be non-negative");
        }
        cfg.freezePaneCols = cols;
        return self();
    }

    /**
     * Sets the number of columns and rows to freeze.
     */
    public SELF freezePane(int cols, int rows) {
        if (cols < 0) {
            throw new IllegalArgumentException("freezePaneCols must be non-negative");
        }
        if (rows < 0) {
            throw new IllegalArgumentException("freezePaneRows must be non-negative");
        }
        cfg.freezePaneCols = cols;
        cfg.freezePaneRows = rows;
        return self();
    }

    // ── Callbacks ──

    /**
     * Registers a callback that writes custom content before the column header row.
     */
    public SELF beforeHeader(BeforeHeaderWriter beforeHeaderWriter) {
        cfg.beforeHeaderWriter = Objects.requireNonNull(beforeHeaderWriter, "beforeHeaderWriter cannot be null");
        return self();
    }

    /**
     * Registers a callback that writes custom content after all data rows on each sheet.
     */
    public SELF afterData(AfterDataWriter afterDataWriter) {
        cfg.afterDataWriter = Objects.requireNonNull(afterDataWriter, "afterDataWriter cannot be null");
        return self();
    }

    // ── Row styling ──

    /**
     * Sets a function that determines the background color for each row.
     * <p>
     * Every distinct color needs a cell style of its own, and Apache POI looks each new style up in
     * the styles created so far, so the cost grows with the square of the number of distinct colors:
     * about 1,000 take a second, 4,000 about ten seconds, 8,000 about forty. An .xlsx workbook holds at
     * most 64,000 styles. Return a small palette (round values into a few buckets) rather than a
     * continuous gradient.
     */
    public SELF rowColor(Function<T, @Nullable ExcelColor> rowColorFunction) {
        cfg.rowColorFunction = Objects.requireNonNull(rowColorFunction, "rowColorFunction cannot be null");
        return self();
    }

    /**
     * Adds a conditional row style that applies to all cells in a row when the predicate matches.
     * <p>
     * Unlike {@link #rowColor(Function)} which only sets background color, this method supports
     * bold, font size, font color, italic, strikethrough, and background color.
     * Multiple row styles can be registered; the first matching predicate wins.
     *
     * <pre>{@code
     * writer.rowStyle(
     *     product -> product.price() > 10000,
     *     style -> style.bold(true).backgroundColor(ExcelColor.LIGHT_YELLOW)
     * );
     * }</pre>
     *
     * @param predicate  condition to test each row
     * @param configurer style configuration to apply when the condition is true
     * @return this writer for chaining
     */
    public SELF rowStyle(Predicate<T> predicate, Consumer<RowStyleConfig> configurer) {
        Objects.requireNonNull(predicate, "predicate cannot be null");
        RowStyleConfig style = new RowStyleConfig();
        configurer.accept(style);
        cfg.rowStyleEntries.add(new SheetConfig.RowStyleEntry<>(predicate, style));
        return self();
    }

    // ── Progress ──

    /**
     * Registers a progress callback that fires every {@code interval} rows.
     */
    public SELF onProgress(int interval, ProgressCallback callback) {
        if (interval <= 0) {
            throw new IllegalArgumentException("progress interval must be positive");
        }
        cfg.progressInterval = interval;
        cfg.progressCallback = Objects.requireNonNull(callback, "callback cannot be null");
        return self();
    }

    /**
     * Sets how row value extraction and cell value write failures are handled.
     * Defaults to {@link WriteErrorPolicy#FAIL_FAST}.
     *
     * @since 0.19.0
     */
    public SELF writeErrorPolicy(WriteErrorPolicy policy) {
        cfg.writeErrorPolicy = Objects.requireNonNull(policy, "policy must not be null");
        return self();
    }

    // ── Width ──

    /**
     * Sets the number of data rows sampled for auto column width calculation. Rows written before the
     * header (see {@link #beforeHeader}) and the header itself do not count, and the sample is taken
     * from the first rows of the whole write, not of each rollover sheet.
     * Defaults to 100. Set to 0 to disable.
     */
    public SELF autoWidthSampleRows(int rows) {
        if (rows < 0) {
            throw new IllegalArgumentException("autoWidthSampleRows must be non-negative");
        }
        cfg.autoWidthSampleRows = rows;
        return self();
    }

    // ── Styles ──

    /**
     * Limits the distinct cell styles this write may create for {@code cellColor}, {@code rowColor} and
     * {@code rowStyle} results, so that a function returning too many distinct colors fails fast with an
     * {@link ExcelWriteException} instead of spending minutes creating styles (the time grows with the
     * square of their number). Defaults to 64,000, the most an .xlsx workbook holds; the count includes the
     * styles the workbook already has.
     *
     * @param maxCellStyles the limit, between 1 and 64,000
     * @return this writer for chaining
     * @since 0.27.0
     */
    public SELF maxCellStyles(int maxCellStyles) {
        if (maxCellStyles < 1 || maxCellStyles > ExcelLimits.MAX_CELL_STYLES) {
            throw new IllegalArgumentException("maxCellStyles must be between 1 and "
                    + ExcelLimits.MAX_CELL_STYLES + ", got " + maxCellStyles);
        }
        cfg.maxCellStyles = maxCellStyles;
        return self();
    }

    // ── Protection ──

    /**
     * Protects the sheet(s) with the given password.
     */
    public SELF protectSheet(String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        cfg.sheetPassword = password;
        return self();
    }

    // ── Conditional formatting ──

    /**
     * Adds a conditional formatting rule.
     */
    public SELF conditionalFormatting(Consumer<ExcelConditionalRule> configurer) {
        cfg.addConditionalRule(configurer);
        return self();
    }

    // ── Chart ──

    /**
     * Configures a chart to be added after all data is written.
     */
    public SELF chart(Consumer<ExcelChartConfig> configurer) {
        ExcelChartConfig config = new ExcelChartConfig();
        configurer.accept(config);
        cfg.chartConfig = config;
        return self();
    }

    // ── Print ──

    /**
     * Configures print setup (page layout) for the sheet(s).
     */
    public SELF printSetup(Consumer<ExcelPrintSetup> configurer) {
        ExcelPrintSetup config = new ExcelPrintSetup();
        configurer.accept(config);
        config.validate();
        cfg.printSetup = config;
        return self();
    }

    // ── Striped rows ──

    /** The default stripe: a light gray that stays readable under black text. */
    private static final ExcelColor DEFAULT_STRIPE = ExcelColor.of(242, 242, 242);

    /**
     * Shades every second data row (the second, the fourth, ...) light gray, so that long rows are easier to follow
     * across a wide sheet. The first data row is not shaded. A rolled-over sheet starts the pattern again.
     * <p>
     * A stripe is the lowest-priority background: {@code cellColor}, a {@code rowStyle} color and {@code rowColor}
     * all win over it. For the stripes of an Excel table use {@code table(...)}, which brings its own.
     *
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF stripedRows() {
        return stripedRows(DEFAULT_STRIPE);
    }

    /**
     * Shades every second data row with the given color; see {@code stripedRows()}.
     *
     * @param color the background of the shaded rows
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF stripedRows(ExcelColor color) {
        cfg.stripeColor = Objects.requireNonNull(color, "color cannot be null");
        return self();
    }

    // ── Page breaks ──

    /**
     * Starts a new printed page where the data changes: a page break goes above a row when {@code startsNewPage}
     * holds for the row before it and the row itself. A rolled-over sheet breaks within itself only. It adds to
     * {@link ExcelPrintSetup#rowBreakEvery(int)}; Excel keeps at most 1026 breaks per sheet, and the write fails
     * with {@link ExcelWriteException} beyond that. Not applied by {@link TemplateListWriter}.
     *
     * @param startsNewPage {@code (previous, current) -> true} to break above {@code current}
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF pageBreakWhen(BiPredicate<T, T> startsNewPage) {
        cfg.pageBreakWhen = Objects.requireNonNull(startsNewPage, "startsNewPage cannot be null");
        return self();
    }

    /**
     * Starts a new printed page whenever the key of the row differs from the key of the row before it, for example
     * one page per department of a list sorted by department: {@code pageBreakOnChange(Employee::department)}.
     * The keys are compared with {@link Objects#equals(Object, Object)}. Otherwise as
     * {@link #pageBreakWhen(BiPredicate)}.
     *
     * @param key extracts the key of a row
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF pageBreakOnChange(Function<T, @Nullable Object> key) {
        Objects.requireNonNull(key, "key cannot be null");
        return pageBreakWhen((previous, current) -> !Objects.equals(key.apply(previous), key.apply(current)));
    }

    // ── Zoom ──

    /**
     * Sets the zoom the sheet opens with, in percent (a wide sheet is easier to take in at 80, a dense one to read at 120).
     * Every sheet of a rollover gets it. The default is Excel's 100.
     *
     * @param percent the zoom, from 10 to 400
     * @return this writer for chaining
     * @throws IllegalArgumentException if {@code percent} is outside 10..400, the range Excel accepts
     * @since 0.29.0
     */
    public SELF zoom(int percent) {
        if (percent < 10 || percent > 400) {
            throw new IllegalArgumentException("zoom must be between 10 and 400 percent, was " + percent);
        }
        cfg.zoom = percent;
        return self();
    }

    // ── Tab color ──

    /**
     * Sets the sheet tab color using RGB values.
     */
    public SELF tabColor(int r, int g, int b) {
        cfg.tabColor = ExcelColor.of(r, g, b).toRgb();
        return self();
    }

    /**
     * Sets the sheet tab color using a preset color.
     */
    public SELF tabColor(ExcelColor color) {
        return tabColor(color.getR(), color.getG(), color.getB());
    }

    // ── Default style ──

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     */
    public SELF defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        cfg.defaultStyleConfig = config;
        return self();
    }

    // ── Summary ──

    /**
     * Configures summary (footer) rows with formulas such as SUM, AVERAGE, COUNT, MIN, MAX.
     */
    public SELF summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        cfg.summaryConfig = summary;
        return self();
    }

    // ── Empty data ──

    /**
     * Writes a line below the header when there are no rows, so that an empty export says so instead of looking like a
     * broken file: the text is centered over the width of the sheet in gray italics. It is not a data row: the filter,
     * conditional formatting and charts do not cover it, and a summary below it counts nothing. A sheet with rows is
     * not touched. Cannot be combined with {@code table(...)} (the line is a merged cell, which Excel does not allow in
     * a table): the write fails with {@link ExcelWriteException}. A file with the line reads back with it as a data row,
     * so leave it off a sheet that people fill in and upload.
     *
     * @param text what the empty sheet says, for example {@code "No data"}
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF emptyMessage(String text) {
        Objects.requireNonNull(text, "text cannot be null");
        if (text.length() > ExcelLimits.MAX_CELL_TEXT) {
            throw new IllegalArgumentException("emptyMessage(...) has " + text.length() + " characters; a cell holds at most "
                    + ExcelLimits.MAX_CELL_TEXT);
        }
        cfg.emptyMessage = text;
        return self();
    }

    // ── Group subtotals ──

    /**
     * Writes a subtotal row below every run of rows that share a key, such as a department over its employees.
     * A subtotal is a {@code SUBTOTAL} formula over the rows of its group, so it follows the data when it is edited,
     * and a {@link #summary(Consumer)} over the sheet counts every row once (the summary switches to
     * {@code SUBTOTAL} too). The rows must be sorted by the key: a key that returns after another key gets a
     * subtotal of its own. The subtotal rows are not data rows: they do not count for {@code rowNumberColumn},
     * {@code stripedRows} or {@code maxRows}, and a rolled-over sheet closes its last group and starts a new one. A sheet
     * also turns over early when the subtotal rows and the summary would not fit the 1,048,576 rows Excel allows.
     * <p>
     * Cannot be combined with {@code table(...)}, {@code chart(...)} or {@code namedRange(...)}, which take the
     * rows of the sheet for one block of data: the write fails with {@link ExcelWriteException}. A file written
     * with subtotals reads back with the subtotal rows as data rows.
     *
     * <pre>{@code
     * writer.groupSubtotal(Employee::department, s -> s
     *     .label("Name", dept -> dept + " total")
     *     .keyColumn("Department")
     *     .sum("Salary").average("Salary"));
     * }</pre>
     *
     * @param key        the group of a row; consecutive rows with equal keys form a group; keys are compared with {@code equals}, so {@code 1} and {@code 1L}, or two
     *                   arrays, are different keys (return a value that compares as you mean, such as a {@code String})
     * @param configurer sets the subtotal columns and the style of the subtotal rows
     * @return this writer for chaining
     * @since 0.29.0
     */
    public SELF groupSubtotal(Function<T, @Nullable Object> key, Consumer<ExcelSubtotal> configurer) {
        Objects.requireNonNull(key, "key cannot be null");
        Objects.requireNonNull(configurer, "configurer cannot be null");
        ExcelSubtotal subtotal = new ExcelSubtotal();
        configurer.accept(subtotal);
        cfg.subtotalKey = key;
        cfg.subtotalConfig = subtotal;
        return self();
    }

    // ── Sheet naming ──

    /**
     * Sets a function that generates sheet names based on the sheet index (0-based).
     */
    public SELF sheetName(Function<Integer, String> sheetNameFunction) {
        cfg.sheetNameFunction = Objects.requireNonNull(sheetNameFunction, "sheetNameFunction cannot be null");
        return self();
    }

    // ── Header ──

    /**
     * Sets the height (in points) applied to every header row (including group header rows).
     * Pass {@code 0} to revert to Excel's default.
     *
     * @since 0.16.11
     */
    public SELF headerRowHeight(float points) {
        if (points < 0) throw new IllegalArgumentException("points must be >= 0");
        cfg.headerRowHeightInPoints = points;
        return self();
    }

    /**
     * Attaches a comment (note) to a group header cell identified by its path
     * (outermost label first). No-op if no column declares this path.
     *
     * @since 0.16.11
     */
    public SELF groupComment(String text, String... path) {
        return groupComment(ExcelCellComment.of(text), path);
    }

    /**
     * Attaches a rich comment (with author / size) to a group header cell identified by path.
     *
     * @since 0.16.11
     */
    public SELF groupComment(ExcelCellComment comment, String... path) {
        Objects.requireNonNull(comment, "comment cannot be null");
        if (path == null || path.length == 0) {
            throw new IllegalArgumentException("path must not be empty");
        }
        cfg.putGroupComment(Arrays.asList(path), comment);
        return self();
    }

    // ── Named ranges ──

    /**
     * Registers a named range for the given column. After data is written,
     * the range covers all data rows in that column (header excluded).
     * <p>
     * The named range can be referenced in formulas on other sheets
     * (e.g., {@code =SUM(PriceData)}).
     *
     * @param name        the named range name
     * @param columnIndex 0-based column index
     * @return this writer for chaining
     * @since 0.17.0
     */
    public SELF namedRange(String name, int columnIndex) {
        if (name == null || name.isBlank()) {
            throw new IllegalArgumentException("name must not be blank");
        }
        if (columnIndex < 0) {
            throw new IllegalArgumentException("columnIndex must be non-negative");
        }
        if (cfg.namedRanges == null) {
            cfg.namedRanges = new LinkedHashMap<>();
        }
        cfg.namedRanges.put(name, columnIndex);
        return self();
    }
}
