package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.read.RowFunction;
import org.jspecify.annotations.Nullable;

import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * Shared sheet-level configuration methods for {@link ExcelWriter} and {@link ExcelSheetWriter}.
 * <p>
 * Contains all methods that delegate to {@link SheetConfig} — layout, callbacks, styling,
 * protection, charting, print setup, etc. Column registration and write logic remain in
 * the concrete subclasses because they differ between the two writers.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete writer type, for fluent method chaining
 * @author dhkim
 * @since 0.17.0
 */
@SuppressWarnings("unchecked")
abstract class AbstractSheetWriter<T, SELF extends AbstractSheetWriter<T, SELF>> {

    final SheetConfig<T> cfg = new SheetConfig<>();

    int maxRows = ExcelWriteSupport.DEFAULT_MAX_ROWS;
    @Nullable TableOptions tableOptions;

    private SELF self() {
        return (SELF) this;
    }

    // ── Columns shared by every sheet writer ──

    /**
     * Adds a column using a simple function.
     *
     * @param name     the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public abstract SELF column(String name, Function<T, @Nullable Object> function);

    /**
     * Adds a column whose function also receives the row's {@link Cursor}.
     *
     * @param name     the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public abstract SELF column(String name, RowFunction<T, @Nullable Object> function);

    /**
     * Conditionally adds a column using a simple function. If {@code condition} is false, nothing is added.
     *
     * @param name      the column header
     * @param condition whether to include this column
     * @param function  function to extract the cell value
     * @return this writer for chaining
     */
    public SELF columnIf(String name, boolean condition, Function<T, @Nullable Object> function) {
        if (condition) {
            column(name, function);
        }
        return self();
    }

    /**
     * Adds a column with a constant value for all rows.
     *
     * @param name  the column header
     * @param value the constant value written in every row
     * @return this writer for chaining
     */
    public SELF constColumn(String name, @Nullable Object value) {
        return column(name, (RowFunction<T, @Nullable Object>) (r, c) -> value);
    }

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
        this.maxRows = ExcelWriteSupport.requireMaxRows(maxRows);
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
        this.tableOptions = options;
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
     * Defaults to {@link ExcelWriteErrorPolicy#FAIL_FAST}.
     *
     * @since 0.19.0
     */
    public SELF writeErrorPolicy(ExcelWriteErrorPolicy policy) {
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
        if (maxCellStyles < 1 || maxCellStyles > ExcelWriteSupport.MAX_CELL_STYLES) {
            throw new IllegalArgumentException("maxCellStyles must be between 1 and "
                    + ExcelWriteSupport.MAX_CELL_STYLES + ", got " + maxCellStyles);
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
