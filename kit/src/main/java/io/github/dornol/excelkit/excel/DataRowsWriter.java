package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * What every writer of rows has in common, whether it writes a sheet of its own ({@link SheetWriter}) or a list into a
 * sheet of a template ({@link TemplateListWriter}): the columns and the settings of how the rows are written (row
 * height and colors, progress, the default column style, a summary below the data, a table over it, what happens to a
 * cell that fails). A setting is written once here.
 * <p>
 * It cannot be extended outside this package.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete writer type, for fluent method chaining
 * @since 0.29.0
 */
public abstract class DataRowsWriter<T, SELF extends DataRowsWriter<T, SELF>> extends ColumnsWriter<T, SELF> {

    DataRowsWriter(SheetConfig<T> cfg) {
        this.cfg = cfg;
    }

    /** The settings of the rows; {@link ExcelWriter} shares those of the sheet it delegates to. */
    final SheetConfig<T> cfg;

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
     * Sets the number of data rows sampled for auto column width calculation. Rows written before the
     * header (see {@code beforeHeader}) and the header itself do not count, and the sample is taken
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

    /**
     * Registers a callback that writes custom content after all data rows on each sheet.
     */
    public SELF afterData(AfterDataWriter afterDataWriter) {
        cfg.afterDataWriter = Objects.requireNonNull(afterDataWriter, "afterDataWriter cannot be null");
        return self();
    }

    /**
     * Configures summary (footer) rows with formulas such as SUM, AVERAGE, COUNT, MIN, MAX.
     */
    public SELF summary(Consumer<ExcelSummary> configurer) {
        ExcelSummary summary = new ExcelSummary();
        configurer.accept(summary);
        cfg.summaryConfig = summary;
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

    /**
     * Sets default column styles that apply to all columns unless overridden per-column.
     */
    public SELF defaultStyle(Consumer<ColumnStyleConfig.DefaultStyleConfig<T>> configurer) {
        ColumnStyleConfig.DefaultStyleConfig<T> config = new ColumnStyleConfig.DefaultStyleConfig<>();
        configurer.accept(config);
        cfg.defaultStyleConfig = config;
        return self();
    }
}
