package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.core.ProgressCallback;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.function.BiPredicate;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * Shared per-sheet configuration used by both {@link ExcelWriter} and {@link ExcelSheetWriter}.
 * <p>
 * Extracted to eliminate duplicated field declarations across the two writer classes.
 * This class is package-private; users interact with it indirectly through the writer APIs.
 *
 * @param <T> the row data type
 * @author dhkim
 * @since 0.8.2
 */
class SheetConfig<T> {

    SheetConfig() {}

    SheetConfig(SheetConfig<T> source) {
        rowHeightInPoints = source.rowHeightInPoints;
        headerRowHeightInPoints = source.headerRowHeightInPoints;
        autoFilter = source.autoFilter;
        freezePaneCols = source.freezePaneCols;
        freezePaneRows = source.freezePaneRows;
        beforeHeaderWriter = source.beforeHeaderWriter;
        afterDataWriter = source.afterDataWriter;
        rowColorFunction = source.rowColorFunction;
        rowStyleEntries.addAll(source.rowStyleEntries);
        progressCallback = source.progressCallback;
        progressInterval = source.progressInterval;
        autoWidthSampleRows = source.autoWidthSampleRows;
        maxCellStyles = source.maxCellStyles;
        sheetPassword = source.sheetPassword;
        conditionalRules = source.conditionalRules == null ? null : List.copyOf(source.conditionalRules);
        chartConfig = source.chartConfig;
        printSetup = source.printSetup;
        tabColor = source.tabColor == null ? null : source.tabColor.clone();
        zoom = source.zoom;
        stripeColor = source.stripeColor;
        pageBreakWhen = source.pageBreakWhen;
        visibility = source.visibility;
        defaultStyleConfig = source.defaultStyleConfig;
        summaryConfig = source.summaryConfig;
        subtotalKey = source.subtotalKey;
        subtotalConfig = source.subtotalConfig;
        emptyMessage = source.emptyMessage;
        sheetNameFunction = source.sheetNameFunction;
        groupComments = source.groupComments == null ? null : Map.copyOf(source.groupComments);
        namedRanges = source.namedRanges == null ? null : Map.copyOf(source.namedRanges);
        writeErrorPolicy = source.writeErrorPolicy;
        maxRows = source.maxRows;
        tableOptions = source.tableOptions;
    }

    static final float DEFAULT_ROW_HEIGHT_POINTS = 20f;

    float rowHeightInPoints = DEFAULT_ROW_HEIGHT_POINTS;
    /** Per-header-row height in points; 0 means use Excel default. */
    float headerRowHeightInPoints = 0f;
    boolean autoFilter = false;
    int freezePaneCols = 0;
    int freezePaneRows = 0;
    @Nullable BeforeHeaderWriter beforeHeaderWriter;
    @Nullable AfterDataWriter afterDataWriter;
    @Nullable Function<T, @Nullable ExcelColor> rowColorFunction;
    final List<RowStyleEntry<T>> rowStyleEntries = new ArrayList<>();
    @Nullable ProgressCallback progressCallback;
    int progressInterval;
    int autoWidthSampleRows = ExcelLimits.AUTO_WIDTH_SAMPLE_ROWS;
    int maxCellStyles = ExcelLimits.MAX_CELL_STYLES;
    @Nullable String sheetPassword;
    @Nullable List<ExcelConditionalRule> conditionalRules;
    @Nullable ExcelChartConfig chartConfig;
    @Nullable ExcelPrintSetup printSetup;
    int @Nullable [] tabColor;
    /** The zoom in percent, or 0 for Excel's default (100). */
    int zoom;
    /** The background of every second data row, or {@code null} for no stripes. */
    @Nullable ExcelColor stripeColor;
    /** Whether a page break goes above a row, given the row before it and the row itself. */
    @Nullable BiPredicate<T, T> pageBreakWhen;
    /** {@code HIDDEN} or {@code VERY_HIDDEN}; {@code null} keeps the sheet visible. Set by a workbook's sheet writer only. */
    org.apache.poi.ss.usermodel.@Nullable SheetVisibility visibility;
    ColumnStyleConfig.@Nullable DefaultStyleConfig<T> defaultStyleConfig;
    @Nullable ExcelSummary summaryConfig;
    /** The key whose changes end a group of rows with a subtotal row, and what the row holds. */
    @Nullable Function<T, @Nullable Object> subtotalKey;
    @Nullable ExcelSubtotal subtotalConfig;
    /** What a sheet says when no row was written, or {@code null} for the header alone. */
    @Nullable String emptyMessage;
    @Nullable Function<Integer, String> sheetNameFunction;
    /** Map from group header path (outermost-first) to its comment. */
    @Nullable Map<List<String>, ExcelCellComment> groupComments;
    /** Named ranges: name → column index. Applied after data is written. */
    @Nullable Map<String, Integer> namedRanges;
    WriteErrorPolicy writeErrorPolicy = WriteErrorPolicy.FAIL_FAST;
    int maxRows = ExcelLimits.DEFAULT_MAX_ROWS;
    @Nullable TableOptions tableOptions;

    record RowStyleEntry<T>(Predicate<T> predicate, RowStyleConfig style) {}

    void putGroupComment(List<String> path, ExcelCellComment comment) {
        if (groupComments == null) {
            groupComments = new HashMap<>();
        }
        groupComments.put(List.copyOf(path), comment);
    }

    void addConditionalRule(Consumer<ExcelConditionalRule> configurer) {
        if (conditionalRules == null) {
            conditionalRules = new ArrayList<>();
        }
        ExcelConditionalRule rule = new ExcelConditionalRule();
        configurer.accept(rule);
        conditionalRules.add(rule);
    }
}
