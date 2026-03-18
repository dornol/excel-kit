package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.ProgressCallback;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Function;

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

    float rowHeightInPoints = 20;
    boolean autoFilter = false;
    int freezePaneRows = 0;
    @Nullable BeforeHeaderWriter beforeHeaderWriter;
    @Nullable AfterDataWriter afterDataWriter;
    @Nullable Function<T, @Nullable ExcelColor> rowColorFunction;
    @Nullable ProgressCallback progressCallback;
    int progressInterval;
    int autoWidthSampleRows = ExcelWriteSupport.AUTO_WIDTH_SAMPLE_ROWS;
    @Nullable String sheetPassword;
    @Nullable List<ExcelConditionalRule> conditionalRules;
    @Nullable ExcelChartConfig chartConfig;
    @Nullable ExcelPrintSetup printSetup;
    int @Nullable [] tabColor;
    ColumnStyleConfig.@Nullable DefaultStyleConfig<T> defaultStyleConfig;
    @Nullable ExcelSummary summaryConfig;
    @Nullable Function<Integer, String> sheetNameFunction;

    void addConditionalRule(Consumer<ExcelConditionalRule> configurer) {
        if (conditionalRules == null) {
            conditionalRules = new ArrayList<>();
        }
        ExcelConditionalRule rule = new ExcelConditionalRule();
        configurer.accept(rule);
        conditionalRules.add(rule);
    }
}
