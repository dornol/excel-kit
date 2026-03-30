package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.List;

/**
 * Builder for conditional formatting rules to apply to Excel sheets.
 * <p>
 * Supports cell-value-based rules, data bars, and icon sets.
 *
 * <pre>{@code
 * new ExcelWriter<Product>()
 *     .addColumn("Price", Product::getPrice, c -> c.type(ExcelDataType.INTEGER))
 *     .conditionalFormatting(cf -> cf
 *         .columns(1)
 *         .greaterThan("1000", ExcelColor.LIGHT_RED)
 *         .lessThan("100", ExcelColor.LIGHT_GREEN))
 *     .conditionalFormatting(cf -> cf
 *         .columns(1)
 *         .dataBar(ExcelColor.BLUE))
 *     .conditionalFormatting(cf -> cf
 *         .columns(1)
 *         .iconSet(ExcelConditionalRule.IconSetType.ARROWS_3))
 *     .write(stream)
 *     .consumeOutputStream(out);
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelConditionalRule {
    private static final Logger log = LoggerFactory.getLogger(ExcelConditionalRule.class);

    /**
     * Supported icon set types for conditional formatting.
     *
     * @since 0.9.2
     */
    public enum IconSetType {
        ARROWS_3,
        ARROWS_4,
        ARROWS_5,
        TRAFFIC_LIGHTS_3,
        SIGNS_3,
        SYMBOLS_3,
        FLAGS_3,
        RATINGS_4,
        RATINGS_5,
        QUARTERS_5
    }

    private final List<RuleEntry> rules = new ArrayList<>();
    private int @Nullable [] columnIndices;
    private int startRow = -1;
    private @Nullable ExcelColor dataBarColor;
    private @Nullable IconSetType iconSetType;

    /**
     * Sets the columns to apply conditional formatting to (0-based indices).
     *
     * @param indices 0-based column indices
     * @return this rule for chaining
     */
    public ExcelConditionalRule columns(int... indices) {
        this.columnIndices = indices;
        return this;
    }

    /**
     * Sets the starting data row (0-based).
     * If not set, defaults to the row after the header.
     *
     * @param startRow 0-based starting row
     * @return this rule for chaining
     */
    public ExcelConditionalRule startRow(int startRow) {
        this.startRow = startRow;
        return this;
    }

    /**
     * Adds a rule: cell value greater than the given value.
     */
    public ExcelConditionalRule greaterThan(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.GT, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value greater than or equal to the given value.
     */
    public ExcelConditionalRule greaterThanOrEqual(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.GE, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value less than the given value.
     */
    public ExcelConditionalRule lessThan(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.LT, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value less than or equal to the given value.
     */
    public ExcelConditionalRule lessThanOrEqual(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.LE, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value equal to the given value.
     */
    public ExcelConditionalRule equalTo(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.EQUAL, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value not equal to the given value.
     */
    public ExcelConditionalRule notEqualTo(String value, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.NOT_EQUAL, value, null, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value between the given values (inclusive).
     */
    public ExcelConditionalRule between(String value1, String value2, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.BETWEEN, value1, value2, bgColor));
        return this;
    }

    /**
     * Adds a rule: cell value not between the given values.
     */
    public ExcelConditionalRule notBetween(String value1, String value2, ExcelColor bgColor) {
        rules.add(new RuleEntry(ComparisonOperator.NOT_BETWEEN, value1, value2, bgColor));
        return this;
    }

    /**
     * Adds a data bar conditional formatting with the specified fill color.
     * <p>
     * Data bars display a gradient bar in each cell proportional to the cell's value.
     *
     * @param color the fill color of the data bar
     * @return this rule for chaining
     * @since 0.9.2
     */
    public ExcelConditionalRule dataBar(ExcelColor color) {
        this.dataBarColor = color;
        return this;
    }

    /**
     * Adds an icon set conditional formatting.
     * <p>
     * Icon sets display icons (arrows, traffic lights, flags, etc.) based on cell values.
     *
     * @param type the icon set type
     * @return this rule for chaining
     * @since 0.9.2
     */
    public ExcelConditionalRule iconSet(IconSetType type) {
        this.iconSetType = type;
        return this;
    }

    /**
     * Applies all configured rules to the given sheet.
     * Package-private, called by the writer.
     */
    void apply(SXSSFSheet sheet, int headerRowIndex, int columnCount, int lastDataRow) {
        if (rules.isEmpty() && dataBarColor == null && iconSetType == null) return;

        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        int dataStartRow = (startRow >= 0) ? startRow : headerRowIndex + 1;
        int endRow = Math.max(dataStartRow, lastDataRow);

        int[] cols = (columnIndices != null) ? columnIndices : defaultColumnRange(columnCount);

        for (int colIdx : cols) {
            CellRangeAddress[] ranges = {
                    new CellRangeAddress(dataStartRow, endRow, colIdx, colIdx)
            };

            for (RuleEntry entry : rules) {
                ConditionalFormattingRule rule = scf.createConditionalFormattingRule(
                        entry.operator, entry.value1, entry.value2);
                PatternFormatting pf = rule.createPatternFormatting();
                pf.setFillBackgroundColor(new XSSFColor(new byte[]{
                        (byte) entry.bgColor.getR(), (byte) entry.bgColor.getG(), (byte) entry.bgColor.getB()}));
                scf.addConditionalFormatting(ranges, rule);
            }

            if (dataBarColor != null) {
                applyDataBar(sheet, ranges);
            }

            if (iconSetType != null) {
                applyIconSet(sheet, ranges);
            }
        }
    }

    private void applyDataBar(SXSSFSheet sheet, CellRangeAddress[] ranges) {
        try {
            XSSFSheet xssfSheet = SXSSFSheetHelper.getXSSFSheetOrThrow(sheet);
            CTWorksheet ctSheet = xssfSheet.getCTWorksheet();
            CTConditionalFormatting cf = ctSheet.addNewConditionalFormatting();
            cf.setSqref(List.of(ranges[0].formatAsString()));

            CTCfRule ctRule = cf.addNewCfRule();
            ctRule.setType(STCfType.DATA_BAR);
            ctRule.setPriority(ctSheet.sizeOfConditionalFormattingArray());

            CTDataBar dataBar = ctRule.addNewDataBar();
            CTCfvo min = dataBar.addNewCfvo();
            min.setType(STCfvoType.MIN);
            CTCfvo max = dataBar.addNewCfvo();
            max.setType(STCfvoType.MAX);

            CTColor color = dataBar.addNewColor();
            color.setRgb(new byte[]{
                    (byte) 0xFF,
                    (byte) dataBarColor.getR(),
                    (byte) dataBarColor.getG(),
                    (byte) dataBarColor.getB()
            });
        } catch (Exception e) {
            log.warn("Failed to apply data bar conditional formatting", e);
        }
    }

    private void applyIconSet(SXSSFSheet sheet, CellRangeAddress[] ranges) {
        try {
            XSSFSheet xssfSheet = SXSSFSheetHelper.getXSSFSheetOrThrow(sheet);
            CTWorksheet ctSheet = xssfSheet.getCTWorksheet();
            CTConditionalFormatting cf = ctSheet.addNewConditionalFormatting();
            cf.setSqref(List.of(ranges[0].formatAsString()));

            CTCfRule ctRule = cf.addNewCfRule();
            ctRule.setType(STCfType.ICON_SET);
            ctRule.setPriority(ctSheet.sizeOfConditionalFormattingArray());

            CTIconSet iconSet = ctRule.addNewIconSet();
            iconSet.setIconSet(mapIconSetType(iconSetType));

            int thresholdCount = getThresholdCount(iconSetType);
            for (int i = 0; i < thresholdCount; i++) {
                CTCfvo cfvo = iconSet.addNewCfvo();
                if (i == 0) {
                    cfvo.setType(STCfvoType.MIN);
                } else {
                    cfvo.setType(STCfvoType.PERCENT);
                    cfvo.setVal(String.valueOf(i * (100 / thresholdCount)));
                }
            }
        } catch (Exception e) {
            log.warn("Failed to apply icon set conditional formatting", e);
        }
    }

    private static STIconSetType.Enum mapIconSetType(@Nullable IconSetType type) {
        if (type == null) return STIconSetType.X_3_ARROWS;
        return switch (type) {
            case ARROWS_3 -> STIconSetType.X_3_ARROWS;
            case ARROWS_4 -> STIconSetType.X_4_ARROWS;
            case ARROWS_5 -> STIconSetType.X_5_ARROWS;
            case TRAFFIC_LIGHTS_3 -> STIconSetType.X_3_TRAFFIC_LIGHTS_1;
            case SIGNS_3 -> STIconSetType.X_3_SIGNS;
            case SYMBOLS_3 -> STIconSetType.X_3_SYMBOLS;
            case FLAGS_3 -> STIconSetType.X_3_FLAGS;
            case RATINGS_4 -> STIconSetType.X_4_RATING;
            case RATINGS_5 -> STIconSetType.X_5_RATING;
            case QUARTERS_5 -> STIconSetType.X_5_QUARTERS;
        };
    }

    private static int getThresholdCount(IconSetType type) {
        return switch (type) {
            case ARROWS_3, TRAFFIC_LIGHTS_3, SIGNS_3, SYMBOLS_3, FLAGS_3 -> 3;
            case ARROWS_4, RATINGS_4 -> 4;
            case ARROWS_5, RATINGS_5, QUARTERS_5 -> 5;
        };
    }

    private int[] defaultColumnRange(int columnCount) {
        int[] result = new int[columnCount];
        for (int i = 0; i < columnCount; i++) {
            result[i] = i;
        }
        return result;
    }

    private record RuleEntry(byte operator, String value1, String value2, ExcelColor bgColor) {
    }
}
