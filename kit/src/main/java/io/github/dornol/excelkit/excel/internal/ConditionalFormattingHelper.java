package io.github.dornol.excelkit.excel.internal;

import io.github.dornol.excelkit.excel.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelConditionalRule;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfRule;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCfvo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConditionalFormatting;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataBar;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconSet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCfvoType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STIconSetType;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * Isolates low-level CT XML API usage for data bar and icon set conditional formatting.
 * <p>
 * Placed in an internal package to prevent {@code org.openxmlformats} types from
 * polluting the public API surface via wildcard imports.
 *
 * @author dhkim
 * @since 0.9.2
 */
public class ConditionalFormattingHelper {
    private static final Logger log = LoggerFactory.getLogger(ConditionalFormattingHelper.class);

    private ConditionalFormattingHelper() {}

    public static void applyDataBar(XSSFSheet xssfSheet, CellRangeAddress[] ranges,
                                    ExcelColor color, @Nullable ExcelColor maxColor) {
        try {
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

            CTColor ctColor = dataBar.addNewColor();
            ctColor.setRgb(new byte[]{
                    (byte) 0xFF,
                    (byte) color.getR(),
                    (byte) color.getG(),
                    (byte) color.getB()
            });

            if (maxColor != null) {
                dataBar.setMinLength(0L);
                dataBar.setMaxLength(100L);
            }
        } catch (Exception e) {
            log.warn("Failed to apply data bar conditional formatting", e);
        }
    }

    public static void applyIconSet(XSSFSheet xssfSheet, CellRangeAddress[] ranges,
                                    ExcelConditionalRule.IconSetType iconSetType) {
        try {
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

    private static STIconSetType.Enum mapIconSetType(ExcelConditionalRule.IconSetType type) {
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

    private static int getThresholdCount(ExcelConditionalRule.IconSetType type) {
        return switch (type) {
            case ARROWS_3, TRAFFIC_LIGHTS_3, SIGNS_3, SYMBOLS_3, FLAGS_3 -> 3;
            case ARROWS_4, RATINGS_4 -> 4;
            case ARROWS_5, RATINGS_5, QUARTERS_5 -> 5;
        };
    }
}
