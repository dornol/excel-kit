package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.WeakHashMap;
import java.util.function.Consumer;

/**
 * Fluent builder for creating rich text content with mixed formatting within a single Excel cell.
 * <p>
 * Supports partial bold, italic, underline, strikethrough, font color, and font size
 * applied to individual text segments.
 * <p>
 * Example usage:
 * <pre>{@code
 * ExcelRichText rt = new ExcelRichText()
 *     .text("Hello ")
 *     .bold("World")
 *     .text(" — ")
 *     .styled("red text", s -> s.color(ExcelColor.RED).italic(true));
 * }</pre>
 *
 * @author dhkim
 * @see ExcelDataType#RICH_TEXT
 */
public class ExcelRichText {

    private static final Map<SXSSFWorkbook, Map<String, Font>> FONT_CACHES =
            Collections.synchronizedMap(new WeakHashMap<>());

    private final List<Segment> segments = new ArrayList<>();

    /**
     * Appends plain (unstyled) text.
     *
     * @param text the text to append
     * @return this builder
     */
    public ExcelRichText text(String text) {
        segments.add(new Segment(text, null));
        return this;
    }

    /**
     * Appends bold text.
     *
     * @param text the text to append in bold
     * @return this builder
     */
    public ExcelRichText bold(String text) {
        return styled(text, s -> s.bold(true));
    }

    /**
     * Appends italic text.
     *
     * @param text the text to append in italic
     * @return this builder
     */
    public ExcelRichText italic(String text) {
        return styled(text, s -> s.italic(true));
    }

    /**
     * Appends text with fully customized font styling.
     *
     * @param text       the text to append
     * @param configurer a consumer that configures the {@link FontStyle}
     * @return this builder
     */
    public ExcelRichText styled(String text, Consumer<FontStyle> configurer) {
        FontStyle style = new FontStyle();
        configurer.accept(style);
        segments.add(new Segment(text, style));
        return this;
    }

    /**
     * Converts this rich text to a POI {@link RichTextString} for cell value assignment.
     *
     * @param wb        the workbook used to create fonts
     * @param fontCache a cache map to reuse fonts with identical styling
     * @return the rich text string ready to be set on a cell
     */
    RichTextString toRichTextString(SXSSFWorkbook wb, Map<String, Font> fontCache) {
        StringBuilder fullText = new StringBuilder();
        for (Segment seg : segments) {
            fullText.append(seg.text);
        }
        XSSFRichTextString rts = new XSSFRichTextString(fullText.toString());

        int pos = 0;
        for (Segment seg : segments) {
            int end = pos + seg.text.length();
            if (seg.style != null && !seg.text.isEmpty()) {
                Font font = seg.style.resolveFont(wb, fontCache);
                rts.applyFont(pos, end, font);
            }
            pos = end;
        }
        return rts;
    }

    /**
     * Returns the font cache associated with the given workbook.
     * <p>
     * The cache is stored in a {@link WeakHashMap} so that entries are automatically
     * removed when the workbook is garbage-collected.
     *
     * @param wb the workbook
     * @return the font cache map
     */
    static Map<String, Font> getFontCache(SXSSFWorkbook wb) {
        return FONT_CACHES.computeIfAbsent(wb, k -> new HashMap<>());
    }

    /**
     * Returns the plain text content (without formatting) for display or width calculation.
     */
    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        for (Segment seg : segments) {
            sb.append(seg.text);
        }
        return sb.toString();
    }

    private record Segment(String text, @Nullable FontStyle style) {
    }

    /**
     * Style configuration for a rich text segment.
     * <p>
     * Supports bold, italic, underline, strikethrough, font color (RGB), and font size.
     */
    public static class FontStyle {
        private boolean bold;
        private boolean italic;
        private boolean underline;
        private boolean strikethrough;
        private int @Nullable [] color;
        private @Nullable Integer fontSize;

        /**
         * Sets bold styling.
         */
        public FontStyle bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        /**
         * Sets italic styling.
         */
        public FontStyle italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        /**
         * Sets underline styling.
         */
        public FontStyle underline(boolean underline) {
            this.underline = underline;
            return this;
        }

        /**
         * Sets strikethrough styling.
         */
        public FontStyle strikethrough(boolean strikethrough) {
            this.strikethrough = strikethrough;
            return this;
        }

        /**
         * Sets the font color using RGB values.
         *
         * @param r red component (0–255)
         * @param g green component (0–255)
         * @param b blue component (0–255)
         */
        public FontStyle color(int r, int g, int b) {
            this.color = new int[]{r, g, b};
            return this;
        }

        /**
         * Sets the font color using a predefined {@link ExcelColor}.
         *
         * @param color the color preset
         */
        public FontStyle color(ExcelColor color) {
            return color(color.getR(), color.getG(), color.getB());
        }

        /**
         * Sets the font size in points.
         *
         * @param size font size in points
         */
        public FontStyle fontSize(int size) {
            this.fontSize = size;
            return this;
        }

        private String cacheKey() {
            return bold + ":" + italic + ":" + underline + ":" + strikethrough
                    + ":" + (color != null ? color[0] + "," + color[1] + "," + color[2] : "null")
                    + ":" + fontSize;
        }

        Font resolveFont(SXSSFWorkbook wb, Map<String, Font> fontCache) {
            String key = cacheKey();
            return fontCache.computeIfAbsent(key, k -> {
                Font font = wb.createFont();
                font.setBold(bold);
                font.setItalic(italic);
                if (underline) {
                    font.setUnderline(Font.U_SINGLE);
                }
                font.setStrikeout(strikethrough);
                if (color != null && font instanceof XSSFFont xf) {
                    xf.setColor(new XSSFColor(new byte[]{(byte) color[0], (byte) color[1], (byte) color[2]}));
                }
                if (fontSize != null) {
                    font.setFontHeightInPoints(fontSize.shortValue());
                }
                return font;
            });
        }
    }
}
