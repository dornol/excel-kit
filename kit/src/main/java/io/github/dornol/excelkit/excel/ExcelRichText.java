package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
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
import java.util.Objects;
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
 * @since 0.6.4
 */
public class ExcelRichText {

    /** Creates a new empty rich text instance. */
    public ExcelRichText() {}

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
        segments.add(new Segment(Objects.requireNonNull(text, "text cannot be null"), null));
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
        Objects.requireNonNull(text, "text cannot be null");
        Objects.requireNonNull(configurer, "configurer cannot be null");
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
    /**
     * The segments, with no run boundary inside a surrogate pair: a character outside the BMP (an emoji) whose
     * halves ended up in two segments is kept whole, in the segment its first half is in. Each run is stored as
     * text of its own, and half of a pair cannot be.
     */
    private List<Segment> wholeCharacterSegments() {
        List<Segment> result = new ArrayList<>(segments.size());
        for (Segment segment : segments) {
            String text = segment.text;
            if (text.isEmpty()) continue;
            if (!result.isEmpty() && Character.isLowSurrogate(text.charAt(0))) {
                Segment previous = result.get(result.size() - 1);
                if (Character.isHighSurrogate(previous.text.charAt(previous.text.length() - 1))) {
                    result.set(result.size() - 1, new Segment(previous.text + text.charAt(0), previous.style));
                    text = text.substring(1);
                    if (text.isEmpty()) continue;
                }
            }
            result.add(new Segment(text, segment.style));
        }
        return result;
    }

    XSSFRichTextString toRichTextString(SXSSFWorkbook wb, Map<String, Font> fontCache) {
        return toRichTextString(wb, fontCache, null);
    }

    XSSFRichTextString toRichTextString(SXSSFWorkbook wb, Map<String, Font> fontCache, @Nullable XSSFFont base) {
        List<Segment> segments = wholeCharacterSegments();
        // Text nothing is formatted in is a plain string: the cell's own font already applies to it.
        if (segments.stream().noneMatch(seg -> seg.style != null && !seg.text.isEmpty())) base = null;
        StringBuilder fullText = new StringBuilder();
        for (Segment seg : segments) {
            fullText.append(seg.text);
        }
        String logical = fullText.toString();
        String encoded = ExcelCellText.encode(logical);
        if (encoded != logical) {
            return encodedRichText(wb, fontCache, segments, logical, encoded, base);
        }
        XSSFRichTextString rts = new XSSFRichTextString(logical);

        int pos = 0;
        for (Segment seg : segments) {
            int end = pos + seg.text.length();
            if (!seg.text.isEmpty()) {
                Font font = fontOf(seg, wb, fontCache, base);
                if (font != null) rts.applyFont(pos, end, font);
            }
            pos = end;
        }
        return rts;
    }

    private XSSFRichTextString encodedRichText(SXSSFWorkbook wb, Map<String, Font> fontCache,
                                              List<Segment> segments, String logical, String encoded,
                                              @Nullable XSSFFont base) {
        var text = new ExcelCellText.StreamingText(logical, encoded);
        int position = 0;
        int storedPosition = 0;
        for (Segment segment : segments) {
            if (segment.text.isEmpty()) continue;
            int storedStart = storedPosition;
            int end = position + segment.text.length();
            while (position < end) {
                storedPosition += ExcelCellText.needsEscape(logical, position++) ? 7 : 1;
            }
            XSSFFont font = (XSSFFont) fontOf(segment, wb, fontCache, base);
            text.append(encoded.substring(storedStart, storedPosition), font);
        }
        return text;
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
    /** The font of a segment: its own style laid over the base font, or the base font itself when it has no style. */
    private static @Nullable Font fontOf(Segment segment, SXSSFWorkbook wb, Map<String, Font> fontCache, @Nullable XSSFFont base) {
        if (segment.style != null) return segment.style.resolveFont(wb, fontCache, base);
        return base;
    }

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
        /** Creates a new font style with defaults. */
        public FontStyle() {}

        /** The largest font size Excel accepts, in points. */
        private static final int MAX_FONT_SIZE = 409;

        private boolean bold;
        private boolean italic;
        private boolean underline;
        private boolean strikethrough;
        private int @Nullable [] color;
        private @Nullable Integer fontSize;

        /**
         * Sets bold styling.
         *
         * @param bold whether to apply bold
         * @return this instance for chaining
         */
        public FontStyle bold(boolean bold) {
            this.bold = bold;
            return this;
        }

        /**
         * Sets italic styling.
         *
         * @param italic whether to apply italic
         * @return this instance for chaining
         */
        public FontStyle italic(boolean italic) {
            this.italic = italic;
            return this;
        }

        /**
         * Sets underline styling.
         *
         * @param underline whether to apply underline
         * @return this instance for chaining
         */
        public FontStyle underline(boolean underline) {
            this.underline = underline;
            return this;
        }

        /**
         * Sets strikethrough styling.
         *
         * @param strikethrough whether to apply strikethrough
         * @return this instance for chaining
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
         * @return this instance for chaining
         */
        public FontStyle color(int r, int g, int b) {
            this.color = ExcelColor.of(r, g, b).toRgb();
            return this;
        }

        /**
         * Sets the font color using a predefined {@link ExcelColor}.
         *
         * @param color the color preset
         * @return this instance for chaining
         */
        public FontStyle color(ExcelColor color) {
            return color(color.getR(), color.getG(), color.getB());
        }

        /**
         * Sets the font size in points.
         *
         * @param size font size in points
         * @return this instance for chaining
         */
        public FontStyle fontSize(int size) {
            if (size < 1 || size > MAX_FONT_SIZE) {
                throw new IllegalArgumentException("fontSize must be between 1 and " + MAX_FONT_SIZE + " points, got " + size);
            }
            this.fontSize = size;
            return this;
        }

        private String cacheKey() {
            return bold + ":" + italic + ":" + underline + ":" + strikethrough
                    + ":" + (color != null ? color[0] + "," + color[1] + "," + color[2] : "null")
                    + ":" + fontSize;
        }

        Font resolveFont(SXSSFWorkbook wb, Map<String, Font> fontCache) {
            return resolveFont(wb, fontCache, null);
        }

        /** This style laid over {@code base}: what the style does not set (name, size, colour, ...) stays the base's. */
        Font resolveFont(SXSSFWorkbook wb, Map<String, Font> fontCache, @Nullable XSSFFont base) {
            String key = cacheKey() + (base == null ? "" : "|" + base.getIndex());
            return fontCache.computeIfAbsent(key, k -> {
                Font font = wb.createFont();
                if (base != null) {
                    font.setFontName(base.getFontName());
                    font.setFontHeightInPoints(base.getFontHeightInPoints());
                    font.setBold(base.getBold());
                    font.setItalic(base.getItalic());
                    font.setUnderline(base.getUnderline());
                    font.setStrikeout(base.getStrikeout());
                    font.setCharSet(base.getCharSet());
                    if (font instanceof XSSFFont xf && base.getXSSFColor() != null) xf.setColor(base.getXSSFColor());
                }
                if (bold) font.setBold(true);
                if (italic) font.setItalic(true);
                if (underline) {
                    font.setUnderline(Font.U_SINGLE);
                }
                if (strikethrough) font.setStrikeout(true);
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
