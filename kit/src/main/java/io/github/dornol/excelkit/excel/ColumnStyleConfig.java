package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.jspecify.annotations.Nullable;

import java.util.function.Function;

/**
 * Shared base for column styling configuration.
 * <p>
 * Used by both {@link ExcelColumn.ExcelColumnBuilder} and
 * {@link ExcelSheetWriter.ColumnConfig} to eliminate duplicated field
 * declarations and setter methods.
 *
 * <h3>Why not split into FontStyle, BorderStyle, LayoutConfig, etc.?</h3>
 * This class is intentionally kept flat. Although it has many fields, it is a pure
 * configuration holder with no logic — each field maps to one fluent setter.
 * Splitting would force nested builders on callers (e.g. {@code column.font(f -> f.bold(true))})
 * and complicate the 4 subclasses that inherit from this base, without reducing actual complexity.
 *
 * @param <T>    the row data type
 * @param <SELF> the concrete subclass type, for fluent method chaining
 * @author dhkim
 * @since 0.7.0
 */
@SuppressWarnings("unchecked")
public abstract class ColumnStyleConfig<T, SELF extends ColumnStyleConfig<T, SELF>> {

    @Nullable ExcelDataType dataType;
    @Nullable String dataFormat;
    HorizontalAlignment alignment = HorizontalAlignment.CENTER;
    boolean alignmentSet;
    int @Nullable [] backgroundColor;
    @Nullable Boolean bold;
    @Nullable Integer fontSize;
    int minWidth;
    int maxWidth;
    boolean fixedWidth;
    String @Nullable [] dropdownOptions;
    @Nullable CellColorFunction<T> cellColorFunction;
    @Nullable String groupName;
    int outlineLevel;
    @Nullable Function<T, @Nullable String> commentFunction;
    @Nullable ExcelBorderStyle borderStyle;
    @Nullable Boolean locked;
    boolean hidden;
    @Nullable ExcelValidation validation;
    @Nullable Short rotation;
    @Nullable ExcelBorderStyle borderTop;
    @Nullable ExcelBorderStyle borderBottom;
    @Nullable ExcelBorderStyle borderLeft;
    @Nullable ExcelBorderStyle borderRight;
    int @Nullable [] fontColor;
    @Nullable Boolean strikethrough;
    @Nullable Boolean underline;
    @Nullable VerticalAlignment verticalAlignment;
    @Nullable Boolean wrapText;
    @Nullable String fontName;
    @Nullable Short indentation;

    // Return self for fluent chaining
    private SELF self() {
        return (SELF) this;
    }

    /**
     * Sets the column's data type (used for styling and value conversion).
     */
    public SELF type(ExcelDataType dataType) {
        this.dataType = dataType;
        return self();
    }

    /**
     * Sets the column's Excel cell data format.
     */
    public SELF format(String dataFormat) {
        this.dataFormat = dataFormat;
        return self();
    }

    /**
     * Sets the column's horizontal text alignment.
     */
    public SELF alignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
        this.alignmentSet = true;
        return self();
    }

    /**
     * Sets the background color for this column's cells.
     *
     * @param r Red component (0-255)
     * @param g Green component (0-255)
     * @param b Blue component (0-255)
     */
    public SELF backgroundColor(int r, int g, int b) {
        this.backgroundColor = new int[]{r, g, b};
        return self();
    }

    /**
     * Sets the background color for this column's cells using a preset color.
     *
     * @param color Preset color
     */
    public SELF backgroundColor(ExcelColor color) {
        return backgroundColor(color.getR(), color.getG(), color.getB());
    }

    /**
     * Sets whether this column's font should be bold.
     */
    public SELF bold(boolean bold) {
        this.bold = bold;
        return self();
    }

    /**
     * Sets the font size for this column's cells.
     *
     * @param fontSize Font size in points (must be positive)
     */
    public SELF fontSize(int fontSize) {
        if (fontSize <= 0) {
            throw new IllegalArgumentException("fontSize must be positive");
        }
        this.fontSize = fontSize;
        return self();
    }

    /**
     * Sets a fixed column width. The column will not auto-resize.
     *
     * @param fixedWidth Fixed width value (in Excel internal units)
     */
    public SELF width(int fixedWidth) {
        this.fixedWidth = true;
        this.minWidth = fixedWidth;
        return self();
    }

    /**
     * Sets the minimum column width. Auto-resize will not shrink below this value.
     *
     * @param minWidth Minimum width value (in Excel internal units)
     */
    public SELF minWidth(int minWidth) {
        this.minWidth = minWidth;
        return self();
    }

    /**
     * Sets the maximum column width. Auto-resize will not grow beyond this value.
     *
     * @param maxWidth Maximum width value (in Excel internal units)
     */
    public SELF maxWidth(int maxWidth) {
        this.maxWidth = maxWidth;
        return self();
    }

    /**
     * Sets dropdown validation options for this column's cells.
     *
     * @param options The list of allowed values for the dropdown
     */
    public SELF dropdown(String... options) {
        this.dropdownOptions = options;
        return self();
    }

    /**
     * Sets a per-cell conditional color function.
     * <p>
     * The function receives the resolved cell value and the row data, and returns
     * an {@link ExcelColor} to apply as the cell background, or {@code null} for no override.
     * Cell-level color takes precedence over row-level {@code rowColor}.
     *
     * @param cellColorFunction function to determine per-cell background color
     */
    public SELF cellColor(CellColorFunction<T> cellColorFunction) {
        this.cellColorFunction = cellColorFunction;
        return self();
    }

    /**
     * Sets the group header name for this column.
     * <p>
     * Adjacent columns with the same group name will share a merged group header row
     * above the regular column header row.
     *
     * @param groupName the group header label
     */
    public SELF group(String groupName) {
        this.groupName = groupName;
        return self();
    }

    /**
     * Sets the outline (grouping) level for this column.
     * <p>
     * Columns with an outline level &gt; 0 can be collapsed/expanded in Excel.
     * Adjacent columns with the same outline level are grouped together.
     *
     * @param level the outline level (1-7, 0 = no outline)
     */
    public SELF outline(int level) {
        if (level < 0 || level > 7) {
            throw new IllegalArgumentException("outline level must be between 0 and 7");
        }
        this.outlineLevel = level;
        return self();
    }

    /**
     * Sets a function that generates a cell comment (note) for each row.
     * <p>
     * The function receives the row data and returns the comment text,
     * or {@code null} if no comment should be added.
     *
     * @param commentFunction function to generate comment text per row
     */
    public SELF comment(Function<T, @Nullable String> commentFunction) {
        this.commentFunction = commentFunction;
        return self();
    }

    /**
     * Sets the border style for this column's cells.
     * <p>
     * Overrides the default THIN border on all sides.
     *
     * @param borderStyle the border style to apply
     */
    public SELF border(ExcelBorderStyle borderStyle) {
        this.borderStyle = borderStyle;
        return self();
    }

    /**
     * Sets the top border style for this column's cells.
     *
     * @param borderStyle the border style to apply to the top border
     */
    public SELF borderTop(ExcelBorderStyle borderStyle) {
        this.borderTop = borderStyle;
        return self();
    }

    /**
     * Sets the bottom border style for this column's cells.
     *
     * @param borderStyle the border style to apply to the bottom border
     */
    public SELF borderBottom(ExcelBorderStyle borderStyle) {
        this.borderBottom = borderStyle;
        return self();
    }

    /**
     * Sets the left border style for this column's cells.
     *
     * @param borderStyle the border style to apply to the left border
     */
    public SELF borderLeft(ExcelBorderStyle borderStyle) {
        this.borderLeft = borderStyle;
        return self();
    }

    /**
     * Sets the right border style for this column's cells.
     *
     * @param borderStyle the border style to apply to the right border
     */
    public SELF borderRight(ExcelBorderStyle borderStyle) {
        this.borderRight = borderStyle;
        return self();
    }

    /**
     * Sets whether this column's cells should be locked when sheet protection is enabled.
     * <p>
     * By default, all cells are locked when sheet protection is active.
     * Set to {@code false} to allow editing of this column's cells even when the sheet is protected.
     *
     * @param locked whether cells should be locked
     */
    public SELF locked(boolean locked) {
        this.locked = locked;
        return self();
    }

    /**
     * Marks this column as hidden in the Excel output.
     */
    public SELF hidden() {
        this.hidden = true;
        return self();
    }

    /**
     * Sets whether this column should be hidden in the Excel output.
     *
     * @param hidden whether the column should be hidden
     */
    public SELF hidden(boolean hidden) {
        this.hidden = hidden;
        return self();
    }

    /**
     * Sets the text rotation angle for this column's cells.
     * <p>
     * Positive values rotate text counter-clockwise (0 to 90 degrees).
     * Negative values rotate text clockwise (-1 to -90 degrees).
     *
     * @param degrees rotation angle (-90 to 90)
     */
    public SELF rotation(int degrees) {
        if (degrees < -90 || degrees > 90) {
            throw new IllegalArgumentException("rotation must be between -90 and 90 degrees");
        }
        this.rotation = toExcelRotation(degrees);
        return self();
    }

    /**
     * Sets the font color for this column's cells using RGB values.
     *
     * @param r Red component (0-255)
     * @param g Green component (0-255)
     * @param b Blue component (0-255)
     */
    public SELF fontColor(int r, int g, int b) {
        this.fontColor = new int[]{r, g, b};
        return self();
    }

    /**
     * Sets the font color for this column's cells using a preset color.
     *
     * @param color Preset color
     */
    public SELF fontColor(ExcelColor color) {
        return fontColor(color.getR(), color.getG(), color.getB());
    }

    /**
     * Enables strikethrough on this column's font.
     */
    public SELF strikethrough() {
        this.strikethrough = true;
        return self();
    }

    /**
     * Sets whether this column's font should be strikethrough.
     *
     * @param strikethrough whether to apply strikethrough
     */
    public SELF strikethrough(boolean strikethrough) {
        this.strikethrough = strikethrough;
        return self();
    }

    /**
     * Enables underline on this column's font.
     */
    public SELF underline() {
        this.underline = true;
        return self();
    }

    /**
     * Sets whether this column's font should be underlined.
     *
     * @param underline whether to apply underline
     */
    public SELF underline(boolean underline) {
        this.underline = underline;
        return self();
    }

    /**
     * Sets the column's vertical text alignment.
     *
     * @param verticalAlignment vertical alignment (e.g., TOP, CENTER, BOTTOM, JUSTIFY)
     */
    public SELF verticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return self();
    }

    /**
     * Enables text wrapping for this column's cells.
     * <p>
     * When enabled, cell content that exceeds the column width wraps to the next line
     * within the cell instead of being clipped.
     */
    public SELF wrapText() {
        this.wrapText = true;
        return self();
    }

    /**
     * Sets whether text wrapping is enabled for this column's cells.
     *
     * @param wrapText whether to enable text wrapping
     */
    public SELF wrapText(boolean wrapText) {
        this.wrapText = wrapText;
        return self();
    }

    /**
     * Sets the font family name for this column's cells.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕", "Times New Roman")
     */
    public SELF fontName(String fontName) {
        this.fontName = fontName;
        return self();
    }

    /**
     * Sets the indentation level for this column's cells.
     * <p>
     * Indentation shifts the cell content to the right by the specified number of levels.
     *
     * @param level the indentation level (0-250)
     */
    public SELF indentation(int level) {
        if (level < 0 || level > 250) {
            throw new IllegalArgumentException("indentation level must be between 0 and 250");
        }
        this.indentation = (short) level;
        return self();
    }

    /**
     * Sets advanced data validation for this column.
     *
     * @param validation the validation configuration
     */
    public SELF validation(ExcelValidation validation) {
        this.validation = validation;
        return self();
    }

    /**
     * Applies defaults from the given config to this config.
     * Only null/default fields in this config are overridden.
     */
    void applyDefaults(ColumnStyleConfig<?, ?> defaults) {
        if (this.dataType == null) this.dataType = defaults.dataType;
        if (this.dataFormat == null) this.dataFormat = defaults.dataFormat;
        if (this.backgroundColor == null) this.backgroundColor = defaults.backgroundColor;
        if (this.bold == null) this.bold = defaults.bold;
        if (this.fontSize == null) this.fontSize = defaults.fontSize;
        if (this.borderStyle == null) this.borderStyle = defaults.borderStyle;
        if (this.locked == null) this.locked = defaults.locked;
        if (this.rotation == null) this.rotation = defaults.rotation;
        if (this.borderTop == null) this.borderTop = defaults.borderTop;
        if (this.borderBottom == null) this.borderBottom = defaults.borderBottom;
        if (this.borderLeft == null) this.borderLeft = defaults.borderLeft;
        if (this.borderRight == null) this.borderRight = defaults.borderRight;
        if (this.fontColor == null) this.fontColor = defaults.fontColor;
        if (this.strikethrough == null) this.strikethrough = defaults.strikethrough;
        if (this.underline == null) this.underline = defaults.underline;
        if (this.verticalAlignment == null) this.verticalAlignment = defaults.verticalAlignment;
        if (this.wrapText == null) this.wrapText = defaults.wrapText;
        if (this.fontName == null) this.fontName = defaults.fontName;
        if (this.indentation == null) this.indentation = defaults.indentation;
        if (!this.alignmentSet && defaults.alignmentSet) {
            this.alignment = defaults.alignment;
            this.alignmentSet = true;
        }
    }

    /**
     * Concrete subclass for defining default column styles at the writer level.
     *
     * @param <T> the row data type
     */
    public static class DefaultStyleConfig<T> extends ColumnStyleConfig<T, DefaultStyleConfig<T>> {
    }

    /**
     * Converts a user-facing rotation angle (-90 to 90) to the POI internal representation.
     * <p>
     * POI uses 0-90 for counter-clockwise and 91-180 for clockwise
     * (e.g., -1 degree maps to 91, -90 degrees maps to 180).
     *
     * @param degrees user-facing rotation angle
     * @return POI-internal rotation value
     */
    static short toExcelRotation(int degrees) {
        return (short) (degrees >= 0 ? degrees : 90 + Math.abs(degrees));
    }
}
