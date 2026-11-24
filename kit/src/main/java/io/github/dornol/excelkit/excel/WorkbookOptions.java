package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;

import java.util.function.Consumer;

/**
 * The settings of the workbook that {@link ExcelWriter} and {@link ExcelWorkbook} both have: document properties, the look
 * of the header, the file password and the protection of the workbook structure. What each setting means is the same on
 * both; only the moment it must be set by differs, which is said where it matters.
 * <p>
 * The type is public so that a setting is written once for both, for example the look of the reports of a team:
 *
 * <pre>{@code
 * static <W extends WorkbookOptions<W>> W teamLook(W writer) {
 *     return writer.headerColor(ExcelColor.STEEL_BLUE).headerFontName("Arial").documentProperty("author", "Reporting");
 * }
 *
 * teamLook(ExcelWriter.<Order>create()).column("Id", Order::id);
 * teamLook(ExcelWorkbook.create()).sheet("Orders");
 * }</pre>
 *
 * It cannot be implemented outside the library.
 *
 * @param <SELF> the writer type, for chaining
 * @since 0.29.0
 */
public interface WorkbookOptions<SELF extends WorkbookOptions<SELF>> {

    /**
     * Sets a document property on the generated Excel file.
     * <p>
     * Standard keys ({@code title}, {@code subject}, {@code author}, {@code keywords},
     * {@code description}, {@code category}) are mapped to Excel core properties.
     * Other keys are stored as custom properties.
     *
     * @param key   the property name (case-insensitive for standard keys)
     * @param value the property value
     * @return this for chaining
     * @since 0.17.0
     */
    SELF documentProperty(String key, String value);

    /**
     * Sets the header background color of the sheets. Use presets like {@link ExcelColor#STEEL_BLUE} or custom via
     * {@link ExcelColor#of(int, int, int)}. Defaults to {@link ExcelColor#WHITE}.
     * <p>
     * A setting of the header look must come before what writes the header: {@code write(...)} of an
     * {@link ExcelWriter}, the {@link ExcelWorkbook#sheet(String)} that should use it for a workbook.
     *
     * @param color header color (must not be null)
     * @return this for chaining
     * @since 0.17.0
     */
    SELF headerColor(ExcelColor color);

    /**
     * Sets the header font name of the sheets (see {@link #headerColor(ExcelColor)} for when it must be set).
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return this for chaining
     */
    SELF headerFontName(String fontName);

    /**
     * Sets the header font size of the sheets in points (see {@link #headerColor(ExcelColor)} for when it must be set).
     *
     * @param fontSize font size in points (must be positive)
     * @return this for chaining
     */
    SELF headerFontSize(int fontSize);

    /**
     * Configures the global header cell style (alignment, bold, border, wrap text).
     * <p>
     * Font name, font size, and background color are set separately via
     * {@link #headerFontName(String)}, {@link #headerFontSize(int)}, and {@link #headerColor(ExcelColor)}.
     *
     * <pre>{@code
     * writer.headerStyle(h -> h
     *     .bold(true)
     *     .alignment(HorizontalAlignment.LEFT)
     *     .border(ExcelBorderStyle.MEDIUM)
     *     .wrapText(true));
     * }</pre>
     *
     * @param configurer consumer to configure header style properties
     * @return this for chaining
     * @since 0.17.0
     */
    SELF headerStyle(Consumer<HeaderStyleConfig> configurer);

    /**
     * Protects the workbook structure with the given password.
     * <p>
     * When enabled, users cannot add, delete, rename, or reorder sheets.
     *
     * @param password the protection password
     * @return this for chaining
     */
    SELF protectWorkbook(String password);

    /**
     * Sets the file encryption password.
     * <p>
     * When set, the resulting Excel file will be encrypted using the "agile" encryption mode,
     * and {@link ExcelHandler#writeTo(java.io.OutputStream)} will automatically
     * apply encryption — no need to pass the password to {@link ExcelHandler#writeTo(java.io.OutputStream, String)}.
     *
     * @param password the encryption password (must not be null or blank)
     * @return this for chaining
     */
    SELF password(String password);
}
