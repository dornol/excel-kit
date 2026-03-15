package io.github.dornol.excelkit.excel;

/**
 * Represents a hyperlink value for Excel cells.
 * <p>
 * Use with {@link ExcelDataType#HYPERLINK} to create clickable links in Excel.
 * If only a URL is provided, it is used as both the display label and the link target.
 *
 * <pre>{@code
 * writer.column("Website", user -> new ExcelHyperlink(user.getUrl(), "Visit"))
 *       .type(ExcelDataType.HYPERLINK)
 *       .write(stream);
 * }</pre>
 *
 * @param url   the hyperlink URL
 * @param label the display text shown in the cell
 * @author dhkim
 */
public record ExcelHyperlink(String url, String label) {

    /**
     * Creates a hyperlink where the URL is also used as the display label.
     *
     * @param url the hyperlink URL
     */
    public ExcelHyperlink(String url) {
        this(url, url);
    }
}
