package io.github.dornol.excelkit.excel;

/**
 * Represents a cell comment (note) to be added to an Excel cell.
 *
 * @param text   the comment text
 * @param author the comment author (nullable, defaults to empty)
 * @author dhkim
 * @since 0.6.0
 */
public record ExcelCellComment(String text, String author) {

    /**
     * Creates a comment with the given text and no author.
     *
     * @param text the comment text
     */
    public ExcelCellComment(String text) {
        this(text, null);
    }
}
