package io.github.dornol.excelkit.excel;

import org.jspecify.annotations.Nullable;

/**
 * Represents a cell comment (note) to be added to an Excel cell.
 * <p>
 * Use the static factory {@link #of(String)} and the wither-style methods
 * {@link #author(String)} / {@link #size(int, int)} to build comments fluently.
 *
 * <pre>{@code
 * ExcelCellComment.of("Enter in YYYY-MM-DD").size(3, 5)
 * ExcelCellComment.of("Reviewed").author("System").size(2, 2)
 * }</pre>
 *
 * @param text   the comment text (must be non-null)
 * @param author optional comment author (POI stores it in the XML but
 *               Excel does not surface it in the note tooltip; kept for
 *               completeness)
 * @param width  comment box width in cells; {@code 0} means default (2)
 * @param height comment box height in rows; {@code 0} means default (3)
 * @author dhkim
 * @since 0.6.0
 */
public record ExcelCellComment(String text, @Nullable String author, int width, int height) {

    /**
     * Validates invariants.
     */
    public ExcelCellComment {
        if (text == null) {
            throw new IllegalArgumentException("text must not be null");
        }
        if (width < 0 || height < 0) {
            throw new IllegalArgumentException("width/height must be >= 0");
        }
    }

    /**
     * Creates a comment with the given text, no author, and default size.
     *
     * @param text the comment text
     * @return a new {@code ExcelCellComment}
     */
    public static ExcelCellComment of(String text) {
        return new ExcelCellComment(text, null, 0, 0);
    }

    /**
     * Returns a copy with the given author.
     *
     * @param author the author name
     * @return a new {@code ExcelCellComment}
     */
    public ExcelCellComment author(@Nullable String author) {
        return new ExcelCellComment(text, author, width, height);
    }

    /**
     * Returns a copy with the given box size (in cells × rows).
     * <p>
     * Pass {@code 0} for either dimension to use the default (2 cols × 3 rows).
     *
     * @param width  width in cells
     * @param height height in rows
     * @return a new {@code ExcelCellComment}
     */
    public ExcelCellComment size(int width, int height) {
        return new ExcelCellComment(text, author, width, height);
    }
}
