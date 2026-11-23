package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.jspecify.annotations.Nullable;

/**
 * Configures page layout settings for printing Excel sheets.
 * <p>
 * Supports orientation, paper size, margins, headers/footers, repeat rows,
 * and fit-to-page scaling. Use with {@link ExcelWriter#printSetup(java.util.function.Consumer)}
 * or {@link ExcelSheetWriter#printSetup(java.util.function.Consumer)}.
 *
 * <h2>Header/Footer special codes</h2>
 * <ul>
 *   <li>{@code &P} — current page number</li>
 *   <li>{@code &N} — total number of pages</li>
 *   <li>{@code &D} — current date</li>
 *   <li>{@code &T} — current time</li>
 *   <li>{@code &F} — file name</li>
 * </ul>
 * An ampersand is the start of such a code, so a literal one is written twice: {@code "R&&D"} prints
 * {@code R&D}, where {@code "R&D"} prints {@code R} and the date. Excel allows 255 characters in a header or
 * footer, counting the {@code &L}, {@code &C} and {@code &R} that mark its sections.
 *
 * @author dhkim
 * @since 0.6.4
 */
public class ExcelPrintSetup {

    /** Creates a new print setup configuration with defaults. */
    public ExcelPrintSetup() {}

    /**
     * Page orientation for printing.
     */
    public enum Orientation {
        /** Portrait orientation (taller than wide). */
        PORTRAIT,
        /** Landscape orientation (wider than tall). */
        LANDSCAPE
    }

    /**
     * Standard paper sizes for printing.
     */
    public enum PaperSize {
        /** US Letter size. */
        LETTER(PrintSetup.LETTER_PAPERSIZE),
        /** US Legal size. */
        LEGAL(PrintSetup.LEGAL_PAPERSIZE),
        /** A3 paper size. */
        A3(PrintSetup.A3_PAPERSIZE),
        /** A4 paper size. */
        A4(PrintSetup.A4_PAPERSIZE),
        /** A5 paper size. */
        A5(PrintSetup.A5_PAPERSIZE),
        /** B4 paper size. */
        B4(12),
        /** B5 paper size. */
        B5(PrintSetup.B5_PAPERSIZE);

        private final short code;

        PaperSize(short code) {
            this.code = code;
        }

        PaperSize(int code) {
            this.code = (short) code;
        }

        short getCode() {
            return code;
        }
    }

    /** The most pages Excel fits a sheet to, in either direction. */
    private static final int MAX_FIT_PAGES = 32_767;
    /** What Excel keeps of a header or a footer. */
    private static final int MAX_HEADER_FOOTER = 255;

    private static double requireMargin(double inches, String side) {
        if (!Double.isFinite(inches) || inches < 0) {
            throw new IllegalArgumentException("The " + side + " margin must be a number of inches that is not "
                    + "negative, got " + inches);
        }
        return inches;
    }

    /** Fails if the header or the footer is longer than Excel keeps; called when the setup is handed to a writer. */
    void validate() {
        checkLength("header", headerLeft, headerCenter, headerRight);
        checkLength("footer", footerLeft, footerCenter, footerRight);
    }

    private static void checkLength(String what, @Nullable String left, @Nullable String center, @Nullable String right) {
        int length = (left == null ? 0 : 2 + left.length()) + (center == null ? 0 : 2 + center.length())
                + (right == null ? 0 : 2 + right.length());
        if (length > MAX_HEADER_FOOTER) {
            throw new IllegalArgumentException("The " + what + " is " + length + " characters (the &L, &C and &R "
                    + "that mark its sections count); Excel allows at most " + MAX_HEADER_FOOTER);
        }
    }

    private @Nullable Orientation orientation;
    private @Nullable PaperSize paperSize;
    private @Nullable Double leftMargin;
    private @Nullable Double rightMargin;
    private @Nullable Double topMargin;
    private @Nullable Double bottomMargin;
    private @Nullable String headerLeft;
    private @Nullable String headerCenter;
    private @Nullable String headerRight;
    private @Nullable String footerLeft;
    private @Nullable String footerCenter;
    private @Nullable String footerRight;
    private boolean repeatHeader;
    private int repeatRowsStart = -1;
    private int repeatRowsEnd = -1;
    private int rowBreakEvery;
    private boolean fitToPage;
    private int fitWidth = 1;
    private int fitHeight;

    /**
     * Sets the page orientation.
     *
     * @param orientation the page orientation
     * @return this instance for chaining
     */
    public ExcelPrintSetup orientation(@Nullable Orientation orientation) {
        this.orientation = orientation;
        return this;
    }

    /**
     * Sets the paper size.
     *
     * @param paperSize the paper size
     * @return this instance for chaining
     */
    public ExcelPrintSetup paperSize(@Nullable PaperSize paperSize) {
        this.paperSize = paperSize;
        return this;
    }

    /**
     * Sets all four page margins at once.
     *
     * @param left   left margin in inches
     * @param right  right margin in inches
     * @param top    top margin in inches
     * @param bottom bottom margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup margins(double left, double right, double top, double bottom) {
        this.leftMargin = requireMargin(left, "left");
        this.rightMargin = requireMargin(right, "right");
        this.topMargin = requireMargin(top, "top");
        this.bottomMargin = requireMargin(bottom, "bottom");
        return this;
    }

    /**
     * Sets the left margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup leftMargin(double inches) {
        this.leftMargin = requireMargin(inches, "left");
        return this;
    }

    /**
     * Sets the right margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup rightMargin(double inches) {
        this.rightMargin = requireMargin(inches, "right");
        return this;
    }

    /**
     * Sets the top margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup topMargin(double inches) {
        this.topMargin = requireMargin(inches, "top");
        return this;
    }

    /**
     * Sets the bottom margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup bottomMargin(double inches) {
        this.bottomMargin = requireMargin(inches, "bottom");
        return this;
    }

    /**
     * Sets the left section of the page header.
     * <p>
     * Supports special codes: {@code &P} (page number), {@code &N} (total pages),
     * {@code &D} (date), {@code &T} (time), {@code &F} (filename).
     *
     * @param text header text
     * @return this instance for chaining
     */
    public ExcelPrintSetup headerLeft(@Nullable String text) {
        this.headerLeft = text;
        return this;
    }

    /**
     * Sets the center section of the page header.
     *
     * @param text header text
     * @return this instance for chaining
     * @see #headerLeft(String) for special codes
     */
    public ExcelPrintSetup headerCenter(@Nullable String text) {
        this.headerCenter = text;
        return this;
    }

    /**
     * Sets the right section of the page header.
     *
     * @param text header text
     * @return this instance for chaining
     * @see #headerLeft(String) for special codes
     */
    public ExcelPrintSetup headerRight(@Nullable String text) {
        this.headerRight = text;
        return this;
    }

    /**
     * Sets the left section of the page footer.
     *
     * @param text footer text
     * @return this instance for chaining
     * @see #headerLeft(String) for special codes
     */
    public ExcelPrintSetup footerLeft(@Nullable String text) {
        this.footerLeft = text;
        return this;
    }

    /**
     * Sets the center section of the page footer.
     *
     * @param text footer text
     * @return this instance for chaining
     * @see #headerLeft(String) for special codes
     */
    public ExcelPrintSetup footerCenter(@Nullable String text) {
        this.footerCenter = text;
        return this;
    }

    /**
     * Sets the right section of the page footer.
     *
     * @param text footer text
     * @return this instance for chaining
     * @see #headerLeft(String) for special codes
     */
    public ExcelPrintSetup footerRight(@Nullable String text) {
        this.footerRight = text;
        return this;
    }

    /**
     * Repeats header rows (from row 0 through the column header row) on every printed page.
     * <p>
     * This includes any rows added by {@code beforeHeader} and group header rows.
     *
     * @return this instance for chaining
     */
    public ExcelPrintSetup repeatHeaderRows() {
        this.repeatHeader = true;
        this.repeatRowsStart = -1;
        this.repeatRowsEnd = -1;
        return this;
    }

    /**
     * Starts a new printed page after every {@code dataRows} data rows (the header rows are not counted), so that
     * each page holds that many rows whatever the paper size would have fitted. Each sheet of a rollover is
     * counted on its own. For a page per group of rows use {@code pageBreakOnChange(...)} on the writer.
     *
     * @param dataRows the rows per page, at least 1
     * @return this instance for chaining
     * @throws IllegalArgumentException if {@code dataRows} is below 1
     * @since 0.29.0
     */
    public ExcelPrintSetup rowBreakEvery(int dataRows) {
        if (dataRows < 1) {
            throw new IllegalArgumentException("rowBreakEvery needs at least 1 row per page, got " + dataRows);
        }
        this.rowBreakEvery = dataRows;
        return this;
    }

    /**
     * Repeats specific rows on every printed page.
     *
     * @param firstRow first row to repeat (0-based)
     * @param lastRow  last row to repeat (0-based, inclusive)
     * @return this instance for chaining
     */
    public ExcelPrintSetup repeatRows(int firstRow, int lastRow) {
        if (firstRow < 0 || lastRow < firstRow || lastRow > ExcelLimits.EXCEL_MAX_ROWS) {
            throw new IllegalArgumentException("repeatRows needs 0 <= firstRow <= lastRow <= "
                    + ExcelLimits.EXCEL_MAX_ROWS + ", got " + firstRow + " and " + lastRow);
        }
        this.repeatHeader = false;
        this.repeatRowsStart = firstRow;
        this.repeatRowsEnd = lastRow;
        return this;
    }

    /**
     * Enables fit-to-page scaling with the specified width and height in pages.
     * <p>
     * Use {@code height = 0} for automatic height scaling.
     *
     * @param width  number of pages wide
     * @param height number of pages tall (0 for auto)
     * @return this instance for chaining
     */
    public ExcelPrintSetup fitToPage(int width, int height) {
        if (width < 0 || width > MAX_FIT_PAGES || height < 0 || height > MAX_FIT_PAGES) {
            throw new IllegalArgumentException("fitToPage needs pages wide and tall between 0 (as many as needed) and "
                    + MAX_FIT_PAGES + ", got " + width + " and " + height);
        }
        this.fitToPage = true;
        this.fitWidth = width;
        this.fitHeight = height;
        return this;
    }

    /**
     * Fits the sheet to one page wide with automatic height.
     *
     * @return this instance for chaining
     */
    public ExcelPrintSetup fitToPageWidth() {
        return fitToPage(1, 0);
    }

    /**
     * Applies this print setup configuration to the given sheet.
     *
     * @param sheet          the sheet to configure
     * @param headerRowIndex the 0-based row index of the header row
     */
    void apply(SXSSFSheet sheet, int headerRowIndex, int dataEndRow) {
        PrintSetup ps = sheet.getPrintSetup();

        if (orientation != null) {
            ps.setLandscape(orientation == Orientation.LANDSCAPE);
        }
        if (paperSize != null) {
            ps.setPaperSize(paperSize.getCode());
        }
        if (fitToPage) {
            sheet.setFitToPage(true);
            ps.setFitWidth((short) fitWidth);
            ps.setFitHeight((short) fitHeight);
        }

        if (leftMargin != null) sheet.setMargin(PageMargin.LEFT, leftMargin);
        if (rightMargin != null) sheet.setMargin(PageMargin.RIGHT, rightMargin);
        if (topMargin != null) sheet.setMargin(PageMargin.TOP, topMargin);
        if (bottomMargin != null) sheet.setMargin(PageMargin.BOTTOM, bottomMargin);

        if (headerLeft != null) sheet.getHeader().setLeft(headerLeft);
        if (headerCenter != null) sheet.getHeader().setCenter(headerCenter);
        if (headerRight != null) sheet.getHeader().setRight(headerRight);
        if (footerLeft != null) sheet.getFooter().setLeft(footerLeft);
        if (footerCenter != null) sheet.getFooter().setCenter(footerCenter);
        if (footerRight != null) sheet.getFooter().setRight(footerRight);

        for (int row = headerRowIndex + 1 + rowBreakEvery; rowBreakEvery > 0 && row <= dataEndRow; row += rowBreakEvery) {
            sheet.setRowBreak(row - 1); // POI breaks below the row it is given; 'row' is the first row of the new page
        }

        if (repeatHeader) {
            sheet.setRepeatingRows(new CellRangeAddress(0, headerRowIndex, -1, -1));
        } else if (repeatRowsStart >= 0) {
            sheet.setRepeatingRows(new CellRangeAddress(repeatRowsStart, repeatRowsEnd, -1, -1));
        }
    }
}
