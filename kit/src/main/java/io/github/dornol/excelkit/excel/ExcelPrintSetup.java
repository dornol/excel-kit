package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
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
 *
 * @author dhkim
 */
public class ExcelPrintSetup {

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
        LETTER(PrintSetup.LETTER_PAPERSIZE),
        LEGAL(PrintSetup.LEGAL_PAPERSIZE),
        A3(PrintSetup.A3_PAPERSIZE),
        A4(PrintSetup.A4_PAPERSIZE),
        A5(PrintSetup.A5_PAPERSIZE),
        B4(12),
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
    private boolean fitToPage;
    private int fitWidth = 1;
    private int fitHeight;

    /**
     * Sets the page orientation.
     *
     * @param orientation the page orientation
     * @return this instance for chaining
     */
    public ExcelPrintSetup orientation(Orientation orientation) {
        this.orientation = orientation;
        return this;
    }

    /**
     * Sets the paper size.
     *
     * @param paperSize the paper size
     * @return this instance for chaining
     */
    public ExcelPrintSetup paperSize(PaperSize paperSize) {
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
        this.leftMargin = left;
        this.rightMargin = right;
        this.topMargin = top;
        this.bottomMargin = bottom;
        return this;
    }

    /**
     * Sets the left margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup leftMargin(double inches) {
        this.leftMargin = inches;
        return this;
    }

    /**
     * Sets the right margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup rightMargin(double inches) {
        this.rightMargin = inches;
        return this;
    }

    /**
     * Sets the top margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup topMargin(double inches) {
        this.topMargin = inches;
        return this;
    }

    /**
     * Sets the bottom margin.
     *
     * @param inches margin in inches
     * @return this instance for chaining
     */
    public ExcelPrintSetup bottomMargin(double inches) {
        this.bottomMargin = inches;
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
    public ExcelPrintSetup headerLeft(String text) {
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
    public ExcelPrintSetup headerCenter(String text) {
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
    public ExcelPrintSetup headerRight(String text) {
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
    public ExcelPrintSetup footerLeft(String text) {
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
    public ExcelPrintSetup footerCenter(String text) {
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
    public ExcelPrintSetup footerRight(String text) {
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
    void apply(SXSSFSheet sheet, int headerRowIndex) {
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

        if (leftMargin != null) sheet.setMargin(Sheet.LeftMargin, leftMargin);
        if (rightMargin != null) sheet.setMargin(Sheet.RightMargin, rightMargin);
        if (topMargin != null) sheet.setMargin(Sheet.TopMargin, topMargin);
        if (bottomMargin != null) sheet.setMargin(Sheet.BottomMargin, bottomMargin);

        if (headerLeft != null) sheet.getHeader().setLeft(headerLeft);
        if (headerCenter != null) sheet.getHeader().setCenter(headerCenter);
        if (headerRight != null) sheet.getHeader().setRight(headerRight);
        if (footerLeft != null) sheet.getFooter().setLeft(footerLeft);
        if (footerCenter != null) sheet.getFooter().setCenter(footerCenter);
        if (footerRight != null) sheet.getFooter().setRight(footerRight);

        if (repeatHeader) {
            sheet.setRepeatingRows(new CellRangeAddress(0, headerRowIndex, -1, -1));
        } else if (repeatRowsStart >= 0) {
            sheet.setRepeatingRows(new CellRangeAddress(repeatRowsStart, repeatRowsEnd, -1, -1));
        }
    }
}
