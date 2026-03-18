package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * Orchestrates multi-sheet Excel workbook creation where each sheet can have
 * a different data type.
 * <p>
 * Unlike {@link ExcelWriter} which handles automatic sheet rollover for a single data type,
 * {@code ExcelWorkbook} allows explicitly writing different data types to separate sheets.
 *
 * <pre>{@code
 * try (ExcelWorkbook workbook = new ExcelWorkbook(ExcelColor.STEEL_BLUE)) {
 *     workbook.<User>sheet("Users")
 *         .column("Name", u -> u.getName())
 *         .column("Status", u -> u.getStatus(), c -> c.dropdown("Active", "Inactive"))
 *         .write(userStream);
 *
 *     workbook.<Order>sheet("Orders")
 *         .column("ID", o -> o.getId())
 *         .column("Amount", o -> o.getAmount(), c -> c.type(ExcelDataType.DOUBLE))
 *         .write(orderStream);
 *
 *     ExcelHandler handler = workbook.finish();
 *     handler.consumeOutputStream(outputStream);
 * }
 * }</pre>
 *
 * @author dhkim
 */
public class ExcelWorkbook implements AutoCloseable {
    private static final Logger log = LoggerFactory.getLogger(ExcelWorkbook.class);

    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final SXSSFWorkbook wb;
    private CellStyle headerStyle;
    private final XSSFColor headerColor;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private final Set<String> usedSheetNames = new HashSet<>();
    private boolean finished = false;
    private @Nullable String workbookPassword;
    private @Nullable String headerFontName;
    private @Nullable Integer headerFontSize;

    /**
     * Creates a workbook with a default white header color.
     */
    public ExcelWorkbook() {
        this(255, 255, 255);
    }

    /**
     * Creates a workbook with a custom RGB header color.
     *
     * @param r Red component (0–255)
     * @param g Green component (0–255)
     * @param b Blue component (0–255)
     */
    public ExcelWorkbook(int r, int g, int b) {
        this(r, g, b, DEFAULT_ROW_ACCESS_WINDOW_SIZE);
    }

    /**
     * Creates a workbook with a custom RGB header color and row access window size.
     *
     * @param r                   Red component (0–255)
     * @param g                   Green component (0–255)
     * @param b                   Blue component (0–255)
     * @param rowAccessWindowSize Number of rows kept in memory by SXSSFWorkbook
     */
    public ExcelWorkbook(int r, int g, int b, int rowAccessWindowSize) {
        this.wb = new SXSSFWorkbook(rowAccessWindowSize);
        this.headerColor = new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b});
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor);
    }

    /**
     * Creates a workbook with a preset header color.
     *
     * @param color Preset header color
     */
    public ExcelWorkbook(ExcelColor color) {
        this(color.getR(), color.getG(), color.getB());
    }

    /**
     * Creates a workbook with a preset header color and row access window size.
     *
     * @param color               Preset header color
     * @param rowAccessWindowSize Number of rows kept in memory by SXSSFWorkbook
     */
    public ExcelWorkbook(ExcelColor color, int rowAccessWindowSize) {
        this(color.getR(), color.getG(), color.getB(), rowAccessWindowSize);
    }

    /**
     * Protects the workbook structure with the given password.
     * <p>
     * When enabled, users cannot add, delete, rename, or reorder sheets.
     *
     * @param password the protection password
     * @return this workbook for chaining
     */
    public ExcelWorkbook protectWorkbook(String password) {
        this.workbookPassword = password;
        return this;
    }

    /**
     * Sets the header font name for all sheets.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return this workbook for chaining
     */
    public ExcelWorkbook headerFontName(String fontName) {
        this.headerFontName = fontName;
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize);
        return this;
    }

    /**
     * Sets the header font size for all sheets.
     *
     * @param fontSize font size in points (must be positive)
     * @return this workbook for chaining
     */
    public ExcelWorkbook headerFontSize(int fontSize) {
        if (fontSize <= 0) {
            throw new IllegalArgumentException("fontSize must be positive");
        }
        this.headerFontSize = fontSize;
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize);
        return this;
    }

    /**
     * Creates a new sheet with the given name and returns a typed writer for it.
     *
     * @param name the sheet name (must be unique within this workbook)
     * @param <T>  the data type for this sheet's rows
     * @return an {@link ExcelSheetWriter} for configuring and writing the sheet
     * @throws ExcelWriteException if the workbook is already finished or the sheet name is duplicate
     */
    public <T> ExcelSheetWriter<T> sheet(String name) {
        if (finished) {
            throw new ExcelWriteException("Workbook is already finished");
        }
        if (!usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name: " + name);
        }
        return new ExcelSheetWriter<>(wb, wb.createSheet(name), name, headerStyle, cellStyleCache, usedSheetNames);
    }

    /**
     * Finishes the workbook and returns an {@link ExcelHandler} for output.
     * <p>
     * After calling this method, no more sheets can be added.
     *
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler finish() {
        finished = true;
        ExcelWriteSupport.applyWorkbookProtection(wb, workbookPassword);
        return new ExcelHandler(wb);
    }

    /**
     * Closes the underlying workbook if it has not been finished.
     * If {@link #finish()} was called, the workbook lifecycle is managed by {@link ExcelHandler}.
     */
    @Override
    public void close() {
        if (!finished) {
            try {
                wb.close();
            } catch (Exception e) {
                log.warn("Failed to close workbook", e);
            }
        }
    }
}
