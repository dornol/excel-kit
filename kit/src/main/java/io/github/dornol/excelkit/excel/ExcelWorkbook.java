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
import java.util.function.Consumer;

/**
 * Orchestrates multi-sheet Excel workbook creation where each sheet can have
 * a different data type.
 * <p>
 * Unlike {@link ExcelWriter} which handles automatic sheet rollover for a single data type,
 * {@code ExcelWorkbook} allows explicitly writing different data types to separate sheets.
 *
 * <pre>{@code
 * try (ExcelWorkbook workbook = ExcelWorkbook.create().headerColor(ExcelColor.STEEL_BLUE)) {
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
 *     handler.write(outputStream);
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
    private XSSFColor headerColor;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private final Set<String> usedSheetNames = new HashSet<>();
    private boolean finished = false;
    private char @Nullable [] password;
    private @Nullable String workbookPassword;
    private @Nullable String headerFontName;
    private @Nullable Integer headerFontSize;
    private @Nullable HeaderStyleConfig headerStyleConfig;

    /**
     * Creates a new ExcelWorkbook with default initialization (white header, 1000 row window).
     *
     * @return a new workbook instance (implements AutoCloseable)
     * @since 0.17.0
     */
    public static ExcelWorkbook create() {
        return create(opts -> {});
    }

    /**
     * Creates a new ExcelWorkbook with initialization options.
     *
     * <pre>{@code
     * try (ExcelWorkbook wb = ExcelWorkbook.create(opts -> opts.rowAccessWindowSize(500))
     *         .headerColor(ExcelColor.STEEL_BLUE)) {
     *     wb.<User>sheet("Users").column("Name", User::getName).write(stream);
     *     wb.finish().write(out);
     * }
     * }</pre>
     *
     * @param configurer consumer that configures {@link InitOptions}
     * @return a new workbook instance (implements AutoCloseable)
     * @since 0.17.0
     */
    public static ExcelWorkbook create(Consumer<InitOptions> configurer) {
        InitOptions opts = new InitOptions();
        configurer.accept(opts);
        return new ExcelWorkbook(opts);
    }

    private ExcelWorkbook(InitOptions opts) {
        this.wb = new SXSSFWorkbook(opts.rowAccessWindowSize);
        ExcelColor defaultColor = ExcelColor.WHITE;
        this.headerColor = new XSSFColor(new byte[]{
                (byte) defaultColor.getR(),
                (byte) defaultColor.getG(),
                (byte) defaultColor.getB()
        });
        this.headerStyle = rebuildHeaderStyle();
    }

    private CellStyle rebuildHeaderStyle() {
        return ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize, headerStyleConfig);
    }

    /**
     * Initialization options for {@link ExcelWorkbook}. Passed to the configurer given to
     * {@link ExcelWorkbook#create(Consumer)}.
     * <p>
     * Restricted to settings that cannot be changed after the underlying SXSSFWorkbook
     * is constructed (currently just {@code rowAccessWindowSize}).
     *
     * @since 0.17.0
     */
    public static final class InitOptions {
        private int rowAccessWindowSize = DEFAULT_ROW_ACCESS_WINDOW_SIZE;

        private InitOptions() {}

        /**
         * Sets the SXSSF row access window size. Must be set at construction time because
         * POI's SXSSFWorkbook takes it as a constructor argument.
         *
         * @param size row window (must be positive)
         * @return this options object for chaining
         */
        public InitOptions rowAccessWindowSize(int size) {
            if (size <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
            this.rowAccessWindowSize = size;
            return this;
        }
    }

    /**
     * Sets a document property on the generated Excel file.
     * <p>
     * Standard keys ({@code title}, {@code subject}, {@code author}, {@code keywords},
     * {@code description}, {@code category}) are mapped to Excel core properties.
     * Other keys are stored as custom properties.
     *
     * @param key   the property name (case-insensitive for standard keys)
     * @param value the property value
     * @return this workbook for chaining
     * @since 0.17.0
     */
    public ExcelWorkbook documentProperty(String key, String value) {
        ExcelWriteSupport.applyDocumentProperty(wb, key, value);
        return this;
    }

    /**
     * Sets the header background color for all sheets. Must be called before any
     * {@link #sheet(String)} that relies on the header style.
     *
     * @param color header color (must not be null)
     * @return this workbook for chaining
     * @since 0.17.0
     */
    public ExcelWorkbook headerColor(ExcelColor color) {
        if (color == null) throw new IllegalArgumentException("color must not be null");
        this.headerColor = new XSSFColor(new byte[]{
                (byte) color.getR(),
                (byte) color.getG(),
                (byte) color.getB()
        });
        this.headerStyle = rebuildHeaderStyle();
        return this;
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
     * Sets the file encryption password.
     * <p>
     * When set, the resulting Excel file will be encrypted using the "agile" encryption mode,
     * and {@link ExcelHandler#writeTo(java.io.OutputStream)} will automatically
     * apply encryption — no need to pass the password to {@link ExcelHandler#writeTo(java.io.OutputStream, String)}.
     *
     * @param password the encryption password (must not be null or blank)
     * @return this workbook for chaining
     */
    public ExcelWorkbook password(String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.password = password.toCharArray();
        return this;
    }

    /**
     * Sets the file encryption password using a char array.
     * <p>
     * The array is copied internally; the caller may zero the original after this call.
     *
     * @param password the encryption password (copied internally; must not be null or empty)
     * @return this workbook for chaining
     * @since 0.17.0
     */
    public ExcelWorkbook password(char[] password) {
        if (password == null || password.length == 0 || isBlank(password)) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.password = password.clone();
        return this;
    }

    private static boolean isBlank(char[] chars) {
        for (char c : chars) {
            if (!Character.isWhitespace(c)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Sets the header font name for all sheets.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return this workbook for chaining
     */
    public ExcelWorkbook headerFontName(String fontName) {
        this.headerFontName = fontName;
        this.headerStyle = rebuildHeaderStyle();
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
        this.headerStyle = rebuildHeaderStyle();
        return this;
    }

    /**
     * Configures the global header cell style (alignment, bold, border, wrap text).
     *
     * @param configurer consumer to configure header style properties
     * @return this workbook for chaining
     * @since 0.17.0
     */
    public ExcelWorkbook headerStyle(java.util.function.Consumer<HeaderStyleConfig> configurer) {
        HeaderStyleConfig config = new HeaderStyleConfig();
        configurer.accept(config);
        this.headerStyleConfig = config;
        this.headerStyle = rebuildHeaderStyle();
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
            throw new ExcelWriteException("Duplicate sheet name: '" + name + "'");
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
        if (finished) {
            throw new ExcelWriteException("Workbook is already finished");
        }
        finished = true;
        ExcelWriteSupport.applyWorkbookProtection(wb, workbookPassword);
        return new ExcelHandler(wb, password);
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
