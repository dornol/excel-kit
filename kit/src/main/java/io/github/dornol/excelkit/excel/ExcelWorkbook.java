package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
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
 * @since 0.3.0
 */
public class ExcelWorkbook implements AutoCloseable {
    private static final Logger log = LoggerFactory.getLogger(ExcelWorkbook.class);

    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final SXSSFWorkbook wb;
    private final WorkbookSettings settings;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    // Styles derived from a column style (headers, row colors, stripes, subtotal rows) are keyed by what they are
    // made of, so the sheets of a workbook share them instead of adding the same styles once per sheet.
    private final Map<String, CellStyle> headerStyleCache = new HashMap<>();
    private final Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache = new HashMap<>();
    private final Set<String> usedSheetNames = new HashSet<>();
    private boolean finished = false;

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
        StreamingOptions streaming = opts.state.options();
        this.wb = ExcelWorkbookSupport.streamingWorkbook(streaming);
        this.settings = new WorkbookSettings(wb);
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
        private final StreamingInitOptionsState state = new StreamingInitOptionsState();

        private InitOptions() { }

        public InitOptions rowAccessWindowSize(int size) {
            state.rowAccessWindowSize(size);
            return this;
        }

        public InitOptions compressTempFiles(boolean enabled) {
            state.compressTempFiles(enabled);
            return this;
        }

        public InitOptions useSharedStrings(boolean enabled) {
            state.useSharedStrings(enabled);
            return this;
        }

        public InitOptions streaming(StreamingOptions options) {
            state.streaming(options);
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
        settings.documentProperty(key, value);
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
        settings.headerColor(color);
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
        settings.protectWorkbook(password);
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
        settings.password(password);
        return this;
    }

    /**
     * Sets the header font name for all sheets.
     *
     * @param fontName the font name (e.g., "Arial", "맑은 고딕")
     * @return this workbook for chaining
     */
    public ExcelWorkbook headerFontName(String fontName) {
        settings.headerFontName(fontName);
        return this;
    }

    /**
     * Sets the header font size for all sheets.
     *
     * @param fontSize font size in points (must be positive)
     * @return this workbook for chaining
     */
    public ExcelWorkbook headerFontSize(int fontSize) {
        settings.headerFontSize(fontSize);
        return this;
    }

    /**
     * Configures the global header cell style (alignment, bold, border, wrap text).
     *
     * @param configurer consumer to configure header style properties
     * @return this workbook for chaining
     * @since 0.17.0
     */
    public ExcelWorkbook headerStyle(Consumer<HeaderStyleConfig> configurer) {
        settings.headerStyle(configurer);
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
        return new ExcelSheetWriter<>(wb, wb.createSheet(name), name, settings.headerStyle(), cellStyleCache, headerStyleCache, rowStyleCache, usedSheetNames);
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
        ExcelWriteSupport.requireVisibleSheet(wb);
        finished = true;
        ExcelWorkbookSupport.applyProtection(wb, settings.workbookPassword());
        return new ExcelHandler(wb, settings.password());
    }

    /**
     * Closes the underlying workbook if it has not been finished.
     * If {@link #finish()} was called, the workbook lifecycle is managed by {@link ExcelHandler}.
     */
    @Override
    public void close() {
        if (!finished) {
            try {
                ExcelWorkbookSupport.close(wb);
            } catch (RuntimeException | IOException e) {
                log.warn("Failed to close workbook", e);
            }
        }
    }
}
