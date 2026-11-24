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
public class ExcelWorkbook implements AutoCloseable, WorkbookOptions<ExcelWorkbook> {
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
     * @param configurer consumer that configures {@link WorkbookInitOptions}
     * @return a new workbook instance (implements AutoCloseable)
     * @since 0.17.0
     */
    public static ExcelWorkbook create(Consumer<WorkbookInitOptions> configurer) {
        WorkbookInitOptions opts = new WorkbookInitOptions();
        configurer.accept(opts);
        return new ExcelWorkbook(opts);
    }

    ExcelWorkbook(WorkbookInitOptions opts) {
        StreamingOptions streaming = opts.options();
        this.wb = ExcelWorkbookSupport.streamingWorkbook(streaming);
        this.settings = new WorkbookSettings(wb);
    }

    @Override
    public ExcelWorkbook documentProperty(String key, String value) {
        settings.documentProperty(key, value);
        return this;
    }

    @Override
    public ExcelWorkbook headerColor(ExcelColor color) {
        settings.headerColor(color);
        return this;
    }

    @Override
    public ExcelWorkbook protectWorkbook(String password) {
        settings.protectWorkbook(password);
        return this;
    }

    @Override
    public ExcelWorkbook password(String password) {
        settings.password(password);
        return this;
    }

    @Override
    public ExcelWorkbook headerFontName(String fontName) {
        settings.headerFontName(fontName);
        return this;
    }

    @Override
    public ExcelWorkbook headerFontSize(int fontSize) {
        settings.headerFontSize(fontSize);
        return this;
    }

    @Override
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
        CellStyle headerStyle = settings.headerStyle();
        return new ExcelSheetWriter<>(wb, wb.createSheet(name), name, () -> headerStyle, cellStyleCache, headerStyleCache, rowStyleCache, usedSheetNames);
    }

    /**
     * The one sheet of an {@link ExcelWriter}: created, named and styled when it is written, so that the settings of
     * the writer can come after it.
     */
    <T> ExcelSheetWriter<T> lazySheet() {
        return new ExcelSheetWriter<>(wb, null, "Sheet", settings::headerStyle, cellStyleCache, headerStyleCache, rowStyleCache, usedSheetNames);
    }

    SXSSFWorkbook wb() {
        return wb;
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
