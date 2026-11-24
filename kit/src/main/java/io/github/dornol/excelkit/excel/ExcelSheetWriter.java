package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.RowFunction;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;


/**
 * Writes data of a specific type to one or more sheets within an {@link ExcelWorkbook}.
 * <p>
 * Rolls over automatically: once a sheet holds {@link #maxRows(int)} data rows (1,000,000 by default),
 * the writer creates an additional sheet.
 * <p>
 * Columns and sheet settings are those of {@link SheetWriter}, shared with {@link ExcelWriter}; a column is configured
 * with a {@code Consumer<ColumnConfig<T>>}.
 *
 * @param <T> the row data type for this sheet
 * @author dhkim
 * @since 0.3.0
 */
public class ExcelSheetWriter<T> extends SheetWriter<T, ExcelSheetWriter<T>> {

    // Shared resources from ExcelWorkbook
    private final SXSSFWorkbook wb;
    private SXSSFSheet sheet;
    private final String baseName;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache;
    private final Set<String> usedSheetNames;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache;
    private final Map<String, CellStyle> headerStyleCache;
    private boolean written = false;

    ExcelSheetWriter(SXSSFWorkbook wb, SXSSFSheet sheet, String baseName,
                     CellStyle headerStyle, Map<String, CellStyle> cellStyleCache,
                     Map<String, CellStyle> headerStyleCache,
                     Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache,
                     Set<String> usedSheetNames) {
        this.headerStyleCache = headerStyleCache;
        this.rowStyleCache = rowStyleCache;
        this.wb = wb;
        this.sheet = sheet;
        this.baseName = baseName;
        this.headerStyle = headerStyle;
        this.cellStyleCache = cellStyleCache;
        this.usedSheetNames = usedSheetNames;
    }

    /**
     * Hides the sheet from the tab bar; the person opening the file can show it again (right-click on a tab, Unhide).
     * Use it for a sheet that only feeds the others, such as the master codes behind a dropdown
     * ({@code ExcelValidation.listFromRange("Codes!$A$2:$A$20")}). Every sheet of a rollover is hidden. The
     * workbook must keep a visible sheet: {@link ExcelWorkbook#finish()} throws {@link ExcelWriteException}
     * otherwise.
     *
     * @return this writer for chaining
     * @since 0.29.0
     */
    public ExcelSheetWriter<T> hideSheet() {
        cfg.visibility = org.apache.poi.ss.usermodel.SheetVisibility.HIDDEN;
        return this;
    }

    /**
     * Hides the sheet so that it cannot be shown from Excel's Unhide menu (only from the VBA editor or a program).
     * Otherwise as {@link #hideSheet()}.
     *
     * @return this writer for chaining
     * @since 0.29.0
     */
    public ExcelSheetWriter<T> veryHideSheet() {
        cfg.visibility = org.apache.poi.ss.usermodel.SheetVisibility.VERY_HIDDEN;
        return this;
    }

    /**
     * Enables auto-filter on the header row (no-arg convenience).
     *
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> autoFilter() {
        return autoFilter(true);
    }

    /**
     * Writes the data stream to this sheet (with optional auto-rollover).
     *
     * <p>
     * The stream is read but <b>not closed</b>: whoever opened it (a JPA or JDBC stream, for example)
     * closes it, typically with try-with-resources around this call.
     * <p>
     * Rows are written one at a time on the calling thread, in the stream's encounter order: a parallel
     * stream is read sequentially, so {@code parallel()} brings no speed-up here.
     *
     * @param stream the data stream to write
     */
    public void write(Stream<T> stream) {
        Objects.requireNonNull(stream, "stream cannot be null");
        if (written) {
            throw new ExcelWriteException("write() has already been called on this sheet");
        }
        written = true;
        SheetWriteEngine.validate(columns, cfg);
        SheetWriteEngine<T> engine = new SheetWriteEngine<>(wb, columns, cfg, headerStyle, headerStyleCache, rowStyleCache,
                cellStyleCache, this::createRolloverSheet);
        engine.start(this.sheet);
        Iterator<T> it = stream.iterator();
        while (it.hasNext()) {
            engine.row(it.next());
        }
        engine.finishData();
        engine.finishSheets();
    }

    /** Writes rows from an Iterable without copying them. */
    public void write(Iterable<T> rows) {
        Objects.requireNonNull(rows, "rows cannot be null");
        write(StreamSupport.stream(rows.spliterator(), false));
    }

    private SXSSFSheet createRolloverSheet(int rolloverIndex) {
        String name;
        if (cfg.sheetNameFunction != null) {
            name = cfg.sheetNameFunction.apply(rolloverIndex);
        } else {
            name = ExcelWriteSupport.rolloverSheetName(baseName, rolloverIndex);
        }
        if (usedSheetNames != null && !usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name: '" + name + "'");
        }
        return wb.createSheet(name);
    }

    @Override
    void addColumn(ExcelColumn<T> column) {
        columns.add(column);
    }

    @Override
    ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function,
                               @Nullable ColumnConfig<T> config) {
        return ExcelWriteSupport.buildColumn(wb, cellStyleCache, cfg.defaultStyleConfig, name, function, config);
    }

}
