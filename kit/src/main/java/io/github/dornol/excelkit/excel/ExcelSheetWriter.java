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
import java.util.function.Supplier;
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
    private @Nullable SXSSFSheet sheet; // null until write() for a sheet that is named late (ExcelWriter)
    private final String baseName;
    private final Supplier<CellStyle> headerStyle; // read at write(): ExcelWriter's header settings may come after the writer
    private final Map<String, CellStyle> cellStyleCache;
    private final Set<String> usedSheetNames;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache;
    private final Map<String, CellStyle> headerStyleCache;
    private boolean written = false;

    ExcelSheetWriter(SXSSFWorkbook wb, @Nullable SXSSFSheet sheet, String baseName,
                     Supplier<CellStyle> headerStyle, Map<String, CellStyle> cellStyleCache,
                     Map<String, CellStyle> headerStyleCache,
                     Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> rowStyleCache,
                     Set<String> usedSheetNames) {
        super(new SheetConfig<>());
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
        write(stream, (rowData, cursor) -> {});
    }

    /**
     * Writes the data stream to this sheet, calling {@code consumer} after each row with the row and the cursor.
     * Otherwise as {@link #write(Stream)}.
     *
     * @param stream   the data stream to write
     * @param consumer called after each row is written
     */
    public void write(Stream<T> stream, WriteRowCallback<T> consumer) {
        Objects.requireNonNull(stream, "stream cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        if (written) {
            throw new ExcelWriteException("write() has already been called on this sheet");
        }
        check();
        written = true;
        // What changes while the data is written does not change this sheet: the settings are taken now.
        List<ExcelColumn<T>> writtenColumns = List.copyOf(columns);
        SheetConfig<T> config = new SheetConfig<>(cfg);
        SheetWriteEngine<T> engine = new SheetWriteEngine<>(wb, writtenColumns, config, headerStyle.get(), headerStyleCache,
                rowStyleCache, cellStyleCache, this::createRolloverSheet);
        engine.start(firstSheet());
        try {
            Iterator<T> it = stream.sequential().iterator();
            while (it.hasNext()) {
                T rowData = it.next();
                engine.row(rowData);
                consumer.accept(rowData, engine.cursor());
            }
            int nextRow = engine.finishData();
            if (config.afterAllWriter != null) {
                config.afterAllWriter.write(new SheetContext(engine.sheet(), wb, nextRow, writtenColumns, engine.headerRow()));
            }
            engine.finishSheets();
        } catch (ExcelWriteException e) {
            throw e;
        } catch (RuntimeException e) {
            throw new ExcelWriteException("Failed to write excel", e);
        }
    }

    /** The checks that need no data: what {@link #write(Stream)} rejects before it writes anything. */
    void check() {
        SheetWriteEngine.validate(columns, cfg);
    }

    /** The first sheet: created with the workbook, or at the first write when it is named by the writer's settings. */
    private SXSSFSheet firstSheet() {
        if (sheet == null) {
            String name = cfg.sheetNameFunction != null ? cfg.sheetNameFunction.apply(0) : baseName;
            if (!usedSheetNames.add(name)) {
                throw new ExcelWriteException("Duplicate sheet name: '" + name + "'");
            }
            sheet = wb.createSheet(name);
        }
        return sheet;
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
        if (!usedSheetNames.add(name)) {
            throw new ExcelWriteException("Duplicate sheet name: '" + name + "'");
        }
        return wb.createSheet(name);
    }

    List<ExcelColumn<T>> columns() {
        return columns;
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
