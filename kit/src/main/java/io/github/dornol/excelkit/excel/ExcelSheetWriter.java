package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.RowFunction;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;


/**
 * Writes data of a specific type to one or more sheets within an {@link ExcelWorkbook}.
 * <p>
 * Rolls over automatically: once a sheet holds {@link #maxRows(int)} data rows (1,000,000 by default),
 * the writer creates an additional sheet.
 * <p>
 * Column configuration uses {@link ColumnConfig} (via {@code Consumer<ColumnConfig<T>>})
 * rather than {@link ExcelColumn.ExcelColumnBuilder}. Unlike ExcelWriter's builder, there
 * is no {@code style(CellStyle)} method — styles are derived from the declarative config
 * properties (type, bold, color, borders, etc.).
 *
 * @param <T> the row data type for this sheet
 * @author dhkim
 * @since 0.3.0
 */
public class ExcelSheetWriter<T> extends AbstractSheetWriter<T, ExcelSheetWriter<T>> {

    // Shared resources from ExcelWorkbook
    private final SXSSFWorkbook wb;
    private SXSSFSheet sheet;
    private final String baseName;
    private final CellStyle headerStyle;
    private final Map<String, CellStyle> cellStyleCache;
    private final Set<String> usedSheetNames;

    // Per-sheet settings
    private final List<ExcelColumn<T>> columns = new ArrayList<>();
    private final Map<String, CellStyle> rowStyleCache = new HashMap<>();
    private final Map<String, CellStyle> headerStyleCache = new HashMap<>();
    private boolean written = false;

    ExcelSheetWriter(SXSSFWorkbook wb, SXSSFSheet sheet, String baseName,
                     CellStyle headerStyle, Map<String, CellStyle> cellStyleCache,
                     Set<String> usedSheetNames) {
        this.wb = wb;
        this.sheet = sheet;
        this.baseName = baseName;
        this.headerStyle = headerStyle;
        this.cellStyleCache = cellStyleCache;
        this.usedSheetNames = usedSheetNames;
    }

    /**
     * Adds a column using a simple function.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function) {
        Objects.requireNonNull(function, "function cannot be null");
        columns.add(buildColumn(name, (r, c) -> function.apply(r), null));
        return this;
    }

    /**
     * Adds a column with additional configuration.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @param cfg consumer to configure column styling
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, Function<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        Objects.requireNonNull(function, "function cannot be null");
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, (r, c) -> function.apply(r), config));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, RowFunction<T, @Nullable Object> function) {
        columns.add(buildColumn(name, function, null));
        return this;
    }

    /**
     * Adds a column using a row function with cursor support and additional configuration.
     *
     * @param name the column header
     * @param function function to extract the cell value
     * @param cfg consumer to configure column styling
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> column(String name, RowFunction<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        ColumnConfig<T> config = new ColumnConfig<>();
        cfg.accept(config);
        columns.add(buildColumn(name, function, config));
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
     * Adds a 1-based sequential row-number column.
     *
     * @since 0.16.11
     */
    public ExcelSheetWriter<T> rowNumberColumn(String name) {
        return column(name, (row, cursor) -> cursor.getCurrentTotal(),
                c -> c.type(ExcelDataType.LONG));
    }

    /**
     * Conditionally adds a column with additional configuration.
     *
     * @param name      Column header name
     * @param condition Whether to include this column
     * @param function  Function to extract cell value from row
     * @param cfg       Consumer to configure column options
     * @return this writer for chaining
     */
    public ExcelSheetWriter<T> columnIf(String name, boolean condition, Function<T, @Nullable Object> function, Consumer<ColumnConfig<T>> cfg) {
        if (condition) {
            column(name, function, cfg);
        }
        return this;
    }    /**
     * Writes the data stream to this sheet (with optional auto-rollover).
     *
     * <p>
     * The stream is read but <b>not closed</b>: whoever opened it (a JPA or JDBC stream, for example)
     * closes it, typically with try-with-resources around this call.
     *
     * @param stream the data stream to write
     */
    public void write(Stream<T> stream) {
        Objects.requireNonNull(stream, "stream cannot be null");
        if (written) {
            throw new ExcelWriteException("write() has already been called on this sheet");
        }
        written = true;
        if (columns.isEmpty()) {
            throw new ExcelWriteException("columns setting required");
        }
        ExcelWriteSupport.validateUniqueColumnNames(columns);

        List<SXSSFSheet> allSheets = new ArrayList<>();
        List<Integer> dataEndRows = new ArrayList<>();
        allSheets.add(this.sheet);

        int currentRow = ExcelWriteSupport.initSheetPreamble(sheet, wb, columns, cfg.beforeHeaderWriter);
        Cursor cursor = new Cursor(currentRow);
        int headerRowIndex = currentRow;

        ExcelHeaderWriter.write(sheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
        int headerRowIdx = cursor.getRowOfSheet() - 1;
        // With group headers the column names sit below the group rows: that row is the header row.
        headerRowIndex = headerRowIdx;
        ExcelWriteSupport.applySheetOptions(sheet, headerRowIdx, cfg.autoFilter, cfg.freezePaneCols, cfg.freezePaneRows, columns.size());

        SXSSFSheet activeSheet = this.sheet;

        {
            Iterator<T> it = stream.iterator();
            while (it.hasNext()) {
                T rowData = it.next();
                cursor.plusTotal();
                if (cursor.getCurrentTotal() > maxRows
                        && (cursor.getCurrentTotal() - 1) % maxRows == 0) {
                    StructuredTableWriter.requireRolloverAllowed(tableOptions);
                    dataEndRows.add(cursor.getRowOfSheet() - 1);
                    ExcelWriteSupport.writeAfterDataAndSummary(activeSheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);
                    activeSheet = createRolloverSheet(allSheets.size());
                    allSheets.add(activeSheet);
                    cursor.initRow();
                    int preambleRow = ExcelWriteSupport.initSheetPreamble(activeSheet, wb, columns, cfg.beforeHeaderWriter);
                    cursor.setRowOfSheet(preambleRow);
                    ExcelHeaderWriter.write(activeSheet, cursor, columns, headerStyle, wb, headerStyleCache, cfg.groupComments, cfg.headerRowHeightInPoints);
                    int hdrIdx = cursor.getRowOfSheet() - 1;
                    headerRowIndex = hdrIdx;
                    ExcelWriteSupport.applySheetOptions(activeSheet, hdrIdx, cfg.autoFilter, cfg.freezePaneCols, cfg.freezePaneRows, columns.size());
                }
                ExcelRowWriter.write(activeSheet, cursor, rowData, columns, cfg, rowStyleCache, wb);
                ExcelWriteSupport.checkProgress(cursor, cfg.progressInterval, cfg.progressCallback);
            }
        }

        dataEndRows.add(cursor.getRowOfSheet() - 1);
        int nextFreeRow = ExcelWriteSupport.writeAfterDataAndSummary(
                activeSheet, wb, cursor.getRowOfSheet(), columns, headerRowIndex, cfg);

        for (int i = 0; i < allSheets.size(); i++) {
            ExcelSheetPostProcessor.apply(allSheets.get(i), columns, headerRowIndex, dataEndRows.get(i), cfg, i);
        }
        if (tableOptions != null) {
            for (int i = 0; i < allSheets.size(); i++) {
                SXSSFSheet target = allSheets.get(i);
                if (target.getLastRowNum() <= headerRowIndex) continue;
                String name = allSheets.size() == 1 ? tableOptions.name() : tableOptions.name() + "_" + (i + 1);
                StructuredTableWriter.apply(target, name, headerRowIndex, target.getLastRowNum(), columns.stream().map(ExcelColumn::getName).toList(),
                        tableOptions.style(), tableOptions.showRowStripes());
            }
        }

        // One chart per sheet over that sheet's own rows: a chart of only the last sheet would plot just the remainder.
        if (cfg.chartConfig != null) {
            for (int i = 0; i < allSheets.size(); i++) {
                boolean last = i == allSheets.size() - 1;
                int dataEnd = dataEndRows.get(i);
                ExcelWriteSupport.applyChart(allSheets.get(i), cfg.chartConfig, headerRowIndex, dataEnd,
                        last ? nextFreeRow : dataEnd + 1, columns.size());
            }
        }
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

    private ExcelColumn<T> buildColumn(String name, RowFunction<T, @Nullable Object> function,
                                       @Nullable ColumnConfig<T> config) {
        return ExcelWriteSupport.buildColumn(wb, cellStyleCache, cfg.defaultStyleConfig, name, function, config);
    }

}
