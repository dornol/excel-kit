package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import java.util.ArrayList;
import java.util.List;

/** Centralized structured-table validation and POI creation. */
final class StructuredTableWriter {
    private StructuredTableWriter() {}

    static void validateName(String name) {
        if (name == null || !name.matches("[A-Za-z_][A-Za-z0-9_.]*"))
            throw new IllegalArgumentException("Invalid Excel table name: " + name);
        if (CellReference.classifyCellReference(name, SpreadsheetVersion.EXCEL2007)
                != CellReference.NameType.NAMED_RANGE)
            throw new IllegalArgumentException("Excel table name cannot be a cell reference: " + name);
    }

    static List<String> validateExistingHeaders(SXSSFSheet sheet, int headerRow, int columns) {
        XSSFSheet xssf = SXSSFSheetHelper.getXSSFSheet(sheet);
        var row = xssf.getRow(headerRow);
        if (row == null) throw new ExcelWriteException("Template table header row not found: " + headerRow);
        var headers = new ArrayList<String>(columns);
        for (int i = 0; i < columns; i++) {
            var cell = row.getCell(i);
            if (cell == null || cell.toString().isBlank())
                throw new ExcelWriteException("Template table header is blank at column " + i);
            headers.add(cell.toString());
        }
        return List.copyOf(headers);
    }

    static void apply(SXSSFSheet sheet, String name, int headerRow, int lastRow, List<String> headers,
                      String style, boolean rowStripes) {
        int columns = headers.size();
        if (columns == 0 || lastRow <= headerRow) return;
        XSSFSheet xssf = SXSSFSheetHelper.getXSSFSheet(sheet);
        var workbook = sheet.getWorkbook().getXSSFWorkbook();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++)
            for (var existing : workbook.getSheetAt(i).getTables())
                if (existing.getName() != null && existing.getName().equalsIgnoreCase(name))
                    throw new ExcelWriteException("Duplicate table name: '" + name + "'");
        AreaReference area = new AreaReference(new CellReference(headerRow, 0),
                new CellReference(lastRow, columns - 1), SpreadsheetVersion.EXCEL2007);
        var table = xssf.createTable(area);
        table.setName(name);
        table.setDisplayName(name);
        table.setStyleName(style);
        // Streaming headers may already be flushed and unavailable to XSSF.
        // Preserve their actual names instead of POI's generated "Column 1" names.
        for (int i = 0; i < columns; i++) {
            table.getCTTable().getTableColumns().getTableColumnArray(i).setName(headers.get(i));
        }
        var styleInfo = table.getCTTable().isSetTableStyleInfo()
                ? table.getCTTable().getTableStyleInfo() : table.getCTTable().addNewTableStyleInfo();
        styleInfo.setName(style);
        styleInfo.setShowRowStripes(rowStripes);
    }
}
