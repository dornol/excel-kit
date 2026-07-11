package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

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

    static void validateExistingHeaders(SXSSFSheet sheet, int headerRow, int columns) {
        XSSFSheet xssf = SXSSFSheetHelper.getXSSFSheet(sheet);
        var row = xssf == null ? null : xssf.getRow(headerRow);
        if (row == null) throw new ExcelWriteException("Template table header row not found: " + headerRow);
        for (int i = 0; i < columns; i++) {
            var cell = row.getCell(i);
            if (cell == null || cell.toString().isBlank())
                throw new ExcelWriteException("Template table header is blank at column " + i);
        }
    }

    static void apply(SXSSFSheet sheet, String name, int headerRow, int lastRow, int columns,
                      String style, boolean rowStripes) {
        XSSFSheet xssf = SXSSFSheetHelper.getXSSFSheet(sheet);
        if (xssf == null || columns == 0 || lastRow <= headerRow) return;
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
        var styleInfo = table.getCTTable().isSetTableStyleInfo()
                ? table.getCTTable().getTableStyleInfo() : table.getCTTable().addNewTableStyleInfo();
        styleInfo.setName(style);
        styleInfo.setShowRowStripes(rowStripes);
    }
}
