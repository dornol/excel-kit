package io.github.dornol.excelkit.testsupport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * A workbook as plain text, for comparing two files that must hold the same thing: the cells (type, value, formula and
 * its cached result, style), the row and column settings, the sheet settings, the tables, the charts. The names of the
 * sheets are replaced by their position, because two writers may name sheets differently.
 */
public final class WorkbookSnapshot {
    private WorkbookSnapshot() {
    }

    public static List<String> of(byte[] xlsx) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            List<String> lines = new ArrayList<>();
            lines.add("sheets " + wb.getNumberOfSheets());
            for (int i = 0; i < wb.getNumberOfSheets(); i++) sheet(wb, i, lines);
            wb.getAllNames().stream().map(n -> "name " + n.getNameName() + " = " + normalize(wb, n.getRefersToFormula()))
                    .sorted().forEach(lines::add);
            return lines;
        }
    }

    private static String normalize(XSSFWorkbook wb, String text) {
        if (text == null) return "";
        String out = text;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            out = out.replace("'" + name.replace("'", "''") + "'!", "#" + i + "!").replace(name + "!", "#" + i + "!");
        }
        return out;
    }

    private static void sheet(XSSFWorkbook wb, int index, List<String> lines) {
        XSSFSheet sheet = wb.getSheetAt(index);
        String p = "S" + index + " ";
        lines.add(p + "state " + wb.getSheetVisibility(index) + " zoom " + sheet.getCTWorksheet().isSetSheetViews()
                + " protect " + sheet.getProtect() + " lastRow " + sheet.getLastRowNum());
        if (sheet.getCTWorksheet().isSetSheetPr() && sheet.getCTWorksheet().getSheetPr().isSetTabColor()) {
            lines.add(p + "tab " + sheet.getCTWorksheet().getSheetPr().getTabColor().xmlText());
        }
        if (sheet.getCTWorksheet().isSetSheetViews()) {
            var view = sheet.getCTWorksheet().getSheetViews().getSheetViewArray(0);
            lines.add(p + "view zoom=" + (view.isSetZoomScale() ? view.getZoomScale() : -1) + " pane=" + (view.isSetPane() ? view.getPane().xmlText() : "-"));
        }
        if (sheet.getCTWorksheet().isSetAutoFilter()) lines.add(p + "autofilter " + sheet.getCTWorksheet().getAutoFilter().getRef());
        for (CellRangeAddress region : sheet.getMergedRegions()) lines.add(p + "merge " + region.formatAsString());
        for (int b : sheet.getRowBreaks()) lines.add(p + "rowbreak " + b);
        for (int c = 0; c <= sheet.getCTWorksheet().getColsArray().length * 0 + maxColumn(sheet); c++) {
            lines.add(p + "col " + c + " width=" + sheet.getColumnWidth(c) + " hidden=" + sheet.isColumnHidden(c)
                    + " outline=" + (sheet.getCTWorksheet().getColsArray().length == 0 ? 0 : outlineOf(sheet, c)));
        }
        var cf = sheet.getSheetConditionalFormatting();
        for (int i = 0; i < cf.getNumConditionalFormattings(); i++) {
            var f = cf.getConditionalFormattingAt(i);
            StringBuilder rules = new StringBuilder();
            for (int r = 0; r < f.getNumberOfRules(); r++) {
                var rule = f.getRule(r);
                rules.append(rule.getConditionType()).append('/').append(rule.getComparisonOperation()).append('/')
                        .append(rule.getFormula1()).append('/').append(rule.getFormula2()).append(';');
            }
            StringBuilder ranges = new StringBuilder();
            for (CellRangeAddress a : f.getFormattingRanges()) ranges.append(a.formatAsString()).append(',');
            lines.add(p + "condfmt " + ranges + " " + rules);
        }
        for (var validation : sheet.getDataValidations()) {
            StringBuilder ranges = new StringBuilder();
            for (CellRangeAddress a : validation.getRegions().getCellRangeAddresses()) ranges.append(a.formatAsString()).append(',');
            var c = validation.getValidationConstraint();
            lines.add(p + "validation " + ranges + " " + c.getValidationType() + "/" + c.getOperator() + "/" + normalize(wb, c.getFormula1())
                    + "/" + c.getFormula2() + " blank=" + validation.getEmptyCellAllowed() + " err=" + validation.getErrorBoxTitle());
        }
        for (XSSFTable t : sheet.getTables()) {
            lines.add(p + "table " + t.getName().replaceAll("_\\d+$", "") + " " + t.getCTTable().getRef() + " cols=" + t.getColumnCount());
        }
        var drawing = sheet.getDrawingPatriarch();
        if (drawing != null) {
            drawing.getCharts().forEach(chart -> lines.add(p + "chart " + normalize(wb, chart.getCTChart().xmlText())
                    .replaceAll("\\s+", " ")));
        }
        for (var entry : new java.util.TreeMap<>(sheet.getCellComments()).entrySet()) {
            lines.add(p + "comment " + entry.getKey().formatAsString() + " " + entry.getValue().getString().getString());
        }
        for (var link : sheet.getHyperlinkList()) lines.add(p + "link " + link.getCellRef() + " " + link.getAddress());
        var print = sheet.getPrintSetup();
        lines.add(p + "print " + print.getLandscape() + " " + print.getPaperSize() + " " + sheet.getMargin(org.apache.poi.ss.usermodel.PageMargin.LEFT)
                + " " + sheet.getHeader().getCenter() + " " + sheet.getFooter().getCenter());
        for (Row r : sheet) {
            org.apache.poi.xssf.usermodel.XSSFRow row = (org.apache.poi.xssf.usermodel.XSSFRow) r;
            lines.add(p + "row " + row.getRowNum() + " h=" + (row.getCTRow().isSetCustomHeight() ? row.getHeightInPoints() : -1)
                    + " hidden=" + row.getZeroHeight() + " outline=" + row.getOutlineLevel());
            for (Cell cell : row) lines.add(p + "cell " + cell.getAddress().formatAsString() + " " + describe((XSSFCell) cell));
        }
    }

    private static int maxColumn(XSSFSheet sheet) {
        int max = 0;
        for (Row row : sheet) max = Math.max(max, row.getLastCellNum());
        return max;
    }

    private static int outlineOf(XSSFSheet sheet, int column) {
        for (var cols : sheet.getCTWorksheet().getColsArray()) {
            for (var col : cols.getColArray()) {
                if (col.getMin() <= column + 1 && col.getMax() >= column + 1) return col.getOutlineLevel();
            }
        }
        return 0;
    }

    private static String describe(XSSFCell cell) {
        StringBuilder out = new StringBuilder();
        CellType type = cell.getCellType();
        out.append(type).append(' ');
        switch (type) {
            case STRING -> out.append(cell.getStringCellValue());
            case NUMERIC -> out.append(cell.getNumericCellValue());
            case BOOLEAN -> out.append(cell.getBooleanCellValue());
            case FORMULA -> {
                out.append(cell.getCellFormula()).append(" => ").append(cell.getCachedFormulaResultType()).append(' ');
                switch (cell.getCachedFormulaResultType()) {
                    case NUMERIC -> out.append(cell.getNumericCellValue());
                    case STRING -> out.append(cell.getStringCellValue());
                    default -> { }
                }
            }
            default -> { }
        }
        XSSFCellStyle style = cell.getCellStyle();
        XSSFFont font = style.getFont();
        out.append(" | fmt=").append(style.getDataFormatString())
                .append(" font=").append(font.getFontName()).append('/').append(font.getFontHeightInPoints()).append('/')
                .append(font.getBold() ? "b" : "-").append(font.getItalic() ? "i" : "-").append(font.getStrikeout() ? "s" : "-")
                .append(font.getUnderline()).append('/').append(color(font.getXSSFColor()))
                .append(" fill=").append(style.getFillPattern()).append('/').append(color(style.getFillForegroundXSSFColor()))
                .append(" border=").append(style.getBorderTop()).append(style.getBorderBottom()).append(style.getBorderLeft()).append(style.getBorderRight())
                .append(" align=").append(style.getAlignment()).append('/').append(style.getVerticalAlignment())
                .append(" wrap=").append(style.getWrapText()).append(" rot=").append(style.getRotation()).append(" ind=").append(style.getIndention())
                .append(" locked=").append(style.getLocked()).append(" hidden=").append(style.getHidden());
        return out.toString();
    }

    private static String color(XSSFColor color) {
        if (color == null) return "-";
        byte[] rgb = color.getRGB();
        return rgb == null ? "theme" + color.getTheme() : String.format("%02X%02X%02X", rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
    }
}
