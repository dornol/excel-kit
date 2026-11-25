package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelDataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.*;

/** Every preset is a format Excel (here POI's formatter) shows the way its description says. */
class ExcelDataFormatPresetTest {

    private static final double NUMBER = 1234.567;
    private static final double FRACTION = 0.1234;
    /** 2025-07-19 14:23:45 as an Excel serial number. */
    private static final double MOMENT = org.apache.poi.ss.usermodel.DateUtil.getExcelDate(LocalDateTime.of(2025, 7, 19, 14, 23, 45));

    private static String show(ExcelDataFormat preset, double value) throws Exception {
        try (var wb = new XSSFWorkbook()) {
            var cell = wb.createSheet("S").createRow(0).createCell(0);
            cell.setCellValue(value);
            var style = wb.createCellStyle();
            style.setDataFormat(wb.createDataFormat().getFormat(preset.getFormat()));
            cell.setCellStyle(style);
            return ascii(new DataFormatter(Locale.US).formatCellValue(cell));
        }
    }

    /** POI writes a date with the digits of the JVM's default locale (Arabic-Indic under ar-EG); the formats are meant, not the digits. */
    private static String ascii(String text) {
        StringBuilder sb = new StringBuilder(text.length());
        for (char c : text.toCharArray()) sb.append(Character.isDigit(c) ? (char) ('0' + Character.digit(c, 10)) : c);
        return sb.toString();
    }

    @Test
    void numbersPercentsAndCurrenciesShowAsDescribed() throws Exception {
        assertEquals("1235", show(ExcelDataFormat.NUMBER_PLAIN, NUMBER));
        assertEquals("1,235", show(ExcelDataFormat.NUMBER, NUMBER));
        assertEquals("1,234.6", show(ExcelDataFormat.NUMBER_1, NUMBER));
        assertEquals("1,234.57", show(ExcelDataFormat.NUMBER_2, NUMBER));
        assertEquals("12%", show(ExcelDataFormat.PERCENT_0, FRACTION));
        assertEquals("12.3%", show(ExcelDataFormat.PERCENT_1, FRACTION));
        assertEquals("12.34%", show(ExcelDataFormat.PERCENT, FRACTION));
        assertEquals("1,235원", show(ExcelDataFormat.CURRENCY_KRW, NUMBER));
        assertEquals("$1,234.57", show(ExcelDataFormat.CURRENCY_USD, NUMBER));
        assertEquals("€1,234.57", show(ExcelDataFormat.CURRENCY_EUR, NUMBER));
        assertEquals("¥1,235", show(ExcelDataFormat.CURRENCY_JPY, NUMBER));
    }

    @Test
    void datesAndTimesShowAsDescribed() throws Exception {
        assertEquals("2025-07-19 14:23:45", show(ExcelDataFormat.DATETIME, MOMENT));
        assertEquals("2025-07-19 14:23", show(ExcelDataFormat.DATETIME_HM, MOMENT));
        assertEquals("2025-07-19", show(ExcelDataFormat.DATE, MOMENT));
        assertEquals("2025.07.19", show(ExcelDataFormat.DATE_DOT, MOMENT));
        assertEquals("2025/07/19", show(ExcelDataFormat.DATE_SLASH, MOMENT));
        assertEquals("2025-07", show(ExcelDataFormat.YEAR_MONTH, MOMENT));
        assertEquals("14:23:45", show(ExcelDataFormat.TIME, MOMENT));
        assertEquals("14:23", show(ExcelDataFormat.TIME_HM, MOMENT));
    }

    @Test
    void theTextPresetKeepsLeadingZerosAndIsExcelsTextFormat() throws Exception {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Object>create()
                .column("Plain", x -> "0012")
                .column("Text", x -> "0012", c -> c.format(ExcelDataFormat.TEXT))
                .write(List.of(new Object())).writeTo(out);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1);
            assertEquals("General", row.getCell(0).getCellStyle().getDataFormatString());
            assertEquals("0012", row.getCell(1).getStringCellValue());
            assertEquals(49, row.getCell(1).getCellStyle().getDataFormat(), "49 is Excel's built-in Text format");
            assertEquals("@", row.getCell(1).getCellStyle().getDataFormatString());
        }
    }

    @Test
    void formatTakesAPresetAndRefusesNull() {
        assertThrows(NullPointerException.class, () -> ExcelWriter.<Object>create()
                .column("A", x -> "a", c -> c.format((ExcelDataFormat) null)));
    }
}
