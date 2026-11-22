package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Locale;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** The default label of a summary row does not depend on the JVM's locale. */
class SummaryLabelLocaleTest {

    @Test
    void defaultLabelsAreTheSameUnderTurkishLocale() throws Exception {
        Locale original = Locale.getDefault();
        Locale.setDefault(Locale.forLanguageTag("tr-TR"));
        try {
            var handler = ExcelWriter.<Integer>create()
                    .column("Name", i -> "n" + i)
                    .column("Price", i -> i, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.min("Price").max("Price"))
                    .write(Stream.of(1, 2));
            var out = new ByteArrayOutputStream();
            handler.writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Min", sheet.getRow(3).getCell(0).getStringCellValue());
                assertEquals("Max", sheet.getRow(4).getCell(0).getStringCellValue());
            }
        } finally {
            Locale.setDefault(original);
        }
    }
}
