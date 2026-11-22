package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** A template list writer applies the table it was configured with when write() starts, as the other writers do. */
class TemplateListTableSnapshotTest {

    private static byte[] template() throws Exception {
        try (var workbook = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Name");
            workbook.write(out);
            return out.toByteArray();
        }
    }

    @Test
    void aTableChangedWhileRowsAreWritten_isIgnoredForThatWrite() throws Exception {
        var output = new ByteArrayOutputStream();
        try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(template()), 1)) {
            TemplateListWriter<String> list = writer.<String>list(0, 1).table("Original");
            list.column("Name", name -> {
                list.table("Late");   // changes the writer while it is writing
                return name;
            });
            list.write(Stream.of("Alice", "Bob"));
            writer.finish().writeTo(output);
        }

        try (var workbook = new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
            var tables = workbook.getSheet("Data").getTables();
            assertEquals(1, tables.size());
            assertEquals("Original", tables.get(0).getName());
        }
    }
}
