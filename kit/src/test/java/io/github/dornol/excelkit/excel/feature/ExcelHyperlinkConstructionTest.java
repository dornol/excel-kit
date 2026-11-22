package io.github.dornol.excelkit.excel.feature;

import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.feature.ExcelHyperlink;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelHyperlinkConstructionTest {

    // -----------------------------------------------------------------------
    // 5. ExcelHyperlink — single-arg constructor
    // -----------------------------------------------------------------------

    @Test
    void excelHyperlink_singleArgConstructor() {
        ExcelHyperlink link = new ExcelHyperlink("https://example.com");
        assertEquals("https://example.com", link.url());
        assertEquals("https://example.com", link.label());
    }

    @Test
    void excelHyperlink_twoArgConstructor() {
        ExcelHyperlink link = new ExcelHyperlink("https://example.com", "Example");
        assertEquals("https://example.com", link.url());
        assertEquals("Example", link.label());
    }

    @Test
    void excelHyperlink_singleArg_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<String>create()
                .column("Link", url -> new ExcelHyperlink(url), cfg -> cfg.type(ExcelDataType.HYPERLINK))
                .write(Stream.of("https://example.com"))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            // single-arg: label == url
            assertEquals("https://example.com", cell.getStringCellValue());
            assertNotNull(cell.getHyperlink());
            assertEquals("https://example.com", cell.getHyperlink().getAddress());
        }
    }
}
