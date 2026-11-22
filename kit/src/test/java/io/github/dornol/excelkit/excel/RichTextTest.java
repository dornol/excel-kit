package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class RichTextTest {

    @Test
    void richText_plainText() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText().text("Hello World"),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Hello World", cell.getRichStringCellValue().getString(),
                    "Plain text segment should produce correct cell value");
        }
    }

    @Test
    void richText_boldText() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText().text("Normal ").bold("Bold"),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Normal Bold", cell.getRichStringCellValue().getString(),
                    "Bold rich text should produce correct concatenated string");
        }
    }

    @Test
    void richText_italicText() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText().text("Normal ").italic("Italic"),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Normal Italic", cell.getRichStringCellValue().getString(),
                    "Italic rich text should produce correct concatenated string");
        }
    }

    @Test
    void richText_styledWithMultipleOptions() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText()
                                .text("Normal ")
                                .styled("Fancy", style -> style
                                        .color(255, 0, 0)
                                        .fontSize(16)
                                        .underline(true)
                                        .strikethrough(true)),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Normal Fancy", cell.getRichStringCellValue().getString(),
                    "Styled rich text should produce correct concatenated string");
        }
    }

    @Test
    void richText_multipleSegments() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText()
                                .text("Hello ")
                                .bold("World")
                                .text(" - ")
                                .italic("End"),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            var rts = (XSSFRichTextString) cell.getRichStringCellValue();
            assertEquals("Hello World - End", rts.getString(),
                    "Multiple segments should be concatenated correctly");
        }
    }

    @Test
    void richText_toStringReturnsPlainText() {
        ExcelRichText rt = new ExcelRichText()
                .text("Hello ")
                .bold("World")
                .italic("!")
                .styled("end", s -> s.color(255, 0, 0).fontSize(20));

        assertEquals("Hello World!end", rt.toString(),
                "toString() should return plain text without formatting");
    }

    @Test
    void richText_dataTypeWritesCorrectly() throws IOException {
        ExcelRichText rt = new ExcelRichText()
                .bold("Important: ")
                .text("read this");

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> rt, c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertTrue(cell.getRichStringCellValue() instanceof XSSFRichTextString,
                    "RICH_TEXT data type should produce XSSFRichTextString");
            assertEquals("Important: read this", cell.getRichStringCellValue().getString());
        }
    }

    @Test
    void richText_fallbackWhenNotExcelRichText() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> "plain string value",
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("plain string value", cell.getStringCellValue(),
                    "RICH_TEXT type should fall back to String.valueOf when value is not ExcelRichText");
        }
    }

    @Test
    void richText_withExcelColorPreset() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText()
                                .styled("Red text", style -> style.color(ExcelColor.RED))
                                .styled("Blue text", style -> style.color(ExcelColor.BLUE)),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Red textBlue text", cell.getRichStringCellValue().getString(),
                    "Rich text with ExcelColor presets should have correct content");
        }
    }

    @Test
    void richText_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create(o -> o.useSharedStrings(true))) {
            workbook.<String>sheet("RichSheet")
                    .column("Col", s -> new ExcelRichText()
                                    .text("Normal ")
                                    .bold("Bold"),
                            c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(Stream.of("data"));
            workbook.finish().writeTo(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Normal Bold", cell.getRichStringCellValue().getString(),
                    "Rich text should work in ExcelSheetWriter (ExcelWorkbook)");
        }
    }

    @Test
    void richText_emptySegments() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create(o -> o.useSharedStrings(true))
                .column("Col", s -> new ExcelRichText()
                                .text("")
                                .bold("")
                                .text("Actual"),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(Stream.of("row"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals("Actual", cell.getRichStringCellValue().getString(),
                    "Empty segments should not affect the final text content");
        }
    }
}
