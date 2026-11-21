package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Rich text takes the font of its column and works in a template; a char[] is written as its characters. */
class CellTextDetailsTest {

    record Row(int id) {
    }

    private static XSSFWorkbook open(ByteArrayOutputStream out) throws Exception {
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    // ---- 1. rich text and the column's font

    @Test
    void richTextTakesTheFontOfItsColumn() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create(o -> o.useSharedStrings(true))
                .column("v", row -> new ExcelRichText().text("plain ").bold("bold").styled(" red", s -> s.color(255, 0, 0)),
                        c -> c.type(ExcelDataType.RICH_TEXT).fontName("Courier New").fontSize(14))
                .write(List.of(new Row(1))).writeTo(out);
        try (XSSFWorkbook workbook = open(out)) {
            var rich = (XSSFRichTextString) workbook.getSheetAt(0).getRow(1).getCell(0).getRichStringCellValue();
            List<CTRElt> runs = List.of(rich.getCTRst().getRArray());
            assertEquals(3, runs.size());
            for (CTRElt run : runs) {
                assertNotNull(run.getRPr(), "every run has its own font, plain ones included: " + run.getT());
                assertEquals("Courier New", run.getRPr().getRFontArray(0).getVal(), run.getT());
                assertEquals(14.0, run.getRPr().getSzArray(0).getVal(), run.getT());
            }
            assertEquals(0, runs.get(0).getRPr().sizeOfBArray(), "plain stays not bold");
            assertTrue(runs.get(1).getRPr().sizeOfBArray() > 0, "bold is added on top of the column font");
            assertTrue(runs.get(2).getRPr().sizeOfColorArray() > 0);
        }
    }

    @Test
    void aBoldColumnStaysBoldInItsPlainPartsAndAStyledPartKeepsIt() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create(o -> o.useSharedStrings(true))
                .column("v", row -> new ExcelRichText().text("a").italic("b"),
                        c -> c.type(ExcelDataType.RICH_TEXT).bold(true))
                .write(List.of(new Row(1))).writeTo(out);
        try (XSSFWorkbook workbook = open(out)) {
            var rich = (XSSFRichTextString) workbook.getSheetAt(0).getRow(1).getCell(0).getRichStringCellValue();
            for (CTRElt run : rich.getCTRst().getRArray()) {
                assertTrue(run.getRPr().sizeOfBArray() > 0, run.getT());
            }
        }
    }

    // ---- 2. rich text in a template

    private static java.io.InputStream template() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("S").createRow(0).createCell(0).setCellValue("title");
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            wb.write(bos);
            return new ByteArrayInputStream(bos.toByteArray());
        }
    }

    @Test
    void aTemplateCellKeepsRichTextOnATemplateRowAndOnANewRow() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter w = new ExcelTemplateWriter(template())) {
            w.cell(0, 1, new ExcelRichText().text("in ").bold("template"))
                    .cell(3, 0, new ExcelRichText().text("on ").italic("new row"))
                    .finish().writeTo(out);
        }
        try (XSSFWorkbook workbook = open(out)) {
            var first = (XSSFRichTextString) workbook.getSheetAt(0).getRow(0).getCell(1).getRichStringCellValue();
            assertEquals("in template", first.getString());
            assertTrue(first.getCTRst().getRArray(1).getRPr().sizeOfBArray() > 0);
            var second = (XSSFRichTextString) workbook.getSheetAt(0).getRow(3).getCell(0).getRichStringCellValue();
            assertEquals("on new row", second.getString());
            assertTrue(second.getCTRst().getRArray(1).getRPr().sizeOfIArray() > 0);
        }
    }

    // ---- 3. char[]

    @Test
    void aCharArrayIsWrittenAsItsCharacters() throws Exception {
        ByteArrayOutputStream excel = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().column("v", row -> new char[]{'a', 'b', 'c'}).write(List.of(new Row(1))).writeTo(excel);
        try (XSSFWorkbook workbook = open(excel)) {
            assertEquals("abc", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }

        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (ExcelTemplateWriter w = new ExcelTemplateWriter(template())) {
            w.cell(0, 1, new char[]{'x', 'y'}).finish().writeTo(template);
        }
        try (XSSFWorkbook workbook = open(template)) {
            assertEquals("xy", workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue());
        }

        ByteArrayOutputStream csv = new ByteArrayOutputStream();
        CsvWriter.<Row>create().bom(false).column("v", row -> new char[]{'a', ',', 'b'}).write(List.of(new Row(1))).writeTo(csv);
        String text = csv.toString(StandardCharsets.UTF_8);
        assertEquals("v\r\n\"a,b\"\r\n", text.replace("\n", "\r\n").replace("\r\r", "\r"));
        assertFalse(text.contains("[C@"));
    }
}
