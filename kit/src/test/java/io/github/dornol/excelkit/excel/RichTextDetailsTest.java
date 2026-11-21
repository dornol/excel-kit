package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** What a rich text cell holds in the file: its runs with their formatting, and what happens where they cannot be kept. */
class RichTextDetailsTest {

    record Row(int id) {
    }

    /** A run as read back: its text and what its font changes. */
    private record Run(String text, boolean bold, boolean italic, boolean underline, boolean strike, Float size, String color) {
    }

    private static List<Run> write(Supplier<ExcelRichText> text, boolean sharedStrings) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create(o -> o.useSharedStrings(sharedStrings))
                .column("v", row -> text.get(), c -> c.type(ExcelDataType.RICH_TEXT))
                .write(List.of(new Row(1))).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var rich = (XSSFRichTextString) workbook.getSheetAt(0).getRow(1).getCell(0).getRichStringCellValue();
            List<Run> runs = new ArrayList<>();
            for (CTRElt run : rich.getCTRst().getRArray()) {
                var properties = run.getRPr();
                runs.add(new Run(run.getT(),
                        properties != null && properties.sizeOfBArray() > 0,
                        properties != null && properties.sizeOfIArray() > 0,
                        properties != null && properties.sizeOfUArray() > 0,
                        properties != null && properties.sizeOfStrikeArray() > 0,
                        properties != null && properties.sizeOfSzArray() > 0 ? (float) properties.getSzArray(0).getVal() : null,
                        properties != null && properties.sizeOfColorArray() > 0 && properties.getColorArray(0).getRgb() != null
                                ? HexFormat.of().formatHex(properties.getColorArray(0).getRgb()) : null));
            }
            return runs;
        }
    }

    // ---- the runs

    @Test
    void everyStyledSegmentIsARunWithItsFormatting() throws Exception {
        List<Run> runs = write(() -> new ExcelRichText().text("Status: ").bold("OK")
                .styled(" fine", s -> s.italic(true).underline(true).strikethrough(true).color(255, 0, 0).fontSize(14)), true);

        assertEquals(3, runs.size(), runs.toString());
        assertEquals("Status: ", runs.get(0).text());
        assertFalse(runs.get(0).bold() || runs.get(0).italic(), "plain text carries no formatting");
        assertEquals("OK", runs.get(1).text());
        assertTrue(runs.get(1).bold());
        assertFalse(runs.get(1).italic());
        assertEquals(" fine", runs.get(2).text());
        assertTrue(runs.get(2).italic() && runs.get(2).underline() && runs.get(2).strike());
        assertFalse(runs.get(2).bold());
        assertEquals(14f, runs.get(2).size());
        assertEquals("ffff0000", runs.get(2).color());
    }

    @Test
    void italicAndColorPresetsAreRuns() throws Exception {
        List<Run> runs = write(() -> new ExcelRichText().italic("slanted").styled("blue", s -> s.color(ExcelColor.BLUE)), true);
        assertTrue(runs.get(0).italic());
        assertEquals("ff0000ff", runs.get(1).color());
    }

    @Test
    void emptySegmentsLeaveNoRunAndNoSegmentsLeavesAnEmptyCell() throws Exception {
        List<Run> runs = write(() -> new ExcelRichText().text("").bold("").text("a").italic("").italic("b"), true);
        assertEquals(List.of("a", "b"), runs.stream().map(Run::text).toList());
        assertTrue(runs.get(1).italic());
        // nothing styled and not empty: the cell is plain text, with no run at all
        assertEquals(0, write(() -> new ExcelRichText().text("").bold("").text("a"), true).size());
        assertEquals(0, write(() -> new ExcelRichText(), true).size());
    }

    @Test
    void charactersXmlCannotCarryAreEscapedPerRunAndComeBackAsTheCharacters() throws Exception {
        List<Run> runs = write(() -> new ExcelRichText().text("a").bold("b\u0001c").text("d_x0041_e"), true);
        assertEquals("b_x0001_c", runs.get(1).text());
        assertTrue(runs.get(1).bold());
        assertEquals("d_x005F_x0041_e", runs.get(2).text(), "a literal escape sequence has its underscore escaped");
    }

    @Test
    void anEmojiSplitBetweenSegmentsIsKeptWhole() throws Exception {
        // the halves of one character are in two segments: the run boundary moves behind the character
        List<Run> runs = write(() -> new ExcelRichText().bold("a\uD83D").italic("\uDE00b"), true);
        assertEquals(2, runs.size(), runs.toString());
        assertEquals("a😀", runs.get(0).text());
        assertTrue(runs.get(0).bold());
        assertEquals("b", runs.get(1).text());
        assertTrue(runs.get(1).italic());

        // a segment that is only the second half disappears into the one before it
        List<Run> only = write(() -> new ExcelRichText().bold("\uD83D").italic("\uDE00"), true);
        assertEquals(List.of("😀"), only.stream().map(Run::text).toList());
        assertTrue(only.get(0).bold());

        // whole characters and pairs elsewhere are untouched
        List<Run> whole = write(() -> new ExcelRichText().bold("😀").italic("😁x"), true);
        assertEquals(List.of("😀", "😁x"), whole.stream().map(Run::text).toList());
    }

    // ---- where the formatting cannot be kept

    @Test
    void withoutSharedStringsTheFormattingWouldBeLostSoTheWriteIsRefused() {
        ExcelWriteException error = assertThrows(ExcelWriteException.class,
                () -> write(() -> new ExcelRichText().bold("x"), false));
        String messages = error.getMessage() + " / " + error.getCause().getMessage();
        assertTrue(messages.contains("useSharedStrings(true)"), messages);
    }

    @Test
    void theLenientPolicyWritesThePlainText() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT)
                .column("v", row -> new ExcelRichText().text("a ").bold("b"), c -> c.type(ExcelDataType.RICH_TEXT))
                .write(List.of(new Row(1))).writeTo(out);
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // the text is kept, as the text of the object, which is the concatenated segments
            assertEquals("a b", workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void anExcelWorkbookWithSharedStringsKeepsTheRuns() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create(o -> o.useSharedStrings(true))) {
            workbook.<Row>sheet("S").column("v", row -> new ExcelRichText().text("a").bold("b"), c -> c.type(ExcelDataType.RICH_TEXT))
                    .write(List.of(new Row(1)));
            workbook.finish().writeTo(out);
        }
        try (XSSFWorkbook workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(2, workbook.getSheetAt(0).getRow(1).getCell(0).getRichStringCellValue().numFormattingRuns());
        }
    }

    @Test
    void plainStringsAreStillWrittenWithoutSharedStrings() {
        assertDoesNotThrow(() -> ExcelWriter.<Row>create().column("v", row -> "plain").write(List.of(new Row(1)))
                .writeTo(new ByteArrayOutputStream()));
    }

    // ---- input

    @Test
    void aFontSizeIsBetween1And409Points() {
        assertDoesNotThrow(() -> new ExcelRichText().styled("a", s -> s.fontSize(1)));
        assertDoesNotThrow(() -> new ExcelRichText().styled("a", s -> s.fontSize(409)));
        for (int bad : new int[]{0, -5, 410, 70_000}) {
            assertThrows(IllegalArgumentException.class, () -> new ExcelRichText().styled("a", s -> s.fontSize(bad)), "size " + bad);
        }
    }

    @Test
    void textAndConfigurerMustBeGiven() {
        assertThrows(NullPointerException.class, () -> new ExcelRichText().text(null));
        assertThrows(NullPointerException.class, () -> new ExcelRichText().bold(null));
        assertThrows(NullPointerException.class, () -> new ExcelRichText().italic(null));
        assertThrows(NullPointerException.class, () -> new ExcelRichText().styled(null, s -> { }));
        assertThrows(NullPointerException.class, () -> new ExcelRichText().styled("a", null));
    }

    @Test
    void colorComponentsAreChecked() {
        assertThrows(IllegalArgumentException.class, () -> new ExcelRichText().styled("a", s -> s.color(256, 0, 0)));
        assertDoesNotThrow(() -> new ExcelRichText().styled("a", s -> s.color(255, 255, 255)));
    }

    @Test
    void toStringIsThePlainText() {
        assertEquals("Hello world", new ExcelRichText().bold("Hello").text(" world").toString());
        assertEquals("", new ExcelRichText().toString());
    }
}
