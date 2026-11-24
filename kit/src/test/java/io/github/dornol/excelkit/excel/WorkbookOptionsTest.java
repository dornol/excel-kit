package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.testsupport.WorkbookSnapshot;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/** The workbook settings are written once for {@link ExcelWriter} and {@link ExcelWorkbook}, and mean the same on both. */
class WorkbookOptionsTest {
    record Row(String name, int amount) {}

    private static final List<Row> ROWS = List.of(new Row("a", 1), new Row("b", 2));

    /** A look written once, for both. */
    private static <W extends WorkbookOptions<W>> W teamLook(W writer) {
        return writer.headerColor(ExcelColor.STEEL_BLUE).headerFontName("Arial").headerFontSize(14)
                .headerStyle(h -> h.bold(true).wrapText(true))
                .documentProperty("author", "Reporting").documentProperty("Department", "Sales")
                .protectWorkbook("secret");
    }

    private static byte[] viaWriter() {
        var out = new ByteArrayOutputStream();
        teamLook(ExcelWriter.<Row>create()).column("Name", Row::name).column("Amount", Row::amount).write(ROWS).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] viaWorkbook() {
        var out = new ByteArrayOutputStream();
        try (var workbook = teamLook(ExcelWorkbook.create())) {
            workbook.<Row>sheet("Data").column("Name", Row::name).column("Amount", Row::amount).write(ROWS);
            workbook.finish().writeTo(out);
        }
        return out.toByteArray();
    }

    @Nested
    class OneLookForBoth {
        @Test
        void aLookWrittenOnceGivesTheSameHeaderOnBothWriters() throws IOException {
            assertEquals(WorkbookSnapshot.of(viaWriter()), WorkbookSnapshot.of(viaWorkbook()));
        }

        @Test
        void theLookIsReallyApplied() throws IOException {
            for (byte[] file : List.of(viaWriter(), viaWorkbook())) {
                try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                    var style = wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
                    var font = wb.getFontAt(style.getFontIndex());
                    assertEquals("Arial", font.getFontName());
                    assertEquals(14, font.getFontHeightInPoints());
                    assertTrue(font.getBold());
                    assertTrue(style.getWrapText());
                    assertTrue(wb.isStructureLocked(), "the structure of the workbook is protected");
                    assertEquals("Reporting", wb.getProperties().getCoreProperties().getCreator());
                    assertEquals("Sales", wb.getProperties().getCustomProperties().getProperty("Department").getLpwstr());
                }
            }
        }

        @Test
        void bothWritersAreOfTheOneType() {
            WorkbookOptions<ExcelWriter<Row>> writer = ExcelWriter.<Row>create();
            assertNotNull(writer);
            try (var workbook = ExcelWorkbook.create()) {
                WorkbookOptions<ExcelWorkbook> options = workbook;
                assertSame(workbook, options.headerFontName("Arial"), "a setting gives back the writer");
            }
        }

        @Test
        void thePasswordsAreThoseOfTheSettingOnBoth() throws IOException {
            for (int path = 0; path < 2; path++) {
                var out = new ByteArrayOutputStream();
                if (path == 0) {
                    ExcelWriter.<Row>create().password("pw").column("Name", Row::name).write(ROWS).writeTo(out);
                } else {
                    try (var workbook = ExcelWorkbook.create().password("pw")) {
                        workbook.<Row>sheet("S").column("Name", Row::name).write(ROWS);
                        workbook.finish().writeTo(out);
                    }
                }
                byte[] file = out.toByteArray();
                assertThrows(Exception.class, () -> new XSSFWorkbook(new ByteArrayInputStream(file)).close(), "encrypted: path " + path);
            }
        }

        @Test
        void aWrongArgumentIsRefusedTheSameWay() {
            try (var workbook = ExcelWorkbook.create()) {
                for (WorkbookOptions<?> options : List.of(ExcelWriter.<Row>create(), workbook)) {
                    String which = options.getClass().getSimpleName();
                    assertThrows(IllegalArgumentException.class, () -> options.headerFontSize(0), which);
                    assertThrows(IllegalArgumentException.class, () -> options.headerColor(null), which);
                    assertThrows(IllegalArgumentException.class, () -> options.password(" "), which);
                    assertThrows(IllegalArgumentException.class, () -> options.protectWorkbook(""), which);
                    assertThrows(NullPointerException.class, () -> options.headerFontName(null), which);
                    assertThrows(NullPointerException.class, () -> options.documentProperty("k", null), which);
                }
            }
        }
    }

    @Nested
    class OptionsGivenWhenTheWorkbookIsCreated {
        private static final java.util.function.Consumer<WorkbookInitOptions> SMALL_WINDOW = o -> o.rowAccessWindowSize(5);

        @Test
        void oneTypeServesTheWriterTheWorkbookAndForMap() throws IOException {
            var rows = IntStream.range(0, 50).mapToObj(i -> new Row("n" + i, i)).toList();
            var a = new ByteArrayOutputStream();
            ExcelWriter.<Row>create(SMALL_WINDOW).column("Name", Row::name).write(rows).writeTo(a);
            var b = new ByteArrayOutputStream();
            try (var workbook = ExcelWorkbook.create(SMALL_WINDOW)) {
                workbook.<Row>sheet("Sheet").column("Name", Row::name).write(rows);
                workbook.finish().writeTo(b);
            }
            var c = new ByteArrayOutputStream();
            ExcelWriter.forMap(SMALL_WINDOW, "Name").write(rows.stream().map(r -> Map.<String, Object>of("Name", r.name())).toList()).writeTo(c);
            assertEquals(WorkbookSnapshot.of(a.toByteArray()).subList(1, 52), WorkbookSnapshot.of(b.toByteArray()).subList(1, 52));
            assertEquals(51, new XSSFWorkbook(new ByteArrayInputStream(c.toByteArray())).getSheetAt(0).getPhysicalNumberOfRows());
        }

        @Test
        void everyOptionIsCheckedTheSameWay() {
            assertThrows(IllegalArgumentException.class, () -> ExcelWriter.create(o -> o.rowAccessWindowSize(0)));
            assertThrows(IllegalArgumentException.class, () -> ExcelWorkbook.create(o -> o.rowAccessWindowSize(-1)));
            assertThrows(NullPointerException.class, () -> ExcelWriter.create(o -> o.streaming(null)));
            assertThrows(NullPointerException.class, () -> ExcelWorkbook.create(o -> o.streaming(null)));
        }

        /** What the streaming workbook was built with: the options leave no mark in the file, so the workbook is asked. */
        private org.apache.poi.xssf.streaming.SXSSFWorkbook streamingOf(Object writer) throws ReflectiveOperationException {
            if (writer instanceof ExcelWriter<?> w) return w.getWb();
            var field = ExcelWorkbook.class.getDeclaredField("wb");
            field.setAccessible(true);
            return (org.apache.poi.xssf.streaming.SXSSFWorkbook) field.get(writer);
        }

        @Test
        void theOptionsReachTheStreamingWorkbookOfBothWriters() throws Exception {
            java.util.function.Consumer<WorkbookInitOptions> options = o -> o.rowAccessWindowSize(7).compressTempFiles(true).useSharedStrings(true);
            var writer = ExcelWriter.<Row>create(options);
            try (var workbook = ExcelWorkbook.create(options)) {
                for (Object w : new Object[]{writer, workbook}) {
                    var wb = streamingOf(w);
                    assertEquals(7, wb.getRandomAccessWindowSize(), w.getClass().getSimpleName());
                    assertTrue(wb.isCompressTempFiles(), w.getClass().getSimpleName());
                }
            }
            writer.column("Name", Row::name).write(ROWS).discard();
            var defaults = ExcelWriter.<Row>create();
            assertEquals(1000, streamingOf(defaults).getRandomAccessWindowSize(), "the default window");
            assertFalse(streamingOf(defaults).isCompressTempFiles());
            defaults.column("Name", Row::name).write(ROWS).discard();
        }

        @Test
        void streamingSetsAllAtOnceAndTheLastSettingWins() throws Exception {
            var writer = ExcelWriter.<Row>create(o -> o.rowAccessWindowSize(10).compressTempFiles(true)
                    .streaming(new StreamingOptions(20, false, false)).rowAccessWindowSize(30));
            var wb = streamingOf(writer);
            assertEquals(30, wb.getRandomAccessWindowSize());
            assertFalse(wb.isCompressTempFiles(), "streaming(...) replaced the earlier compress setting");
            writer.column("Name", Row::name).write(ROWS).discard();
        }

        @Test
        void aMissingOptionObjectIsNamed() {
            var e = assertThrows(NullPointerException.class, () -> ExcelWriter.create(o -> o.streaming(null)));
            assertEquals("options cannot be null", e.getMessage());
        }
    }
}
