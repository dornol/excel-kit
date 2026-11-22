package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * An application that lets users bring their own template gets a library exception for a file it cannot use, not
 * whatever POI or its XML layer throws. A template that cannot even be read from the stream still gives an
 * {@link IOException}, which the constructor declares.
 */
class TemplateInputFailuresTest {

    private static byte[] validTemplate() throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            wb.createSheet("S").createRow(0).createCell(0).setCellValue("Title");
            wb.write(out);
            return out.toByteArray();
        }
    }

    /** The template with the text of one of its parts changed. */
    private static byte[] withPart(byte[] xlsx, String part, java.util.function.UnaryOperator<String> change) throws Exception {
        Path tmp = Files.createTempFile("template", ".xlsx");
        try {
            Files.write(tmp, xlsx);
            Map<String, byte[]> parts = new LinkedHashMap<>();
            try (ZipFile zip = new ZipFile(tmp.toFile())) {
                for (var e = zip.entries(); e.hasMoreElements(); ) {
                    ZipEntry entry = e.nextElement();
                    parts.put(entry.getName(), zip.getInputStream(entry).readAllBytes());
                }
            }
            String text = new String(parts.get(part), StandardCharsets.UTF_8);
            String changed = change.apply(text);
            assertTrue(!changed.equals(text), "the change had no effect on " + part);
            parts.put(part, changed.getBytes(StandardCharsets.UTF_8));
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ZipOutputStream zip = new ZipOutputStream(out)) {
                for (var e : parts.entrySet()) {
                    zip.putNextEntry(new ZipEntry(e.getKey()));
                    zip.write(e.getValue());
                    zip.closeEntry();
                }
            }
            return out.toByteArray();
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    @Test
    void aFileThatIsNotAWorkbookIsRefusedWithALibraryException() {
        ExcelWriteException text = assertThrows(ExcelWriteException.class,
                () -> new ExcelTemplateWriter(new ByteArrayInputStream("hello, this is not a workbook".getBytes(StandardCharsets.UTF_8))));
        assertEquals("Failed to read the template: it is not a usable .xlsx workbook", text.getMessage());
        assertNotNull(text.getCause(), "what POI said stays as the cause");

        assertThrows(ExcelKitException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(new byte[0])));
    }

    @Test
    void aTruncatedWorkbookIsRefusedAsALibraryExceptionOrAsTheIoExceptionOfTheConstructor() throws Exception {
        byte[] whole = validTemplate();
        byte[] cut = java.util.Arrays.copyOf(whole, whole.length / 2);

        Throwable thrown = assertThrows(Throwable.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(cut)));

        assertTrue(thrown instanceof ExcelKitException || thrown instanceof IOException, String.valueOf(thrown));
    }

    @Test
    void aWorkbookWithoutSheetsIsRefused() throws Exception {
        byte[] noSheets;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            wb.write(out);
            noSheets = out.toByteArray();
        }

        ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(noSheets)));

        assertEquals("Failed to read the template: it has no sheets", e.getMessage());
    }

    @Test
    void theArgumentsAreStillCheckedFirst() throws Exception {
        byte[] valid = validTemplate();

        assertThrows(NullPointerException.class, () -> new ExcelTemplateWriter(null));
        assertThrows(IllegalArgumentException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(valid), 0));
        assertThrows(IllegalArgumentException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(valid), -2));
        // not even looking at the stream: a window that is refused wins over a stream that is garbage
        assertThrows(IllegalArgumentException.class, () -> new ExcelTemplateWriter(new ByteArrayInputStream(new byte[]{1, 2}), 0));
    }

    @Test
    void aTemplateWhoseStylesCannotBeUsedFailsTheCellWriteWithALibraryException() throws Exception {
        // every number format id of the styles is garbage: the template opens, but a cell cannot be given a format
        byte[] damaged = withPart(validTemplate(), "xl/styles.xml", text -> text.replace("numFmtId=\"0\"", "numFmtId=\"A0\""));

        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(damaged))) {
            ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> writer.cell("A1", LocalDate.of(2024, 1, 5)));

            assertEquals("Failed to write the cell at row 0, column 0", e.getMessage());
            assertNotNull(e.getCause());
        }
    }

    @Test
    void aFailureWhileTheListRowsAreWrittenIsALibraryExceptionAsForExcelWriter() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(validTemplate()))) {
            var list = writer.<String>list(3).column("v", v -> v).rowColor(v -> {
                throw new IllegalStateException("boom");
            });

            ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> list.write(List.of("one")));

            assertEquals("Failed to write the list into the template", e.getMessage());
            assertInstanceOf(IllegalStateException.class, e.getCause());
            assertEquals("boom", e.getCause().getMessage());
        }
    }

    @Test
    void theSameCallbackFailureInExcelWriterIsAnExcelWriteExceptionToo() {
        // the behaviour the template list now matches
        assertThrows(ExcelWriteException.class, () -> ExcelWriter.<String>create().column("v", v -> v)
                .rowColor(v -> {
                    throw new IllegalStateException("boom");
                }).write(List.of("one")).writeTo(OutputStream.nullOutputStream()));
    }

    @Test
    void aValueTheCellCannotHoldIsStillAnIllegalArgumentException() throws Exception {
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(validTemplate()))) {
            assertThrows(IllegalArgumentException.class, () -> writer.cell("A1", "x".repeat(40_000)));
        }
    }

    @Test
    void aTemplateThatOpensButCannotBeSavedFailsTheWriteWithALibraryException() throws Exception {
        // the document properties part is not XML: the template opens (POI reads it lazily) but cannot be saved again
        byte[] damaged = withPart(validTemplate(), "docProps/core.xml", text -> "this is not xml");

        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(damaged))) {
            writer.cell("A2", "filled in");
            ExcelHandler handler = writer.finish();

            ExcelWriteException e = assertThrows(ExcelWriteException.class, () -> handler.writeTo(OutputStream.nullOutputStream()));

            assertEquals("Failed to write Excel", e.getMessage());
            assertNotNull(e.getCause());
        }
    }

    @Test
    void aValidTemplateIsUnaffected() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(validTemplate()))) {
            writer.cell("A2", "below the title");
            writer.<String>list(3).column("v", v -> v).write(List.of("one", "two"));
            writer.finish().writeTo(out);
        }
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("Title", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("below the title", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("one", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue());
            assertEquals("two", wb.getSheetAt(0).getRow(4).getCell(0).getStringCellValue());
        }
    }
}
