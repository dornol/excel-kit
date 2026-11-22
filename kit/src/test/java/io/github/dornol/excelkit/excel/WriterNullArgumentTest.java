package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.RowFunction;
import io.github.dornol.excelkit.csv.CsvWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.function.Executable;

import java.util.function.Function;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** A null callback, function or password is rejected where it is passed, not when the first row is written. */
class WriterNullArgumentTest {

    private static void npe(String argument, Executable call) {
        NullPointerException error = assertThrows(NullPointerException.class, call, argument);
        assertTrue(error.getMessage() != null && error.getMessage().contains(argument), error.getMessage());
    }

    private static void iae(Executable call) {
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class, call);
        assertTrue(error.getMessage().contains("null or blank"), error.getMessage());
    }

    @Test
    void sharedSheetSettingsRejectNullCallbacks() {
        ExcelWriter<String> writer = ExcelWriter.create();
        npe("beforeHeaderWriter", () -> writer.beforeHeader(null));
        npe("afterDataWriter", () -> writer.afterData(null));
        npe("rowColorFunction", () -> writer.rowColor(null));
        npe("predicate", () -> writer.rowStyle(null, style -> { }));
        npe("callback", () -> writer.onProgress(10, null));
        npe("sheetNameFunction", () -> writer.sheetName((Function<Integer, String>) null));
        npe("comment", () -> writer.groupComment((ExcelCellComment) null, "group"));
    }

    @Test
    void excelWriterRejectsNullFunctionsAndCallbacks() {
        ExcelWriter<String> writer = ExcelWriter.create();
        npe("afterAllWriter", () -> writer.afterAll(null));
        npe("name", () -> writer.sheetName((String) null));
        npe("fontName", () -> writer.headerFontName(null));
        npe("function", () -> writer.column("c", (Function<String, Object>) null));
        npe("function", () -> writer.column("c", (RowFunction<String, Object>) null));
        npe("function", () -> writer.column("c", (Function<String, Object>) null, c -> { }));
        npe("stream", () -> writer.column("c", s -> s).write((Stream<String>) null));
        npe("consumer", () -> writer.column("c", s -> s).write(Stream.of("a"), null));
    }

    @Test
    void protectionPasswordsRejectNullAndBlank() {
        ExcelWriter<String> writer = ExcelWriter.create();
        iae(() -> writer.protectSheet(null));
        iae(() -> writer.protectSheet(" "));
        iae(() -> writer.protectWorkbook(null));
        iae(() -> writer.protectWorkbook(""));
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            iae(() -> workbook.protectWorkbook(null));
        }
    }

    @Test
    void sheetAndTemplateWritersRejectNullFunctions() {
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            ExcelSheetWriter<String> sheet = workbook.sheet("S");
            npe("function", () -> sheet.column("c", (Function<String, Object>) null));
            npe("function", () -> sheet.column("c", (RowFunction<String, Object>) null));
            npe("function", () -> sheet.column("c", (Function<String, Object>) null, c -> { }));
        }
    }

    @Test
    void csvWriterRejectsNullSettings() {
        CsvWriter<String> writer = CsvWriter.create();
        npe("afterDataWriter", () -> writer.afterData(null));
        npe("function", () -> writer.column("c", (Function<String, Object>) null));
        npe("function", () -> writer.column("c", (RowFunction<String, Object>) null));
        npe("callback", () -> writer.onProgress(10, null));
        npe("quoting", () -> writer.quoting(null));
        npe("charset", () -> writer.charset(null));
        npe("stream", () -> writer.column("c", s -> s).write((Stream<String>) null));
    }

    @Test
    void anOptionalColumnConfigurerMayStillBeNull() {
        assertDoesNotThrow(() -> ExcelWriter.<String>create().column("c", s -> s, null));
    }
}
