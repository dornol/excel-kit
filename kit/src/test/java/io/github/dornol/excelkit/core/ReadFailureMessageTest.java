package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReadException;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A read that fails because of what is in the file says so in its message, not only in the cause
 * chain: an API that reports just {@code getMessage()} would otherwise show "Failed to read CSV".
 */
class ReadFailureMessageTest {

    static class Row {
        String name;
        String price;
    }

    private static CsvReader<Row> csvReader() {
        return CsvReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .column("Price", (row, cell) -> row.price = cell.asString());
    }

    private static ExcelReader<Row> excelReader() {
        return ExcelReader.setter(Row::new)
                .column("Name", (row, cell) -> row.name = cell.asString())
                .column("Price", (row, cell) -> row.price = cell.asString());
    }

    private static CsvReadException csvFailure(String text) {
        return assertThrows(CsvReadException.class, () -> csvReader()
                .read(new ByteArrayInputStream(text.getBytes(StandardCharsets.UTF_8)), result -> { }));
    }

    private static ExcelReadException excelFailure(String[]... rows) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            var sheet = workbook.createSheet("Sheet1");
            for (int r = 0; r < rows.length; r++) {
                var row = sheet.createRow(r);
                for (int c = 0; c < rows[r].length; c++) {
                    row.createCell(c).setCellValue(rows[r][c]);
                }
            }
            workbook.write(out);
            return assertThrows(ExcelReadException.class,
                    () -> excelReader().read(new ByteArrayInputStream(out.toByteArray()), result -> { }));
        }
    }

    @Test
    void csvNamesTheMissingHeaderAndTheOnesThatWereFound() {
        String message = csvFailure("Name,Cost\nA,1\n").getMessage();

        assertTrue(message.startsWith("Failed to read CSV: "), message);
        assertTrue(message.contains("Header aliases [Price] not found"), message);
        assertTrue(message.contains("Available headers: [Name, Cost]"), message);
    }

    @Test
    void csvNamesADuplicatedHeader() {
        String message = csvFailure("Name,Price,Name\nA,1,B\n").getMessage();

        assertTrue(message.startsWith("Failed to read CSV: "), message);
        assertTrue(message.contains("Duplicate header 'Name'"), message);
    }

    @Test
    void csvSaysWhenThereIsNoHeaderRow() {
        String message = csvFailure("").getMessage();

        assertTrue(message.contains("empty or missing header row"), message);
    }

    @Test
    void excelNamesTheMissingHeaderAndTheOnesThatWereFound() throws IOException {
        String message = excelFailure(new String[]{"Name", "Cost"}, new String[]{"A", "1"}).getMessage();

        assertTrue(message.startsWith("Failed to read excel: "), message);
        assertTrue(message.contains("Header aliases [Price] not found"), message);
        assertTrue(message.contains("Available headers: [Name, Cost]"), message);
    }

    @Test
    void excelNamesADuplicatedHeader() throws IOException {
        String message = excelFailure(new String[]{"Name", "Price", "Name"}, new String[]{"A", "1", "B"}).getMessage();

        assertTrue(message.startsWith("Failed to read excel: "), message);
        assertTrue(message.contains("Duplicate header 'Name'"), message);
    }

    @Test
    void theOriginalCauseIsStillAttached() {
        CsvReadException error = csvFailure("Name,Cost\nA,1\n");

        assertTrue(error.getCause() instanceof ExcelKitException);
        assertTrue(error.getMessage().endsWith(error.getCause().getMessage()));
    }

    // ------------------------------------------------------------ the helper itself

    @Test
    void aLibraryCheckAddsItsMessage() {
        assertEquals("Failed to read X: header missing",
                AbstractReadHandler.readFailureMessage("Failed to read X", new ExcelKitException("header missing")));
    }

    @Test
    void otherCausesKeepTheGenericMessage() {
        assertEquals("Failed to read X",
                AbstractReadHandler.readFailureMessage("Failed to read X", new IllegalStateException("parser state")));
        assertEquals("Failed to read X",
                AbstractReadHandler.readFailureMessage("Failed to read X", new IOException("disk")));
    }

    @Test
    void aLibraryCheckWithoutAMessageKeepsTheGenericMessage() {
        assertEquals("Failed to read X",
                AbstractReadHandler.readFailureMessage("Failed to read X", new ExcelKitException((String) null)));
        assertEquals("Failed to read X",
                AbstractReadHandler.readFailureMessage("Failed to read X", new ExcelKitException("  ")));
    }
}
