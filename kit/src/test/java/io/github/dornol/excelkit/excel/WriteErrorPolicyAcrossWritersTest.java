package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.WriteErrorPolicy;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.csv.CsvWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

/** One policy ({@link WriteErrorPolicy}) with the same two behaviors on every writer. */
class WriteErrorPolicyAcrossWritersTest {
    record Row(String name) {}

    private static final List<Row> ROWS = List.of(new Row("a"), new Row("boom"), new Row("c"));
    private static final Function<Row, Object> FAILS_ON_BOOM = r -> {
        if (r.name().equals("boom")) throw new IllegalStateException("boom");
        return r.name();
    };

    private static void writeCsv(WriteErrorPolicy policy, ByteArrayOutputStream out) {
        var writer = CsvWriter.<Row>create();
        if (policy != null) writer.writeErrorPolicy(policy);
        writer.column("N", FAILS_ON_BOOM, c -> c.nullValue("?")).write(ROWS).writeTo(out);
    }

    private static List<String> excelColumn(byte[] file) throws IOException {
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
            var values = new java.util.ArrayList<String>();
            var sheet = wb.getSheetAt(0);
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                var cell = sheet.getRow(r).getCell(0);
                values.add(cell == null || cell.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK ? "" : cell.getStringCellValue());
            }
            return values;
        }
    }

    @Test
    void everyWriterFailsFastByDefault() {
        assertThrows(CsvWriteException.class, () -> writeCsv(null, new ByteArrayOutputStream()));
        assertThrows(ExcelWriteException.class, () -> ExcelWriter.<Row>create().column("N", FAILS_ON_BOOM).write(ROWS)
                .writeTo(new ByteArrayOutputStream()));
        try (var workbook = ExcelWorkbook.create()) {
            var sheet = workbook.<Row>sheet("S").column("N", FAILS_ON_BOOM);
            assertThrows(ExcelWriteException.class, () -> sheet.write(ROWS));
        }
    }

    @Test
    void failFastAndLenientAreTheSameEnumEverywhere() throws IOException {
        // the same constant is accepted by all of them
        for (WriteErrorPolicy policy : WriteErrorPolicy.values()) {
            CsvWriter.<Row>create().writeErrorPolicy(policy);
            ExcelWriter.<Row>create().writeErrorPolicy(policy);
            try (var workbook = ExcelWorkbook.create()) {
                workbook.<Row>sheet("S").writeErrorPolicy(policy);
            }
        }
        assertEquals(List.of(WriteErrorPolicy.LENIENT, WriteErrorPolicy.FAIL_FAST), List.of(WriteErrorPolicy.values()));
    }

    @Test
    void lenientLeavesAMissingValueAndGoesOnOnEveryWriter() throws IOException {
        var csv = new ByteArrayOutputStream();
        writeCsv(WriteErrorPolicy.LENIENT, csv);
        assertEquals(List.of("N", "a", "?", "c"), List.of(csv.toString(StandardCharsets.UTF_8).replace("﻿", "").split("\r\n")),
                "CSV: the column's nullValue stands where the function failed");

        var excel = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().writeErrorPolicy(WriteErrorPolicy.LENIENT).column("N", FAILS_ON_BOOM, c -> c.nullValue("?")).write(ROWS).writeTo(excel);
        assertEquals(List.of("a", "?", "c"), excelColumn(excel.toByteArray()), "Excel: the same");

        var sheetOut = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<Row>sheet("S").writeErrorPolicy(WriteErrorPolicy.LENIENT).column("N", FAILS_ON_BOOM, c -> c.nullValue("?")).write(ROWS);
            workbook.finish().writeTo(sheetOut);
        }
        assertEquals(List.of("a", "?", "c"), excelColumn(sheetOut.toByteArray()), "workbook sheet: the same");
    }

    @Test
    void theFailureIsOneOfTheLibrarysExceptionsOnEveryWriter() {
        assertInstanceOf(ExcelKitException.class,
                assertThrows(RuntimeException.class, () -> writeCsv(WriteErrorPolicy.FAIL_FAST, new ByteArrayOutputStream())));
        assertInstanceOf(ExcelKitException.class, assertThrows(RuntimeException.class,
                () -> ExcelWriter.<Row>create().writeErrorPolicy(WriteErrorPolicy.FAIL_FAST).column("N", FAILS_ON_BOOM).write(ROWS)
                        .writeTo(new ByteArrayOutputStream())));
    }

    @Test
    void thePolicyMustNotBeNull() {
        assertThrows(NullPointerException.class, () -> CsvWriter.<Row>create().writeErrorPolicy(null));
        assertThrows(NullPointerException.class, () -> ExcelWriter.<Row>create().writeErrorPolicy(null));
    }
}
