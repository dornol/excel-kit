package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class CsvReaderColumnBuilderTest {

    // -----------------------------------------------------------------------
    // 1. CsvReader — skipColumn, skipColumns, columnAt
    // -----------------------------------------------------------------------

    @Test
    void csvBuilder_skipColumn_shouldSkipOneColumnAndContinue() {
        String csv = "A,B,C\na1,b1,c1\na2,b2,c2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .skipColumn()  // builder chain: skipColumn returns CsvReader, then column starts new builder
                .column((r, cell) -> r.third = cell.asString())
                .read(is, result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertEquals("a1", results.get(0).first);
        assertNull(results.get(0).second);
        assertEquals("c1", results.get(0).third);
        assertEquals("a2", results.get(1).first);
        assertEquals("c2", results.get(1).third);
    }

    @Test
    void csvBuilder_skipColumns_shouldSkipMultipleColumnsAndContinue() {
        String csv = "A,B,C,D\na1,b1,c1,d1\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .skipColumns(2)  // skip B and C
                .column((r, cell) -> r.fourth = cell.asString())
                .read(is, result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("d1", results.get(0).fourth);
    }

    @Test
    void csvBuilder_columnAt_shouldReadByIndex() {
        String csv = "A,B,C,D\na1,b1,c1,d1\na2,b2,c2,d2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .columnAt(3, (r, cell) -> r.fourth = cell.asString())
                .read(is, result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("d1", results.get(0).fourth);
        assertEquals("a2", results.get(1).first);
        assertEquals("d2", results.get(1).fourth);
    }

    @Test
    void csvBuilder_columnAt_chainMultiple() {
        String csv = "A,B,C,D,E\na1,b1,c1,d1,e1\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .columnAt(2, (r, cell) -> r.third = cell.asString())
                .columnAt(4, (r, cell) -> r.fifth = cell.asString())
                .read(is, result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("c1", results.get(0).third);
        assertEquals("e1", results.get(0).fifth);
    }

    @Test
    void csvBuilder_column_namedChain() {
        String csv = "Name,Age,City\nAlice,30,Seoul\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column("City", (r, cell) -> r.third = cell.asString())
                .column("Name", (r, cell) -> r.first = cell.asString())
                .read(is, result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("Seoul", results.get(0).third);
        assertEquals("Alice", results.get(0).first);
    }

    // -----------------------------------------------------------------------
    // Test data classes
    // -----------------------------------------------------------------------

    public static class CsvRow {
        String first;
        String second;
        String third;
        String fourth;
        String fifth;
    }
}
