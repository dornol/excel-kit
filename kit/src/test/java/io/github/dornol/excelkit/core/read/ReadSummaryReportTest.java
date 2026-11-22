package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

class ReadSummaryReportTest {

    @Test void summaryAndBoundedReportCountRows() {
        ReadSummary summary = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()))
                .readWithSummary(csv("Name\nA\nB\n"), result -> {});
        assertEquals(2, summary.totalRows());
        assertEquals(2, summary.successRows());
        assertEquals(0, summary.errorRows());
        assertFalse(summary.stoppedEarly());
        assertFalse(summary.duration().isNegative());

        ReadReport report = CsvReader.<Person>mapping(row -> { throw new IllegalArgumentException("bad"); })
                .readReport(csv("Name\nA\nB\n"), 1);
        assertEquals(2, report.summary().errorRows());
        assertEquals(1, report.errors().size());
        assertTrue(report.errorsTruncated());
    }

    @Test void summaryReportAndWhileSupportPathAndSource() throws Exception {
        var path = Files.createTempFile("excel-kit-summary", ".csv");
        try {
            Files.writeString(path, "Name\nA\nB\n");
            var reader = CsvReader.<Person>mapping(row -> new Person(row.get("Name").asString()));
            assertEquals(2, reader.readWithSummary(path, result -> {}).totalRows());
            assertEquals(0, reader.readReport((InputStreamSource) () -> Files.newInputStream(path), 1)
                    .summary().errorRows());
            ReadSummary stopped = reader.readWhile(path, result -> false);
            assertEquals(1, stopped.totalRows());
            assertTrue(stopped.stoppedEarly());
        } finally { Files.deleteIfExists(path); }
    }

    record Person(String name) {}

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
