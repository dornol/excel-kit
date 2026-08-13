package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.EnumSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReaderInputSourceMatrixTest {
    private enum SourceKind {
        STREAM,
        PATH,
        SOURCE
    }

    private record Person(String name) {
    }

    @ParameterizedTest(name = "CSV supports {0}")
    @EnumSource(SourceKind.class)
    void csvSupportsAllInputSources(SourceKind sourceKind) throws Exception {
        byte[] content = "Name\nAlice\nBob\n".getBytes();
        TrackingInputStream callerStream = new TrackingInputStream(content);
        AtomicReference<TrackingInputStream> sourceStream = new AtomicReference<>();
        Path path = Files.createTempFile("excel-kit-reader", ".csv");
        Files.write(path, content);
        try {
            CsvReader<Person> reader = CsvReader.mapping(row -> new Person(row.get("Name").asString()));
            ReadSummary summary = switch (sourceKind) {
                case STREAM -> reader.readWithSummary(callerStream, result -> {});
                case PATH -> reader.readWithSummary(path, result -> {});
                case SOURCE -> reader.readWithSummary(() -> {
                    TrackingInputStream input = new TrackingInputStream(content);
                    sourceStream.set(input);
                    return input;
                }, result -> {});
            };

            assertSuccessful(summary);
            if (sourceKind == SourceKind.STREAM) assertFalse(callerStream.closed);
            if (sourceKind == SourceKind.SOURCE) assertTrue(sourceStream.get().closed);
        } finally {
            Files.deleteIfExists(path);
        }
    }

    @ParameterizedTest(name = "Excel supports {0}")
    @EnumSource(SourceKind.class)
    void excelSupportsAllInputSources(SourceKind sourceKind) throws Exception {
        byte[] content = excelContent();
        TrackingInputStream callerStream = new TrackingInputStream(content);
        AtomicReference<TrackingInputStream> sourceStream = new AtomicReference<>();
        Path path = Files.createTempFile("excel-kit-reader", ".xlsx");
        Files.write(path, content);
        try {
            ExcelReader<Person> reader = ExcelReader.mapping(row -> new Person(row.get("Name").asString()));
            ReadSummary summary = switch (sourceKind) {
                case STREAM -> reader.readWithSummary(callerStream, result -> {});
                case PATH -> reader.readWithSummary(path, result -> {});
                case SOURCE -> reader.readWithSummary(() -> {
                    TrackingInputStream input = new TrackingInputStream(content);
                    sourceStream.set(input);
                    return input;
                }, result -> {});
            };

            assertSuccessful(summary);
            if (sourceKind == SourceKind.STREAM) assertFalse(callerStream.closed);
            if (sourceKind == SourceKind.SOURCE) assertTrue(sourceStream.get().closed);
        } finally {
            Files.deleteIfExists(path);
        }
    }

    private static void assertSuccessful(ReadSummary summary) {
        assertEquals(2, summary.totalRows());
        assertEquals(2, summary.successRows());
        assertEquals(0, summary.errorRows());
        assertFalse(summary.stoppedEarly());
    }

    private static byte[] excelContent() throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create()
                .column("Name", Person::name)
                .write(List.of(new Person("Alice"), new Person("Bob")))
                .writeTo(output);
        return output.toByteArray();
    }

    private static final class TrackingInputStream extends ByteArrayInputStream {
        private boolean closed;

        private TrackingInputStream(byte[] content) {
            super(content);
        }

        @Override
        public void close() {
            closed = true;
        }
    }
}
