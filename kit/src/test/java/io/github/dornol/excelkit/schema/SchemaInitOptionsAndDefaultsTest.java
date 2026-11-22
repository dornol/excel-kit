package io.github.dornol.excelkit.schema;

import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

class SchemaInitOptionsAndDefaultsTest {

    @Test
    void schemaExcelWriter_shouldAcceptInitOptions() {
        ExcelKitSchema<String> schema = ExcelKitSchema.<String>builder()
                .column("Name", value -> value, (value, cell) -> {})
                .build();

        ExcelWriter<String> writer = schema.excelWriter(opts -> opts.rowAccessWindowSize(10));

        writer.write(Stream.of("Alice")).writeTo(java.io.OutputStream.nullOutputStream());
    }

    @Test
    void schemaDefaults_shouldApplyToReaders() {
        ExcelKitSchema<String> schema = ExcelKitSchema.<String>builder()
                .column("Name", value -> value, (value, cell) -> {})
                .strictHeaders()
                .maxRows(1)
                .build();
        byte[] csv = "Name\nAlice\nBob\n".getBytes(StandardCharsets.UTF_8);
        List<String> names = new ArrayList<>();

        schema.csvReader(() -> "")
                .read(new ByteArrayInputStream(csv), result -> names.add(result.data()));

        assertEquals(1, names.size());
    }
}
