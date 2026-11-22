package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.detect.TabularFileDetector;
import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

class ValueTypeArgumentValidationTest {

    @Test void valueTypesValidateInvalidArguments() {
        assertThrows(IllegalArgumentException.class, () -> new ReadLimits(-2, -1, -1, -1));
        assertThrows(IllegalArgumentException.class, () ->
                new ReadSummary(-1, 0, 0, false, java.time.Duration.ZERO));
        assertThrows(IllegalArgumentException.class, () -> CsvReader.forMap().readReport(csv("A\n1\n"), -1));
        assertThrows(IllegalArgumentException.class, () ->
                TabularFileDetector.detect(java.io.InputStream.nullInputStream()));
    }

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
