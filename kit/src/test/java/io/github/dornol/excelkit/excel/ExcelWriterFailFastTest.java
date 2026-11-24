package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.WriteErrorPolicy;
import org.junit.jupiter.api.Test;

import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelWriterFailFastTest {

    @Test
    void excelWriter_failFast_shouldThrowWhenValueExtractorFails() {
        ExcelWriter<String> writer = ExcelWriter.<String>create()
                .writeErrorPolicy(WriteErrorPolicy.FAIL_FAST)
                .column("Name", value -> {
                    throw new IllegalStateException("boom");
                });

        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("x")));
    }

    @Test
    void excelWriter_failFast_shouldThrowWhenCellWriteFails() {
        ExcelWriter<String> writer = ExcelWriter.<String>create()
                .writeErrorPolicy(WriteErrorPolicy.FAIL_FAST)
                .column("Age", value -> "not-number", c -> c.type(ExcelDataType.INTEGER));

        assertThrows(ExcelWriteException.class, () -> writer.write(Stream.of("x")));
    }
}
