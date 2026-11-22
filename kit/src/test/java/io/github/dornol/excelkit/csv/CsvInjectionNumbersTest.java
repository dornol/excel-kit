package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** The injection defense treats a number as a number and a text that looks like a formula as text. */
class CsvInjectionNumbersTest {

    @Test
    void negativeNumbersKeepTheirSignWhileTextIsDefended() {
        var out = new ByteArrayOutputStream();
        CsvWriter.<Object>create().bom(false)
                .column("v", o -> o)
                .write(Stream.of(-1, -2.5, -1L, new java.math.BigDecimal("-3.25"), "-1", "+1", "=1+1"))
                .writeTo(out);

        assertEquals("v\r\n-1\r\n-2.5\r\n-1\r\n-3.25\r\n'-1\r\n'+1\r\n'=1+1\r\n", out.toString(java.nio.charset.StandardCharsets.UTF_8));
    }
}
