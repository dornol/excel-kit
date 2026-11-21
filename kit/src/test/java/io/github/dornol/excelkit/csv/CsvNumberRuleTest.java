package io.github.dornol.excelkit.csv;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** NON_NUMERIC quoting and the injection exemption for numbers use one definition of a number. */
class CsvNumberRuleTest {

    private static String field(Object value, CsvQuoting quoting) {
        return field(value, quoting, true);
    }

    private static String field(Object value, CsvQuoting quoting, boolean injectionDefense) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        CsvWriter.<Object>create().bom(false).quoting(quoting).csvInjectionDefense(injectionDefense).lineSeparator("\n")
                .column("v", v -> v).write(List.of(value)).writeTo(out);
        return out.toString(StandardCharsets.UTF_8).split("\n")[1];
    }

    @Test
    void plainDecimalsStayUnquotedUnderNonNumeric() {
        for (String number : List.of("0", "42", "-7", "+7", "3.14", ".5", "5.", "1e5", "1E-5", "-2.5e+10")) {
            assertEquals(number, field(number, CsvQuoting.NON_NUMERIC, false), number);
        }
    }

    @Test
    void thingsThatOnlyLookLikeNumbersAreQuoted() {
        for (String text : List.of(".", "-", "+", "1e", "1e+", "1.2.3", "12a", "١٢", "1,5", " 1")) {
            assertEquals("\"" + text + "\"", field(text, CsvQuoting.NON_NUMERIC, false), text);
        }
    }

    @Test
    void aNegativeNumberKeepsItsSignButNegativeTextIsPrefixed() {
        assertEquals("-5", field(-5, CsvQuoting.MINIMAL));
        assertEquals("-1.5E-7", field(-1.5E-7, CsvQuoting.MINIMAL));
        assertEquals("'-5", field("-5", CsvQuoting.MINIMAL));
    }
}
