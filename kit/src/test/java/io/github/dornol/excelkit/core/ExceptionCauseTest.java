package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;

/** A failure that is converted into another keeps what went wrong as its cause. */
class ExceptionCauseTest {

    @Test
    void aBigDecimalThatCannotBeParsedKeepsTheParseFailure() {
        var error = org.junit.jupiter.api.Assertions.assertThrows(IllegalArgumentException.class,
                () -> new CellData(0, "x1", CellConversionConfig.builder().locale(java.util.Locale.US).build()).asBigDecimal());
        assertInstanceOf(NumberFormatException.class, error.getCause());
    }

    @Test
    void aReadSecurityExceptionCanCarryItsCause() {
        var cause = new IllegalStateException("broken xml");
        var error = new ReadSecurityException(ReadSecurityException.Reason.FORMULA, "could not scan", cause);
        assertSame(cause, error.getCause());
        assertEquals(ReadSecurityException.Reason.FORMULA, error.reason());
        assertEquals("could not scan", error.getMessage());
    }
}
