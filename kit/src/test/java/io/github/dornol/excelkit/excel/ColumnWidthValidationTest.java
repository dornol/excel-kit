package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

/** A width Excel cannot hold is refused where it is set, and says which value. */
class ColumnWidthValidationTest {

    private static void write(Consumer<ColumnConfig<String>> config) {
        ExcelWriter.<String>create().column("A", s -> s, config)
                .write(List.of("x").stream()).writeTo(OutputStream.nullOutputStream());
    }

    @Test
    void aNegativeOrTooWideWidthIsRefusedWhereItIsSet() {
        var tooWide = assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create()
                .column("A", s -> s, c -> c.width(255 * 256 + 1)));
        assertTrue(tooWide.getMessage().contains("width") && tooWide.getMessage().contains("65281"), tooWide.getMessage());
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().column("A", s -> s, c -> c.width(-1)));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().column("A", s -> s, c -> c.minWidth(-1)));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().column("A", s -> s, c -> c.maxWidth(-1)));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create().column("A", s -> s, c -> c.maxWidth(70000)));
    }

    @Test
    void theLimitsThemselvesAreAccepted() {
        assertDoesNotThrow(() -> write(c -> c.width(255 * 256)));
        assertDoesNotThrow(() -> write(c -> c.minWidth(0).maxWidth(0)));
    }

    @Test
    void aMinimumAboveTheMaximumIsRefused() {
        var e = assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<String>create()
                .column("A", s -> s, c -> c.minWidth(5000).maxWidth(3000)));
        assertTrue(e.getMessage().contains("minWidth") && e.getMessage().contains("maxWidth"), e.getMessage());
    }

    @Test
    void schemaColumnRefusesANullWriteConfigurer() {
        assertThrows(NullPointerException.class, () -> ExcelKitSchema.<String>builder()
                .column("A", s -> s, (s, cell) -> { }, null));
        assertThrows(NullPointerException.class, () -> ExcelKitSchema.<String>builder()
                .requiredColumn("A", s -> s, (s, cell) -> { }, null));
    }
}
