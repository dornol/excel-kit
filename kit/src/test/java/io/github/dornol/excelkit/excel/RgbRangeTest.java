package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.function.Executable;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** Every RGB setter rejects a component outside 0-255 instead of wrapping it into another color. */
class RgbRangeTest {

    private static void rejected(Executable call) {
        IllegalArgumentException error = assertThrows(IllegalArgumentException.class, call);
        org.junit.jupiter.api.Assertions.assertTrue(error.getMessage().contains("0 and 255"), error.getMessage());
    }

    @Test
    void columnColorsRejectOutOfRangeComponents() {
        ColumnConfig<String> column = new ColumnConfig<>();
        rejected(() -> column.backgroundColor(300, 0, 0));
        rejected(() -> column.fontColor(0, -1, 0));
        rejected(() -> column.headerFontColor(0, 0, 256));
        rejected(() -> column.headerBackgroundColor(-5, 0, 0));
    }

    @Test
    void tabColorRejectsOutOfRangeComponents() {
        ExcelWriter<String> writer = ExcelWriter.create();
        rejected(() -> writer.tabColor(300, 0, 0));
    }

    @Test
    void richTextColorRejectsOutOfRangeComponents() {
        rejected(() -> new ExcelRichText().styled("x", style -> style.color(0, 256, 0)));
    }

    @Test
    void theBoundariesAreAccepted() {
        ColumnConfig<String> column = new ColumnConfig<>();
        assertDoesNotThrow(() -> column.backgroundColor(0, 0, 0).fontColor(255, 255, 255));
        assertDoesNotThrow(() -> ExcelWriter.<String>create().tabColor(0, 255, 0));
    }
}
