package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadLimitExceededException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import java.io.InputStream;
import java.util.Iterator;

/** Sheet counting, limit enforcement, and selected-sheet navigation. */
final class ExcelSheetNavigator {
    private ExcelSheetNavigator() {}
    @FunctionalInterface interface SheetConsumer { void accept(InputStream sheet) throws Exception; }

    static void enforceLimit(XSSFReader reader, int maximum) throws Exception {
        if (maximum < 0) return;
        int count = 0;
        Iterator<InputStream> iterator = reader.getSheetsData();
        while (iterator.hasNext()) {
            try (InputStream ignored = iterator.next()) { count++; }
            if (count > maximum) throw new ReadLimitExceededException(
                    ReadLimitExceededException.Limit.SHEETS, maximum, count);
        }
    }

    static void consume(XSSFReader reader, int selected, SheetConsumer consumer) throws Exception {
        Iterator<InputStream> iterator = reader.getSheetsData();
        int index = 0;
        while (iterator.hasNext()) {
            try (InputStream sheet = iterator.next()) {
                if (index == selected) { consumer.accept(sheet); return; }
            }
            index++;
        }
        throw new ExcelReadException("Sheet index " + selected + " not found. File has " + index + " sheet(s).");
    }
}
