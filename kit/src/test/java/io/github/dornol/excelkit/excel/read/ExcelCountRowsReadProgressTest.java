package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadProgress;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.*;

/** {@code countRows()} gives the total to the read progress callback, with or without the older {@code onProgress}. */
class ExcelCountRowsReadProgressTest {
    record Item(String name, int value) {}

    private static byte[] file(int rows) {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name).column("Value", Item::value, c -> c.type(ExcelDataType.INTEGER))
                .write(IntStream.range(0, rows).mapToObj(i -> new Item("n" + i, i))).writeTo(out);
        return out.toByteArray();
    }

    private static ExcelReader<Item> reader() {
        return ExcelReader.<Item>mapping(row -> new Item(row.get("Name").asString(), row.get("Value").asInt()));
    }

    @Test
    void theReadProgressCallbackKnowsTheTotalWithoutTheOlderCallback() {
        List<ReadProgress> seen = new ArrayList<>();
        reader().countRows().onReadProgress(2, seen::add).read(new ByteArrayInputStream(file(5)), r -> {});
        assertFalse(seen.isEmpty());
        for (ReadProgress progress : seen) {
            assertEquals(5, progress.totalRows(), "processed " + progress.processedRows());
        }
    }

    @Test
    void withoutCountRowsTheTotalIsUnknown() {
        List<ReadProgress> seen = new ArrayList<>();
        reader().onReadProgress(2, seen::add).read(new ByteArrayInputStream(file(5)), r -> {});
        assertFalse(seen.isEmpty());
        assertTrue(seen.stream().noneMatch(p -> p.totalRows() >= 0 && !p.completed()), "no total while reading: " + seen);
    }

    @Test
    void bothCallbacksTogetherAgree() {
        List<ReadProgress> seen = new ArrayList<>();
        List<Long> totals = new ArrayList<>();
        reader().countRows().onReadProgress(1, seen::add).onProgress(1, (count, cursor) -> totals.add(cursor.getTotalRows()))
                .read(new ByteArrayInputStream(file(3)), r -> {});
        assertEquals(List.of(3L, 3L, 3L), totals);
        assertTrue(seen.stream().allMatch(p -> p.totalRows() == 3), seen.toString());
    }
}
