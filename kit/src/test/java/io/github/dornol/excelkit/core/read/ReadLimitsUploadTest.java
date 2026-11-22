package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReadLimitsUploadTest {

    static class Row {
        String name;
    }

    @Test
    void presetBoundsEveryDimension() {
        ReadLimits upload = ReadLimits.UPLOAD;
        assertEquals(50L * 1024 * 1024, upload.maxInputBytes());
        assertEquals(50, upload.maxSheets());
        assertEquals(1_000, upload.maxColumns());
        assertEquals(32_767, upload.maxCellCharacters());
        assertEquals(10_000_000L, upload.maxTotalCells());
    }

    @Test
    void readerRejectsACellLongerThanTheExcelLimit() {
        String csv = "name\n" + "x".repeat(32_768) + "\n";
        CsvReader<Row> reader = CsvReader.setter(Row::new)
                .column("name", (row, cell) -> row.name = cell.asString())
                .limits(ReadLimits.UPLOAD);
        ReadLimitExceededException e = assertThrows(ReadLimitExceededException.class,
                () -> reader.readAll(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8))));
        assertEquals(ReadLimitExceededException.Limit.CELL_CHARACTERS, e.limit());
    }

    @Test
    void readerAcceptsAnOrdinaryFile() {
        CsvReader<Row> reader = CsvReader.setter(Row::new)
                .column("name", (row, cell) -> row.name = cell.asString())
                .limits(ReadLimits.UPLOAD);
        var rows = reader.readAll(new ByteArrayInputStream("name\nann\n".getBytes(StandardCharsets.UTF_8)));
        assertTrue(rows.size() == 1 && "ann".equals(rows.get(0).name));
    }
}
