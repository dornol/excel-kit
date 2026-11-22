package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ReaderColumnSnapshotTest {

    @Test
    void executionUsesColumnSnapshot() {
        CsvReader<MutablePerson> reader = new CsvReader<>(MutablePerson::new)
                .column("Name", (person, cell) -> person.name = cell.asString());
        List<ReadResult<MutablePerson>> results = new ArrayList<>();
        reader.read(input("Name,Later\nA,1\nB,2\n"), result -> {
            results.add(result);
            if (results.size() == 1) {
                reader.column("Later", (person, cell) -> person.later = cell.asString());
            }
        });
        assertEquals(2, results.size());
        assertNull(results.get(1).data().later);
    }

    static final class MutablePerson { String name; String later; }

    private static ByteArrayInputStream input(String csv) {
        return new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
    }
}
