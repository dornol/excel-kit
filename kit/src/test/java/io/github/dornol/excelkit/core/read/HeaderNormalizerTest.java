package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class HeaderNormalizerTest {

    @Test
    void headerNormalizerAppliesToHeadersAndLookups() {
        List<Person> people = new ArrayList<>();
        CsvReader.<Person>mapping(row -> new Person(row.get("name").asString()))
                .headerNormalizer(value -> value.trim().toLowerCase(java.util.Locale.ROOT))
                .read(input(" Name \nAlice\n"), result -> people.add(result.data()));
        assertEquals(List.of(new Person("Alice")), people);
    }

    @Test
    void normalizedDuplicateHeadersHonorPolicy() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .headerNormalizer(value -> value.trim().toLowerCase(java.util.Locale.ROOT))
                .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                .read(input("Name, name \nA,B\n"), result -> {}));
    }

    record Person(String name) {}

    private static ByteArrayInputStream input(String csv) {
        return new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));
    }
}
