package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class HeaderPolicyNormalizationTest {

    @Test void headerPolicyMatchesNormalizedNames() {
        var values = new java.util.ArrayList<Person>();
        CsvReader.<Person>mapping(row -> new Person(row.get("user name").asString()))
                .headerPolicy(HeaderPolicy.NORMALIZED_CASE_INSENSITIVE)
                .read(csv(" User   Name \nA\n"), result -> values.add(result.data()));
        assertEquals(List.of(new Person("A")), values);
    }

    record Person(String name) {}

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
