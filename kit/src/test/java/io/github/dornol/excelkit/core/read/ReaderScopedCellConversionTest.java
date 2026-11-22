package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ReaderScopedCellConversionTest {

    @Test
    void readerScopedCellConversion_shouldParseCustomDateWithoutGlobalMutation() {
        byte[] csv = "Date\n25.12.2024\n".getBytes(StandardCharsets.UTF_8);
        List<LocalDate> dates = new ArrayList<>();

        CsvReader.<LocalDate>mapping(row -> row.get("Date").asLocalDate())
                .cellConversion(c -> c.addDateFormat("dd.MM.yyyy"))
                .read(new ByteArrayInputStream(csv), result -> dates.add(result.data()));

        assertEquals(List.of(LocalDate.of(2024, 12, 25)), dates);
    }

    @Test
    void readerScopedCellConversion_shouldUseConfiguredLocale() {
        byte[] csv = "Amount\n1.234,5\n".getBytes(StandardCharsets.UTF_8);
        List<Double> values = new ArrayList<>();

        CsvReader.<Double>mapping(row -> row.get("Amount").asDouble())
                .delimiter(';')
                .cellConversion(c -> c.locale(Locale.GERMANY))
                .read(new ByteArrayInputStream(csv), result -> values.add(result.data()));

        assertEquals(1234.5d, values.get(0), 0.0001);
    }
}
