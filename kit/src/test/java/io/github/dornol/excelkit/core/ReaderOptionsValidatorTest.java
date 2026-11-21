package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.NotBlank;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;

class ReaderOptionsValidatorTest {

    public static class Row {
        @NotBlank
        public String name;
    }

    private static final Validator VALIDATOR = Validation.byDefaultProvider().configure()
            .messageInterpolator(new ParameterMessageInterpolator())
            .buildValidatorFactory().getValidator();

    @Test
    void optionsWithoutValidatorKeepTheFactoryValidator() {
        List<Boolean> successes = new ArrayList<>();
        CsvReader.setter(Row::new, VALIDATOR)
                .column("name", (row, cell) -> row.name = cell.asString())
                .options(ReaderOptions.builder().build())
                .read(new ByteArrayInputStream("name\n\n".getBytes(StandardCharsets.UTF_8)),
                        result -> successes.add(result.success()));
        assertEquals(List.of(false), successes);
    }

    @Test
    void optionsWithValidatorOverrideTheFactoryValidator() {
        Validator other = Validation.byDefaultProvider().configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory().getValidator();
        CsvReader<Row> reader = CsvReader.setter(Row::new, VALIDATOR)
                .options(ReaderOptions.builder().validator(other).build());
        assertSame(other, reader.configuredValidator());
    }
}
