package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.csv.CsvReader;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.lang.reflect.RecordComponent;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.function.Consumer;
import java.util.function.UnaryOperator;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A common reader setting can be given in two ways: by chaining on the reader, or through a
 * reusable {@link ReaderOptions}. Both must produce exactly the same configuration.
 * <p>
 * Every component of {@link ReadOptions} (plus the validator and the header row index, which are not
 * part of that snapshot) needs a case here, so adding a setting to one path and forgetting the other
 * fails this test instead of silently configuring only one of them.
 */
class ReaderSettingsEquivalenceTest {

    static class Row {
    }

    private static final Validator VALIDATOR = Validation.byDefaultProvider().configure()
            .messageInterpolator(new ParameterMessageInterpolator())
            .buildValidatorFactory()
            .getValidator();
    private static final CellConversionConfig CONVERSION =
            CellConversionConfig.builder().locale(Locale.GERMANY).build();
    private static final UnaryOperator<String> NORMALIZER = String::toUpperCase;
    private static final ReadLimits LIMITS = new ReadLimits(1_000, 2, 10, 100);
    private static final CancellationToken TOKEN = () -> true;
    private static final ProgressCallback PROGRESS = (rows, cursor) -> { };
    private static final ReadProgressCallback READ_PROGRESS = progress -> { };

    /** One setting: how to give it by chaining, how to give it through options, and what it changes. */
    record Case(String name, Set<String> changes, Consumer<CsvReader<Row>> chained,
                Consumer<ReaderOptions.Builder> viaOptions) {
        @Override
        public String toString() {
            return name;
        }
    }

    private static Case setting(String name, Set<String> changes, Consumer<CsvReader<Row>> chained,
                                Consumer<ReaderOptions.Builder> viaOptions) {
        return new Case(name, changes, chained, viaOptions);
    }

    static Stream<Case> settings() {
        return Stream.of(
                setting("validator", Set.of("validator"),
                        r -> r.validator(VALIDATOR), o -> o.validator(VALIDATOR)),
                setting("headerRowIndex", Set.of("headerRowIndex"),
                        r -> r.headerRowIndex(3), o -> o.headerRowIndex(3)),
                setting("onProgress", Set.of("progressInterval", "progressCallback"),
                        r -> r.onProgress(5, PROGRESS), o -> o.onProgress(5, PROGRESS)),
                setting("strictHeaders()", Set.of("strictHeaders"),
                        CsvReader::strictHeaders, ReaderOptions.Builder::strictHeaders),
                setting("strictHeaders(true)", Set.of("strictHeaders"),
                        r -> r.strictHeaders(true), o -> o.strictHeaders(true)),
                setting("duplicateHeaderPolicy", Set.of("duplicateHeaderPolicy"),
                        r -> r.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST),
                        o -> o.duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)),
                setting("cellConversion", Set.of("cellConversionConfig"),
                        r -> r.cellConversion(CONVERSION), o -> o.cellConversion(CONVERSION)),
                setting("maxRows", Set.of("maxRows"),
                        r -> r.maxRows(7), o -> o.maxRows(7)),
                setting("skipBlankRows()", Set.of("skipBlankRows"),
                        CsvReader::skipBlankRows, ReaderOptions.Builder::skipBlankRows),
                setting("skipBlankRows(true)", Set.of("skipBlankRows"),
                        r -> r.skipBlankRows(true), o -> o.skipBlankRows(true)),
                setting("stopAtBlankRows", Set.of("stopAtBlankRows"),
                        r -> r.stopAtBlankRows(2), o -> o.stopAtBlankRows(2)),
                setting("maxErrors", Set.of("maxErrors"),
                        r -> r.maxErrors(4), o -> o.maxErrors(4)),
                setting("headerNormalizer", Set.of("headerNormalizer"),
                        r -> r.headerNormalizer(NORMALIZER), o -> o.headerNormalizer(NORMALIZER)),
                setting("limits", Set.of("limits"),
                        r -> r.limits(LIMITS), o -> o.limits(LIMITS)),
                setting("cancellationToken", Set.of("cancellationToken"),
                        r -> r.cancellationToken(TOKEN), o -> o.cancellationToken(TOKEN)),
                setting("onReadProgress", Set.of("readProgressInterval", "readProgressCallback"),
                        r -> r.onReadProgress(6, READ_PROGRESS), o -> o.onReadProgress(6, READ_PROGRESS)),
                setting("securityPolicy", Set.of("securityPolicy"),
                        r -> r.securityPolicy(ReadSecurityPolicy.STRICT),
                        o -> o.securityPolicy(ReadSecurityPolicy.STRICT)),
                setting("uniqueKey", Set.of("uniqueKey"),
                        r -> r.uniqueKey("Dept", "ID"), o -> o.uniqueKey("Dept", "ID"))
        );
    }

    private static CsvReader<Row> reader() {
        return CsvReader.setter(Row::new);
    }

    @ParameterizedTest
    @MethodSource("settings")
    void chainingAndOptionsConfigureTheReaderIdentically(Case setting) {
        CsvReader<Row> chained = reader();
        setting.chained().accept(chained);

        ReaderOptions.Builder builder = ReaderOptions.builder();
        setting.viaOptions().accept(builder);
        CsvReader<Row> viaOptions = reader().options(builder.build());

        assertEquals(chained.snapshotReadOptions(), viaOptions.snapshotReadOptions());
        assertEquals(chained.configuredValidator(), viaOptions.configuredValidator());
        assertEquals(chained.configuredHeaderRowIndex(), viaOptions.configuredHeaderRowIndex());
    }

    @ParameterizedTest
    @MethodSource("settings")
    void everyCaseActuallyChangesTheConfiguration(Case setting) {
        CsvReader<Row> untouched = reader();
        CsvReader<Row> chained = reader();
        setting.chained().accept(chained);

        boolean changed = !untouched.snapshotReadOptions().equals(chained.snapshotReadOptions())
                || untouched.configuredValidator() != chained.configuredValidator()
                || untouched.configuredHeaderRowIndex() != chained.configuredHeaderRowIndex();
        assertTrue(changed, setting.name() + " did not change anything, so the comparison proves nothing");
    }

    @Test
    void everyReadOptionsComponentHasACase() {
        Set<String> required = new LinkedHashSet<>();
        for (RecordComponent component : ReadOptions.class.getRecordComponents()) {
            required.add(component.getName());
        }
        required.add("validator");
        required.add("headerRowIndex");

        Set<String> covered = new LinkedHashSet<>();
        settings().forEach(setting -> covered.addAll(setting.changes()));

        List<String> missing = new ArrayList<>(required);
        missing.removeAll(covered);
        assertTrue(missing.isEmpty(),
                "settings without a case in ReaderSettingsEquivalenceTest: " + missing);
    }

    @Test
    void optionsReplaceTheCommonSettingsInsteadOfMergingWithThem() {
        CsvReader<Row> reader = reader().maxRows(10).skipBlankRows().headerRowIndex(2);

        reader.options(ReaderOptions.builder().maxErrors(3).build());

        ReadOptions snapshot = reader.snapshotReadOptions();
        assertEquals(-1, snapshot.maxRows());
        assertEquals(false, snapshot.skipBlankRows());
        assertEquals(3, snapshot.maxErrors());
        assertEquals(0, reader.configuredHeaderRowIndex());
    }

    @Test
    void aReaderNeverChangesTheOptionsItWasGiven() {
        ReaderOptions shared = ReaderOptions.builder().maxRows(5).build();

        CsvReader<Row> first = reader().options(shared);
        first.maxRows(99).skipBlankRows();
        CsvReader<Row> second = reader().options(shared);

        assertEquals(5, second.snapshotReadOptions().maxRows());
        assertEquals(false, second.snapshotReadOptions().skipBlankRows());
        assertNotEquals(first.snapshotReadOptions(), second.snapshotReadOptions());
    }

    @Test
    void validationIsTheSameOnBothPaths() {
        assertThrowsBoth(IllegalArgumentException.class, r -> r.maxRows(-1), o -> o.maxRows(-1));
        assertThrowsBoth(IllegalArgumentException.class, r -> r.maxErrors(-1), o -> o.maxErrors(-1));
        assertThrowsBoth(IllegalArgumentException.class, r -> r.stopAtBlankRows(-1), o -> o.stopAtBlankRows(-1));
        assertThrowsBoth(IllegalArgumentException.class,
                r -> r.onProgress(0, PROGRESS), o -> o.onProgress(0, PROGRESS));
        assertThrowsBoth(IllegalArgumentException.class,
                r -> r.onReadProgress(0, READ_PROGRESS), o -> o.onReadProgress(0, READ_PROGRESS));
        assertThrowsBoth(NullPointerException.class, r -> r.limits(null), o -> o.limits(null));
        assertThrowsBoth(NullPointerException.class,
                r -> r.cancellationToken(null), o -> o.cancellationToken(null));
        assertThrowsBoth(NullPointerException.class, r -> r.securityPolicy(null), o -> o.securityPolicy(null));
        assertThrowsBoth(NullPointerException.class,
                r -> r.duplicateHeaderPolicy(null), o -> o.duplicateHeaderPolicy(null));
        assertThrowsBoth(NullPointerException.class,
                r -> r.headerNormalizer(null), o -> o.headerNormalizer(null));
    }

    private static <E extends Throwable> void assertThrowsBoth(Class<E> expected, Consumer<CsvReader<Row>> chained,
                                                               Consumer<ReaderOptions.Builder> viaOptions) {
        org.junit.jupiter.api.Assertions.assertThrows(expected, () -> chained.accept(reader()));
        org.junit.jupiter.api.Assertions.assertThrows(expected, () -> viaOptions.accept(ReaderOptions.builder()));
    }
}
