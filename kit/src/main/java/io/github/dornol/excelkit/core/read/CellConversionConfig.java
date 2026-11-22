package io.github.dornol.excelkit.core.read;

import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.ResolverStyle;
import java.time.temporal.ChronoField;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * Per-reader conversion settings used by {@link CellData}.
 * <p>
 * A {@code CellData} created without a config uses {@link #defaults()}. Readers attach an
 * immutable config to every cell, so locale and date parsing rules are scoped to that reader
 * and never change process-wide state.
 *
 * @since 0.19.0
 */
public final class CellConversionConfig {
    private final Locale locale;
    private final List<DateTimeFormatter> dateFormats;
    private final List<DateTimeFormatter> dateTimeFormats;

    private CellConversionConfig(Locale locale,
                                 List<DateTimeFormatter> dateFormats,
                                 List<DateTimeFormatter> dateTimeFormats) {
        this.locale = Objects.requireNonNull(locale, "locale must not be null");
        this.dateFormats = List.copyOf(dateFormats);
        this.dateTimeFormats = List.copyOf(dateTimeFormats);
    }

    private static final List<DateTimeFormatter> BUILT_IN_DATE_FORMATS = List.of(
            CellConversionSupport.strictDatePattern("yyyy-MM-dd"),
            CellConversionSupport.strictDatePattern("yyyy/MM/dd"),
            CellConversionSupport.strictDatePattern("MM/dd/yy"),
            CellConversionSupport.strictDatePattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE);
    private static final List<DateTimeFormatter> BUILT_IN_DATE_TIME_FORMATS = List.of(
            dateTimePattern("yyyy-MM-dd"),
            dateTimePattern("yyyy/MM/dd"),
            dateTimePattern("MM/dd/yy"),
            dateTimePattern("M/d/yy"),
            DateTimeFormatter.ISO_LOCAL_DATE_TIME);
    private static final CellConversionConfig DEFAULTS = new CellConversionConfig(
            Locale.getDefault(), BUILT_IN_DATE_FORMATS, BUILT_IN_DATE_TIME_FORMATS);

    private static DateTimeFormatter dateTimePattern(String date) {
        return new DateTimeFormatterBuilder().append(CellConversionSupport.strictDatePattern(date))
                .optionalStart().appendPattern(" HH:mm")
                .optionalStart().appendPattern(":ss")
                .optionalStart().appendFraction(ChronoField.NANO_OF_SECOND, 1, 9, true).optionalEnd()
                .optionalEnd().optionalEnd()
                .toFormatter().withResolverStyle(ResolverStyle.STRICT);
    }

    /**
     * Returns the built-in settings: the JVM default locale at class initialization and the
     * built-in date and date-time patterns. The returned instance is immutable.
     */
    public static CellConversionConfig defaults() {
        return DEFAULTS;
    }

    /**
     * Creates a mutable builder initialized from the built-in {@linkplain #defaults() defaults}.
     */
    public static Builder builder() {
        return new Builder(defaults());
    }

    /**
     * Creates a mutable builder initialized from this config.
     */
    public Builder toBuilder() {
        return new Builder(this);
    }

    public Locale locale() {
        return locale;
    }

    public List<DateTimeFormatter> dateFormats() {
        return dateFormats;
    }

    public List<DateTimeFormatter> dateTimeFormats() {
        return dateTimeFormats;
    }

    public static final class Builder {
        private Locale locale;
        private final List<DateTimeFormatter> dateFormats;
        private final List<DateTimeFormatter> dateTimeFormats;

        private Builder(CellConversionConfig source) {
            this.locale = source.locale;
            this.dateFormats = new CopyOnWriteArrayList<>(source.dateFormats);
            this.dateTimeFormats = new CopyOnWriteArrayList<>(source.dateTimeFormats);
        }

        public Builder locale(Locale locale) {
            this.locale = Objects.requireNonNull(locale, "locale must not be null");
            return this;
        }

        public Builder addDateFormat(String pattern) {
            dateFormats.add(0, CellConversionSupport.strictDatePattern(pattern));
            return this;
        }

        public Builder addDateTimeFormat(String pattern) {
            dateTimeFormats.add(0, CellConversionSupport.strictDatePattern(pattern));
            return this;
        }

        public Builder dateFormats(List<DateTimeFormatter> formats) {
            dateFormats.clear();
            dateFormats.addAll(Objects.requireNonNull(formats, "formats must not be null"));
            return this;
        }

        public Builder dateTimeFormats(List<DateTimeFormatter> formats) {
            dateTimeFormats.clear();
            dateTimeFormats.addAll(Objects.requireNonNull(formats, "formats must not be null"));
            return this;
        }

        public CellConversionConfig build() {
            return new CellConversionConfig(locale, dateFormats, dateTimeFormats);
        }
    }
}
