package io.github.dornol.excelkit.core;

import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * Per-reader conversion settings used by {@link CellData}.
 * <p>
 * Existing no-arg {@code CellData} conversions use global defaults for backward
 * compatibility. Readers can attach an immutable config to every cell so locale
 * and date parsing rules do not have to be changed globally.
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

    /**
     * Creates a config from the current global {@link CellData} defaults.
     */
    public static CellConversionConfig defaults() {
        return new CellConversionConfig(
                CellData.getDefaultLocale(),
                CellData.getDateFormats(),
                CellData.getDateTimeFormats());
    }

    /**
     * Creates a mutable builder initialized from current global defaults.
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
            dateFormats.add(0, DateTimeFormatter.ofPattern(pattern));
            return this;
        }

        public Builder addDateTimeFormat(String pattern) {
            dateTimeFormats.add(0, DateTimeFormatter.ofPattern(pattern));
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
