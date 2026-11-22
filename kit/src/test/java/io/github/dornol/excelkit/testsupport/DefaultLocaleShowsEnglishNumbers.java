package io.github.dornol.excelkit.testsupport;

import org.junit.jupiter.api.condition.EnabledIf;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * The test asserts the text of numbers as an English-style default locale shows them ({@code 1,234.5}, {@code -5},
 * ASCII digits), or reads text written that way with the default conversion locale. Readers take their conversion
 * locale from the JVM default (see {@code CellConversionConfig.defaults()}), so under another default locale the
 * same call correctly gives another result (German reads {@code 2.718} as 2718). The test is skipped there instead of
 * failing; what must hold in every locale is covered by {@code LocaleIndependenceTest}.
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.METHOD, ElementType.TYPE})
@EnabledIf(value = "io.github.dornol.excelkit.testsupport.LocaleConditions#englishNumbers",
        disabledReason = "asserts number text as the default English locale shows it; the default locale here differs")
public @interface DefaultLocaleShowsEnglishNumbers {
}
