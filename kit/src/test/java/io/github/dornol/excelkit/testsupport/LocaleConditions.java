package io.github.dornol.excelkit.testsupport;

import java.text.DecimalFormatSymbols;

/** Conditions on the JVM default locale for {@code @EnabledIf}. */
public final class LocaleConditions {
    private LocaleConditions() {
    }

    /** Whether the JVM default locale writes numbers the English way: {@code 1,234.5}, {@code -5}, ASCII digits. */
    public static boolean englishNumbers() {
        DecimalFormatSymbols symbols = DecimalFormatSymbols.getInstance();
        return symbols.getDecimalSeparator() == '.' && symbols.getGroupingSeparator() == ','
                && symbols.getMinusSign() == '-' && symbols.getZeroDigit() == '0';
    }
}
