package io.github.dornol.excelkit.core;

import java.text.Normalizer;
import java.util.Locale;
import java.util.function.UnaryOperator;

/** Common header normalization presets. */
public enum HeaderPolicy {
    EXACT(UnaryOperator.identity()),
    TRIM(String::trim),
    TRIM_CASE_INSENSITIVE(value -> value.trim().toLowerCase(Locale.ROOT)),
    NORMALIZED(value -> Normalizer.normalize(value.trim(), Normalizer.Form.NFKC)),
    NORMALIZED_CASE_INSENSITIVE(value -> Normalizer.normalize(value.trim(), Normalizer.Form.NFKC)
            .replaceAll("\\s+", " ").toLowerCase(Locale.ROOT));

    private final UnaryOperator<String> normalizer;
    HeaderPolicy(UnaryOperator<String> normalizer) { this.normalizer = normalizer; }
    public UnaryOperator<String> normalizer() { return normalizer; }
}
