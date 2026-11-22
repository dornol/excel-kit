package io.github.dornol.excelkit.core.read;

import java.text.Normalizer;
import java.util.Locale;
import java.util.function.UnaryOperator;

/**
 * Common header normalization presets.
 *
 * @since 0.21.0
 */
public enum HeaderPolicy {
    /** Headers must match exactly. */
    EXACT(UnaryOperator.identity()),
    /** Ignores leading and trailing whitespace. */
    TRIM(String::trim),
    /** {@link #TRIM}, and ignores case. */
    TRIM_CASE_INSENSITIVE(value -> value.trim().toLowerCase(Locale.ROOT)),
    /** {@link #TRIM}, and Unicode NFKC normalization (full-width letters and digits match their ASCII forms). */
    NORMALIZED(value -> Normalizer.normalize(value.trim(), Normalizer.Form.NFKC)),
    /** {@link #NORMALIZED}, collapses runs of whitespace to one space, and ignores case. */
    NORMALIZED_CASE_INSENSITIVE(value -> Normalizer.normalize(value.trim(), Normalizer.Form.NFKC)
            .replaceAll("\\s+", " ").toLowerCase(Locale.ROOT));

    private final UnaryOperator<String> normalizer;
    HeaderPolicy(UnaryOperator<String> normalizer) { this.normalizer = normalizer; }
    /** The header transformation this preset applies before names and aliases are compared. */
    public UnaryOperator<String> normalizer() { return normalizer; }
}
