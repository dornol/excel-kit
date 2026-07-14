package io.github.dornol.excelkit.core;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.IntFunction;
import java.util.function.IntUnaryOperator;
import java.util.function.UnaryOperator;

/** Header normalization, duplicate handling, and column-index resolution. */
final class HeaderResolver {
    private final boolean strict;
    private final DuplicateHeaderPolicy duplicatePolicy;
    private final UnaryOperator<String> normalizer;
    private final int maximumColumns;

    HeaderResolver(boolean strict, DuplicateHeaderPolicy duplicatePolicy,
                   UnaryOperator<String> normalizer, int maximumColumns) {
        this.strict = strict;
        this.duplicatePolicy = duplicatePolicy;
        this.normalizer = normalizer;
        this.maximumColumns = maximumColumns;
    }

    int[] resolve(int count, IntFunction<List<String>> aliases, IntUnaryOperator explicitIndex,
                  List<String> headers, String source) {
        Map<String, Integer> index = index(headers, source);
        int[] resolved = new int[count];
        for (int i = 0; i < count; i++) {
            int explicit = explicitIndex.applyAsInt(i);
            if (explicit >= 0) {
                validateIndex(explicit, headers, source);
                resolved[i] = explicit;
                continue;
            }
            List<String> candidates = aliases.apply(i);
            if (candidates == null || candidates.isEmpty()) {
                validateIndex(i, headers, source);
                resolved[i] = i;
                continue;
            }
            Integer match = candidates.stream().map(this::normalize).map(index::get)
                    .filter(java.util.Objects::nonNull).findFirst().orElse(null);
            if (match == null) throw new ExcelKitException("Header aliases " + candidates
                    + " not found in " + source + ". Available headers: " + headers);
            resolved[i] = match;
        }
        return resolved;
    }

    Map<String, Integer> index(List<String> headers, String source) {
        if (maximumColumns >= 0 && headers.size() > maximumColumns) {
            throw new ReadLimitExceededException(ReadLimitExceededException.Limit.COLUMNS,
                    maximumColumns, headers.size());
        }
        Map<String, Integer> result = new LinkedHashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String original = headers.get(i);
            String normalized = normalize(original);
            if (duplicatePolicy == DuplicateHeaderPolicy.FAIL && result.containsKey(normalized))
                throw new ExcelKitException("Duplicate header '" + original + "' found in " + source);
            if (duplicatePolicy == DuplicateHeaderPolicy.LAST) result.put(normalized, i);
            else result.putIfAbsent(normalized, i);
        }
        return result;
    }

    void validateSelected(Set<String> selected, Map<String, Integer> index,
                          List<String> headers, String source) {
        if (!strict || selected == null || selected.isEmpty()) return;
        List<String> missing = selected.stream().filter(name -> !index.containsKey(normalize(name))).toList();
        if (!missing.isEmpty()) throw new ExcelKitException("Selected headers " + missing
                + " not found in " + source + ". Available headers: " + headers);
    }

    String normalize(String header) {
        String normalized = normalizer.apply(header);
        if (normalized == null) throw new ExcelKitException("Header normalizer returned null for: " + header);
        return normalized;
    }

    private void validateIndex(int index, List<String> headers, String source) {
        if (strict && index >= headers.size()) throw new ExcelKitException("Column index " + index
                + " has no header in " + source + ". Available headers: " + headers);
    }
}
