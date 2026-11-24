package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.read.ReadSessionConfig;

import java.nio.charset.Charset;

/** Immutable execution configuration consumed by the CSV engine. */
record CsvReadSessionConfig<T>(ReadSessionConfig<T> common, char delimiter, Charset charset,
        char quoteChar, char escapeChar, boolean strictQuotes, boolean ignoreLeadingWhiteSpace) {}
