package io.github.dornol.excelkit.csv;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * Predefined CSV dialect configurations for common formats.
 * <p>
 * Use with {@link CsvWriter#dialect(CsvDialect)} or {@link CsvReader#dialect(CsvDialect)}
 * to quickly configure delimiter, charset, BOM and line separator settings.
 *
 * <pre>{@code
 * // Tab-separated values
 * CsvWriter.<Item>create()
 *     .dialect(CsvDialect.TSV)
 *     .column("Name", Item::name)
 *     .write(stream);
 *
 * // Excel-compatible CSV (UTF-8 BOM)
 * CsvWriter.<Item>create()
 *     .dialect(CsvDialect.EXCEL)
 *     .column("Name", Item::name)
 *     .write(stream);
 * }</pre>
 *
 * @author dhkim
 * @since 0.9.2
 */
public enum CsvDialect {

    /**
     * RFC 4180 standard: comma delimiter, UTF-8, no BOM, CRLF line ends.
     */
    RFC4180(',', StandardCharsets.UTF_8, false, "\r\n"),

    /**
     * Excel-compatible: comma delimiter, UTF-8 with BOM for proper Korean/CJK display, CRLF line ends.
     */
    EXCEL(',', StandardCharsets.UTF_8, true, "\r\n"),

    /**
     * Tab-separated values: tab delimiter, UTF-8, no BOM, LF line ends
     * (TSV has no single standard, and Unix tools expect LF).
     */
    TSV('\t', StandardCharsets.UTF_8, false, "\n"),

    /**
     * Pipe-separated values: pipe delimiter, UTF-8, no BOM, LF line ends.
     */
    PIPE('|', StandardCharsets.UTF_8, false, "\n");

    private final char delimiter;
    private final Charset charset;
    private final boolean bom;
    private final String lineSeparator;

    CsvDialect(char delimiter, Charset charset, boolean bom, String lineSeparator) {
        this.lineSeparator = lineSeparator;
        this.delimiter = delimiter;
        this.charset = charset;
        this.bom = bom;
    }

    /**
     * Returns the field delimiter character.
     *
     * @return the delimiter character
     */
    public char getDelimiter() {
        return delimiter;
    }

    /**
     * Returns the character encoding.
     *
     * @return the charset
     */
    public Charset getCharset() {
        return charset;
    }

    /**
     * Returns whether a UTF-8 BOM should be written.
     *
     * @return true if BOM is enabled
     */
    public boolean isBom() {
        return bom;
    }

    /**
     * Returns the line terminator a {@link CsvWriter} writes after every row.
     *
     * @return {@code "\r\n"} or {@code "\n"}
     * @since 0.27.0
     */
    public String getLineSeparator() {
        return lineSeparator;
    }
}
