package io.github.dornol.excelkit.csv;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * Predefined CSV dialect configurations for common formats.
 * <p>
 * Use with {@link CsvWriter#dialect(CsvDialect)} or {@link CsvReader#dialect(CsvDialect)}
 * to quickly configure delimiter, charset, and BOM settings.
 *
 * <pre>{@code
 * // Tab-separated values
 * new CsvWriter<Item>()
 *     .dialect(CsvDialect.TSV)
 *     .column("Name", Item::name)
 *     .write(stream);
 *
 * // Excel-compatible CSV (UTF-8 BOM)
 * new CsvWriter<Item>()
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
     * RFC 4180 standard: comma delimiter, UTF-8, no BOM.
     */
    RFC4180(',', StandardCharsets.UTF_8, false),

    /**
     * Excel-compatible: comma delimiter, UTF-8 with BOM for proper Korean/CJK display.
     */
    EXCEL(',', StandardCharsets.UTF_8, true),

    /**
     * Tab-separated values: tab delimiter, UTF-8, no BOM.
     */
    TSV('\t', StandardCharsets.UTF_8, false),

    /**
     * Pipe-separated values: pipe delimiter, UTF-8, no BOM.
     */
    PIPE('|', StandardCharsets.UTF_8, false);

    private final char delimiter;
    private final Charset charset;
    private final boolean bom;

    CsvDialect(char delimiter, Charset charset, boolean bom) {
        this.delimiter = delimiter;
        this.charset = charset;
        this.bom = bom;
    }

    /**
     * Returns the field delimiter character.
     */
    public char getDelimiter() {
        return delimiter;
    }

    /**
     * Returns the character encoding.
     */
    public Charset getCharset() {
        return charset;
    }

    /**
     * Returns whether a UTF-8 BOM should be written.
     */
    public boolean isBom() {
        return bom;
    }
}
