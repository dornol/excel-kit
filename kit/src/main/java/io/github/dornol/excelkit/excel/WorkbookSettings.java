package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.jspecify.annotations.Nullable;

import java.util.Objects;
import java.util.function.Consumer;

/**
 * The settings of one workbook that {@link ExcelWriter} and {@link ExcelWorkbook} both expose: document
 * properties, the header style, the file password and the structure protection. The two classes only add their
 * own return type to each setter; what a setter means lives here.
 * <p>
 * The header style is rebuilt when it is asked for after a setting changed, so a series of setters creates one
 * cell style, not one per setter.
 */
final class WorkbookSettings {
    private final SXSSFWorkbook wb;
    private XSSFColor headerColor = color(ExcelColor.WHITE);
    private @Nullable String headerFontName;
    private @Nullable Integer headerFontSize;
    private @Nullable HeaderStyleConfig headerStyleConfig;
    private @Nullable CellStyle headerStyle;
    private @Nullable String password;
    private @Nullable String workbookPassword;

    WorkbookSettings(SXSSFWorkbook wb) {
        this.wb = wb;
    }

    private static XSSFColor color(ExcelColor color) {
        return new XSSFColor(new byte[]{(byte) color.getR(), (byte) color.getG(), (byte) color.getB()});
    }

    void documentProperty(String key, String value) {
        Objects.requireNonNull(value, "value cannot be null");
        ExcelWorkbookSupport.applyDocumentProperty(wb, key, value);
    }

    // ── Header style ──

    void headerColor(ExcelColor color) {
        if (color == null) throw new IllegalArgumentException("color must not be null");
        this.headerColor = color(color);
        this.headerStyle = null;
    }

    void headerFontName(String fontName) {
        this.headerFontName = Objects.requireNonNull(fontName, "fontName cannot be null");
        this.headerStyle = null;
    }

    void headerFontSize(int fontSize) {
        if (fontSize <= 0) {
            throw new IllegalArgumentException("fontSize must be positive");
        }
        this.headerFontSize = fontSize;
        this.headerStyle = null;
    }

    void headerStyle(Consumer<HeaderStyleConfig> configurer) {
        HeaderStyleConfig config = new HeaderStyleConfig();
        configurer.accept(config);
        this.headerStyleConfig = config;
        this.headerStyle = null;
    }

    /** The header style for the settings so far; built on first use and again after a setting changed. */
    CellStyle headerStyle() {
        if (headerStyle == null) {
            headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor, headerFontName, headerFontSize,
                    headerStyleConfig);
        }
        return headerStyle;
    }

    // ── Passwords ──

    void protectWorkbook(String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.workbookPassword = password;
    }

    void password(String password) {
        if (password == null || password.isBlank()) {
            throw new IllegalArgumentException("Password cannot be null or blank");
        }
        this.password = password;
    }

    /** The file encryption password, or {@code null} when the file is not encrypted. */
    @Nullable String password() {
        return password;
    }

    /** The structure protection password, or {@code null}. */
    @Nullable String workbookPassword() {
        return workbookPassword;
    }
}
