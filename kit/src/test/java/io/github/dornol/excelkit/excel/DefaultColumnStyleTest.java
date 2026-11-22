package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class DefaultColumnStyleTest {

    @Test
    void excelWriter_defaultStyle_allColumnsInheritBoldFontNameAlignment() throws Exception {
        var handler = ExcelWriter.<String>create()
                .defaultStyle(d -> d.bold(true).fontName("Arial").alignment(HorizontalAlignment.LEFT))
                .column("Name", s -> s)
                .column("Value", s -> s)
                .write(Stream.of("Test"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var row = wb.getSheetAt(0).getRow(1); // data row
            for (int i = 0; i < 2; i++) {
                var cell = row.getCell(i);
                var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
                assertTrue(font.getBold(), "Column " + i + " should be bold from default");
                assertEquals("Arial", font.getFontName(),
                        "Column " + i + " should have Arial from default");
                assertEquals(HorizontalAlignment.LEFT, cell.getCellStyle().getAlignment(),
                        "Column " + i + " should have LEFT alignment from default");
            }
        }
    }

    @Test
    void columnLevelOverride_winsOverDefault() throws Exception {
        var handler = ExcelWriter.<String>create()
                .defaultStyle(d -> d.bold(true).fontName("Arial"))
                .column("Name", s -> s, c -> c
                    .bold(false))  // override bold
                .write(Stream.of("Test"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
            assertFalse(font.getBold(), "Column override (bold=false) should win over default");
            assertEquals("Arial", font.getFontName(),
                    "Non-overridden fontName should come from default");
        }
    }

    @Test
    void excelSheetWriter_defaultStyle_appliedToColumn() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Data")
                    .defaultStyle(d -> d.fontName("Courier New"))
                    .column("Name", s -> s)
                    .write(Stream.of("Test"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            var font = wb.getFontAt(cell.getCellStyle().getFontIndex());
            assertEquals("Courier New", font.getFontName());
        }
    }

    @Test
    void defaultStyle_withVerticalAlignmentWrapTextFontSizeFontColor() throws Exception {
        var handler = ExcelWriter.<String>create()
                .defaultStyle(d -> d
                        .verticalAlignment(VerticalAlignment.TOP)
                        .wrapText(true)
                        .fontSize(14)
                        .fontColor(255, 0, 0))
                .column("Col", s -> s)
                .write(Stream.of("Test"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            var style = cell.getCellStyle();
            assertEquals(VerticalAlignment.TOP, style.getVerticalAlignment());
            assertTrue(style.getWrapText(), "wrapText should be enabled from default");
            var font = wb.getFontAt(style.getFontIndex());
            assertEquals(14, font.getFontHeightInPoints());
        }
    }

    @Test
    void noDefaultStyle_columnsUseNormalDefaults() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Col", s -> s)
                .write(Stream.of("Test"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            var style = cell.getCellStyle();
            // default alignment is CENTER per ColumnStyleConfig
            assertEquals(HorizontalAlignment.CENTER, style.getAlignment());
        }
    }

    @Test
    void applyDefaults_fillsNullFields() {
        var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
        defaults.bold(true).fontName("Arial").fontSize(14).alignment(HorizontalAlignment.LEFT);

        var config = new ColumnConfig<String>();
        // Only set bold -- everything else should come from defaults
        config.bold(false);

        config.applyDefaults(defaults);

        assertEquals(false, config.bold, "Explicit bold=false should be kept");
        assertEquals("Arial", config.fontName, "fontName should come from defaults");
        assertEquals(14, config.fontSize, "fontSize should come from defaults");
        assertEquals(HorizontalAlignment.LEFT, config.alignment,
                "alignment should come from defaults");
    }

    @Test
    void applyDefaults_doesNotOverrideExplicitValues() {
        var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
        defaults.bold(true).fontName("Arial");

        var config = new ColumnConfig<String>();
        config.bold(false).fontName("Courier");

        config.applyDefaults(defaults);

        assertEquals(false, config.bold, "Explicit bold=false should not be overridden");
        assertEquals("Courier", config.fontName, "Explicit fontName should not be overridden");
    }

    @Test
    void applyDefaults_alignmentSetFlagTracked() {
        var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
        defaults.alignment(HorizontalAlignment.RIGHT);

        var configWithoutAlignment = new ColumnConfig<String>();
        // alignmentSet is false by default, so defaults should apply
        configWithoutAlignment.applyDefaults(defaults);
        assertEquals(HorizontalAlignment.RIGHT, configWithoutAlignment.alignment,
                "Default alignment should apply when column did not set alignment");
        assertTrue(configWithoutAlignment.alignmentSet,
                "alignmentSet flag should be propagated from defaults");

        var configWithAlignment = new ColumnConfig<String>();
        configWithAlignment.alignment(HorizontalAlignment.LEFT);
        configWithAlignment.applyDefaults(defaults);
        assertEquals(HorizontalAlignment.LEFT, configWithAlignment.alignment,
                "Explicit alignment should not be overridden by defaults");
    }

    @Test
    void applyDefaults_nullDefaultFieldsDoNotOverwrite() {
        var defaults = new ColumnStyleConfig.DefaultStyleConfig<String>();
        // defaults has no explicit settings

        var config = new ColumnConfig<String>();
        config.bold(true).fontName("Helvetica");

        config.applyDefaults(defaults);

        assertEquals(true, config.bold, "Existing bold should remain");
        assertEquals("Helvetica", config.fontName, "Existing fontName should remain");
    }
}
