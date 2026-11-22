package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalTime;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ExcelDataTypeRoundTripTest {

    // -----------------------------------------------------------------------
    // 4. ExcelDataType — TIME, FLOAT, FLOAT_PERCENT, BIG_DECIMAL_TO_DOUBLE,
    //    BIG_DECIMAL_TO_LONG, BOOLEAN_TO_YN write and read back
    // -----------------------------------------------------------------------

    @Test
    void excelDataType_TIME_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<LocalTime>create()
                .column("Time", t -> t, cfg -> cfg.type(ExcelDataType.TIME))
                .write(Stream.of(LocalTime.of(14, 30, 15)))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            // TIME values are stored as date-time in Excel; verify format string
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("hh:mm:ss", format);
        }
    }

    @Test
    void excelDataType_FLOAT_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<Float>create()
                .column("Float", f -> f, cfg -> cfg.type(ExcelDataType.FLOAT))
                .write(Stream.of(3.14f))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(3.14, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void excelDataType_FLOAT_PERCENT_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<Float>create()
                .column("Pct", f -> f, cfg -> cfg.type(ExcelDataType.FLOAT_PERCENT))
                .write(Stream.of(0.25f))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(0.25, cell.getNumericCellValue(), 0.001);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("0.00%", format);
        }
    }

    @Test
    void excelDataType_BIG_DECIMAL_TO_DOUBLE_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<BigDecimal>create()
                .column("BD", bd -> bd, cfg -> cfg.type(ExcelDataType.BIG_DECIMAL_TO_DOUBLE))
                .write(Stream.of(new BigDecimal("123.45")))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(123.45, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void excelDataType_BIG_DECIMAL_TO_LONG_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<BigDecimal>create()
                .column("BD", bd -> bd, cfg -> cfg.type(ExcelDataType.BIG_DECIMAL_TO_LONG))
                .write(Stream.of(new BigDecimal("99999")))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(99999.0, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0", format);
        }
    }

    @Test
    void excelDataType_BOOLEAN_TO_YN_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        ExcelWriter.<Boolean>create()
                .column("Flag", b -> b, cfg -> cfg.type(ExcelDataType.BOOLEAN_TO_YN))
                .write(Stream.of(true, false))
                .writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals("Y", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("N", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }
}
