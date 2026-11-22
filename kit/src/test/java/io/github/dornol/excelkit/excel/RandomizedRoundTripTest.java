package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Writes a randomly shaped table (column kinds, null positions, row count, sheet rollover, optional
 * encryption and structure options) and reads it back twice: with plain POI, which does not share any code
 * with excel-kit, and with {@link ExcelReader}. Every seed is fixed, so a failure names the seed that
 * reproduces it.
 */
class RandomizedRoundTripTest {

    private enum Kind {STRING, INT, LONG, DOUBLE, DATE, DATETIME}

    private static final String[] ALPHABET = {"a", "Z", "0", " ", "한", "글", "é", "ß", "😀", "=", "+", "-",
            "@", "'", "\"", ",", ";", "<", "&", "\n", "\t", "\\"};

    private static IntStream seeds() {
        return IntStream.range(0, 150);
    }

    private static String randomText(Random r) {
        int len = 1 + r.nextInt(r.nextInt(10) == 0 ? 200 : 12);
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < len; i++) sb.append(ALPHABET[r.nextInt(ALPHABET.length)]);
        return sb.toString();
    }

    private static Object randomValue(Kind kind, Random r) {
        return switch (kind) {
            case STRING -> randomText(r);
            case INT -> r.nextInt();
            case LONG -> (long) r.nextInt(1_000_000_000) * 1_000L + r.nextInt(1000) - 500; // < 15 digits
            case DOUBLE -> r.nextInt(2_000_000) / 1024.0 - 1000.0; // exactly representable
            case DATE -> LocalDate.of(1901, 1, 1).plusDays(r.nextInt(40_000));
            case DATETIME -> LocalDateTime.of(1901, 1, 1, 0, 0).plusSeconds(r.nextLong(1_000_000_000L));
        };
    }

    @ParameterizedTest(name = "seed {0}")
    @MethodSource("seeds")
    void randomTable_comesBackUnchanged(int seed) throws Exception {
        Random r = new Random(seed);
        int columns = 1 + r.nextInt(6);
        Kind[] kinds = new Kind[columns];
        double[] nullRate = new double[columns];
        for (int c = 0; c < columns; c++) {
            kinds[c] = Kind.values()[r.nextInt(Kind.values().length)];
            nullRate[c] = new double[]{0, 0.1, 0.5, 1.0}[r.nextInt(4)];
        }
        int rows = r.nextInt(5) == 0 ? 0 : r.nextInt(60);
        Object[][] data = new Object[rows][columns];
        for (int i = 0; i < rows; i++) {
            for (int c = 0; c < columns; c++) {
                data[i][c] = r.nextDouble() < nullRate[c] ? null : randomValue(kinds[c], r);
            }
        }

        int[] limits = {0, 1, 2, 3, 7, 10, Math.max(rows, 1), rows + 1};
        int maxRows = limits[r.nextInt(limits.length)];
        boolean rowNumbers = r.nextBoolean();
        boolean encrypted = r.nextInt(5) == 0;
        boolean autoFilter = r.nextBoolean();
        boolean freeze = r.nextBoolean();
        String describe = "seed=" + seed + " kinds=" + List.of(kinds) + " nullRate=" + java.util.Arrays.toString(nullRate)
                + " rows=" + rows + " maxRows=" + maxRows + " rowNumbers=" + rowNumbers + " encrypted=" + encrypted
                + " autoFilter=" + autoFilter + " freeze=" + freeze;

        ExcelWriter<Object[]> writer = ExcelWriter.create();
        if (encrypted) writer.password("pw-" + seed);
        if (maxRows > 0) writer.maxRows(maxRows);
        if (rowNumbers) writer.rowNumberColumn("No.");
        for (int c = 0; c < columns; c++) {
            final int idx = c;
            writer.column("C" + c, row -> row[idx]);
        }
        writer.autoFilter(autoFilter);
        if (freeze) writer.freezeRows(1);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(java.util.Arrays.asList(data)).writeTo(out);
        byte[] bytes = out.toByteArray();

        int perSheet = maxRows > 0 ? maxRows : Integer.MAX_VALUE;
        int expectedSheets = Math.max(1, (int) Math.ceil(rows / (double) perSheet));
        int offset = rowNumbers ? 1 : 0;

        // 1. Plain POI, independent of excel-kit's reader.
        XSSFWorkbook wb;
        if (encrypted) {
            var fs = new org.apache.poi.poifs.filesystem.POIFSFileSystem(new ByteArrayInputStream(bytes));
            var info = new org.apache.poi.poifs.crypt.EncryptionInfo(fs);
            var dec = org.apache.poi.poifs.crypt.Decryptor.getInstance(info);
            assertTrue(dec.verifyPassword("pw-" + seed), describe);
            wb = new XSSFWorkbook(dec.getDataStream(fs));
        } else {
            wb = new XSSFWorkbook(new ByteArrayInputStream(bytes));
        }
        try (wb) {
            assertEquals(expectedSheets, wb.getNumberOfSheets(), describe);
            int global = 0;
            for (int s = 0; s < expectedSheets; s++) {
                Sheet sheet = wb.getSheetAt(s);
                int inSheet = Math.min(perSheet, rows - global);
                assertEquals(Math.max(inSheet, 0), sheet.getLastRowNum(), describe + " sheet " + s);
                Row header = sheet.getRow(0);
                for (int c = 0; c < columns; c++) {
                    assertEquals("C" + c, header.getCell(c + offset).getStringCellValue(), describe);
                }
                for (int i = 0; i < inSheet; i++, global++) {
                    Row row = sheet.getRow(i + 1);
                    String at = describe + " row " + global;
                    if (rowNumbers) assertEquals(global + 1, (long) row.getCell(0).getNumericCellValue(), at);
                    for (int c = 0; c < columns; c++) {
                        assertCell(kinds[c], data[global][c], row.getCell(c + offset), at + " col " + c);
                    }
                }
            }
            assertEquals(rows, global, describe);
        }

        // 2. excel-kit's own reader, sheet by sheet.
        int global = 0;
        for (int s = 0; s < expectedSheets; s++) {
            List<ReadResult<Object[]>> results = new ArrayList<>();
            ExcelReader<Object[]> reader = ExcelReader.<Object[]>mapping(row -> {
                Object[] values = new Object[columns];
                for (int c = 0; c < columns; c++) values[c] = readBack(kinds[c], row.get("C" + c));
                return values;
            }).sheetIndex(s);
            if (encrypted) reader.password("pw-" + seed);
            reader.read(new ByteArrayInputStream(bytes), results::add);
            int inSheet = Math.min(perSheet, rows - global);
            assertEquals(Math.max(inSheet, 0), results.size(), describe + " sheet " + s);
            for (ReadResult<Object[]> result : results) {
                assertTrue(result.success(), describe + " row " + global + ": " + result.messages());
                for (int c = 0; c < columns; c++) {
                    assertEquals(data[global][c], result.data()[c], describe + " row " + global + " col " + c);
                }
                global++;
            }
        }
    }

    private static Object readBack(Kind kind, io.github.dornol.excelkit.core.read.CellData cell) {
        return switch (kind) {
            case STRING -> cell.asString().isEmpty() ? null : cell.asString();
            case INT -> cell.asInt();
            case LONG -> cell.asLong();
            case DOUBLE -> cell.asDouble();
            case DATE -> cell.asLocalDate();
            case DATETIME -> cell.asLocalDateTime();
        };
    }

    private static void assertCell(Kind kind, Object expected, Cell cell, String at) {
        if (expected == null) {
            if (cell != null) {
                assertEquals(org.apache.poi.ss.usermodel.CellType.BLANK, cell.getCellType(), at);
            }
            return;
        }
        switch (kind) {
            case STRING -> assertEquals(expected, cell.getStringCellValue(), at);
            case INT -> assertEquals(((Integer) expected).doubleValue(), cell.getNumericCellValue(), at);
            case LONG -> assertEquals(((Long) expected).doubleValue(), cell.getNumericCellValue(), at);
            case DOUBLE -> assertEquals((Double) expected, cell.getNumericCellValue(), at);
            case DATE -> assertEquals(expected, cell.getLocalDateTimeCellValue().toLocalDate(), at);
            case DATETIME -> assertEquals(expected, cell.getLocalDateTimeCellValue(), at);
        }
        assertNull(cell.getCellComment(), at);
    }
}
