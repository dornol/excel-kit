package io.github.dornol.excelkit.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.SpringBootTest.WebEnvironment;
import org.springframework.boot.test.web.server.LocalServerPort;

import java.io.ByteArrayInputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.junit.jupiter.api.Assertions.*;

/**
 * The showcase endpoints over a real HTTP connection (embedded server, streamed response): every workbook is opened
 * as a file would be, every formula is recalculated and compared with the value the file carries, and the subtotal
 * structure is checked cell by cell.
 */
@SpringBootTest(webEnvironment = WebEnvironment.RANDOM_PORT, properties = {
        "spring.docker.compose.enabled=false",
        "spring.datasource.url=jdbc:h2:mem:example-e2e;MODE=MariaDB;DB_CLOSE_DELAY=-1;DATABASE_TO_LOWER=TRUE",
        "spring.datasource.driver-class-name=org.h2.Driver",
        "spring.datasource.username=sa",
        "spring.datasource.password=",
        "spring.jpa.hibernate.ddl-auto=create-drop",
        "spring.jpa.show-sql=false",
        "example.dummy-count=0"
})
class ShowcaseEndToEndTest {
    private static final String XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private static final String[] SUBTOTAL_ENDPOINTS = {"subtotal", "subtotal-functions", "subtotal-layout",
            "subtotal-rollover", "subtotal-protection", "subtotal-rules", "subtotal-workbook", "subtotal-products",
            "subtotal-unsorted"};
    private static final Pattern SUBTOTAL = Pattern.compile("SUBTOTAL\\((\\d+),([A-Z]+)(\\d+):([A-Z]+)(\\d+)\\)");

    @LocalServerPort
    private int port;

    private final HttpClient client = HttpClient.newHttpClient();

    private HttpResponse<byte[]> get(String path) throws Exception {
        return client.send(HttpRequest.newBuilder(URI.create("http://localhost:" + port + path)).GET().build(),
                HttpResponse.BodyHandlers.ofByteArray());
    }

    private XSSFWorkbook download(String path) throws Exception {
        HttpResponse<byte[]> response = get(path);
        assertEquals(200, response.statusCode(), path);
        assertEquals(XLSX, response.headers().firstValue("Content-Type").orElse(""), path);
        assertTrue(response.headers().firstValue("Content-Disposition").orElse("").contains(".xlsx"), path);
        assertTrue(response.body().length > 0, path);
        return new XSSFWorkbook(new ByteArrayInputStream(response.body()));
    }

    private static boolean isSubtotalRow(Row row) {
        for (Cell cell : row) {
            if (cell.getCellType() == CellType.FORMULA && cell.getCellFormula().startsWith("SUBTOTAL(")) return true;
        }
        return false;
    }

    @Test
    void everyFormulaCarriesTheValueExcelCalculates() throws Exception {
        for (String endpoint : SUBTOTAL_ENDPOINTS) {
            try (XSSFWorkbook wb = download("/showcase/" + endpoint)) {
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                int checked = 0;
                for (Sheet sheet : wb) {
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (cell.getCellType() != CellType.FORMULA) continue;
                            String where = endpoint + " " + sheet.getSheetName() + "!" + cell.getAddress() + " " + cell.getCellFormula();
                            assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType(), where + " has no cached result");
                            double cached = cell.getNumericCellValue();
                            double calculated = evaluator.evaluate(cell).getNumberValue();
                            assertEquals(calculated, cached, 1e-6, where);
                            checked++;
                        }
                    }
                }
                assertTrue(checked > 0, endpoint + " has formulas");
            }
        }
    }

    @Test
    void everySubtotalCoversExactlyItsGroupAndNothingElse() throws Exception {
        for (String endpoint : SUBTOTAL_ENDPOINTS) {
            try (XSSFWorkbook wb = download("/showcase/" + endpoint)) {
                for (Sheet sheet : wb) {
                    int previousEnd = 0; // the header row (1-based row 1) ends the "previous" block
                    for (Row row : sheet) {
                        if (!isSubtotalRow(row)) continue;
                        int rowNumber = row.getRowNum() + 1;
                        int start = Integer.MAX_VALUE;
                        int end = -1;
                        for (Cell cell : row) {
                            if (cell.getCellType() != CellType.FORMULA) continue;
                            Matcher m = SUBTOTAL.matcher(cell.getCellFormula());
                            if (!m.matches()) continue;
                            String where = endpoint + " " + sheet.getSheetName() + " row " + rowNumber;
                            assertEquals(m.group(2), cell.getAddress().formatAsString().replaceAll("\\d+", ""), where + ": same column");
                            assertEquals(m.group(2), m.group(4), where);
                            start = Math.min(start, Integer.parseInt(m.group(3)));
                            end = Math.max(end, Integer.parseInt(m.group(5)));
                            assertEquals(Integer.parseInt(m.group(3)), start, where + ": every formula of a row starts alike");
                            assertEquals(Integer.parseInt(m.group(5)), end, where + ": and ends alike");
                        }
                        String where = endpoint + " " + sheet.getSheetName() + " row " + rowNumber;
                        assertTrue(end >= start && start >= 2, where + " range " + start + ":" + end);
                        if (start <= previousEnd) {
                            // a summary row: it spans the whole data block, subtotal rows included, which SUBTOTAL skips
                            assertEquals(2, start, where + ": a summary starts at the first data row");
                            assertTrue(end >= previousEnd, where + ": and reaches the last subtotal row");
                            for (int r = start - 1; r <= end - 1; r++) {
                                if (isSubtotalRow(sheet.getRow(r))) continue;
                                for (Cell inRange : sheet.getRow(r)) {
                                    assertFalse(inRange.getCellType() == CellType.FORMULA, where + ": only subtotal rows hold formulas in the range");
                                }
                            }
                            continue;
                        }
                        assertEquals(rowNumber - 1, end, where + ": the group ends right above its subtotal row");
                        assertTrue(start > previousEnd, where + ": starts after the previous subtotal row (" + previousEnd + ")");
                        for (int r = start - 1; r <= end - 1; r++) {
                            assertFalse(isSubtotalRow(sheet.getRow(r)), where + ": row " + (r + 1) + " is another subtotal row");
                        }
                        previousEnd = rowNumber;
                    }
                }
            }
        }
    }

    @Test
    void aSumSubtotalEqualsTheSumOfItsDataCells() throws Exception {
        int sums = 0;
        for (String endpoint : SUBTOTAL_ENDPOINTS) {
            try (XSSFWorkbook wb = download("/showcase/" + endpoint)) {
                for (Sheet sheet : wb) {
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (cell.getCellType() != CellType.FORMULA) continue;
                            Matcher m = SUBTOTAL.matcher(cell.getCellFormula());
                            if (!m.matches() || !m.group(1).equals("9")) continue;
                            int col = cell.getColumnIndex();
                            double expected = 0;
                            for (int r = Integer.parseInt(m.group(3)) - 1; r <= Integer.parseInt(m.group(5)) - 1; r++) {
                                Cell data = sheet.getRow(r).getCell(col);
                                if (data != null && data.getCellType() == CellType.NUMERIC) expected += data.getNumericCellValue();
                            }
                            assertEquals(expected, cell.getNumericCellValue(), 1e-6,
                                    endpoint + " " + sheet.getSheetName() + "!" + cell.getAddress());
                            sums++;
                        }
                    }
                }
            }
        }
        assertTrue(sums > 10, "sums checked: " + sums);
    }

    @Test
    void theRolloverWorkbookHasEverySheetCompleteAndSelfContained() throws Exception {
        try (XSSFWorkbook wb = download("/showcase/subtotal-rollover")) {
            assertTrue(wb.getNumberOfSheets() > 1, "it rolls over");
            List<String> names = new ArrayList<>();
            for (Sheet sheet : wb) {
                assertFalse(names.contains(sheet.getSheetName()), "unique sheet names");
                names.add(sheet.getSheetName());
                Row last = sheet.getRow(sheet.getLastRowNum());
                assertNotNull(last, sheet.getSheetName());
                boolean endsWithSubtotal = isSubtotalRow(last);
                boolean summary = false;
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.FORMULA && !cell.getCellFormula().startsWith("SUBTOTAL(")) summary = true;
                    }
                }
                assertTrue(endsWithSubtotal || summary || sheet.getLastRowNum() > 0, sheet.getSheetName() + " has content");
                // no group runs over into the next sheet: the last data row of a sheet is followed by its subtotal row
                int lastData = -1;
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue;
                    if (!isSubtotalRow(row)) lastData = row.getRowNum();
                }
                assertTrue(lastData >= 0, sheet.getSheetName());
                assertTrue(isSubtotalRow(sheet.getRow(lastData + 1)) || sheet.getRow(lastData + 1) == null
                                || !isSubtotalRow(sheet.getRow(lastData + 1)),
                        "structure readable");
                int dataBeforeEnd = lastData;
                assertTrue(isSubtotalRow(sheet.getRow(dataBeforeEnd + 1)), sheet.getSheetName() + ": the last group is closed on its own sheet");
            }
        }
    }

    @Test
    void protectionWorkbookLocksTheSubtotalRowsOnly() throws Exception {
        try (XSSFWorkbook wb = download("/showcase/subtotal-protection")) {
            Sheet sheet = wb.getSheetAt(0);
            assertTrue(sheet.getProtect(), "the sheet is protected");
            int subtotalRows = 0;
            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue;
                boolean subtotal = isSubtotalRow(row);
                for (Cell cell : row) {
                    boolean locked = cell.getCellStyle().getLocked();
                    if (subtotal) assertTrue(locked, "subtotal cell " + cell.getAddress() + " is locked");
                }
                if (subtotal) subtotalRows++;
            }
            assertTrue(subtotalRows > 0);
        }
    }

    @Test
    void theComboChartHasItsSeriesOnTheRightAxesOverHttp() throws Exception {
        try (XSSFWorkbook wb = download("/showcase/chart-combo")) {
            var plot = wb.getSheetAt(0).getDrawingPatriarch().getCharts().get(0).getCTChart().getPlotArea();
            assertEquals(1, plot.sizeOfBarChartArray());
            assertEquals(2, plot.sizeOfLineChartArray());
            assertEquals(4, plot.sizeOfCatAxArray() + plot.sizeOfValAxArray());
            int series = plot.getBarChartArray(0).sizeOfSerArray()
                    + plot.getLineChartArray(0).sizeOfSerArray() + plot.getLineChartArray(1).sizeOfSerArray();
            assertTrue(series >= 3, "bars plus a line on each axis: " + series);
        }
    }

    @Test
    void theEmptyMessageSheetHasOneMergedRowAndNoFormulas() throws Exception {
        try (XSSFWorkbook wb = download("/showcase/empty-message")) {
            Sheet empty = wb.getSheetAt(0);
            assertEquals(1, empty.getNumMergedRegions());
            CellRangeAddress merged = empty.getMergedRegion(0);
            assertEquals(1, merged.getFirstRow());
            assertEquals(1, merged.getLastRow());
            assertEquals(0, merged.getFirstColumn());
            assertTrue(merged.getLastColumn() > 0, "it spans the columns");
            assertEquals(CellType.STRING, empty.getRow(1).getCell(0).getCellType());
            assertFalse(empty.getRow(1).getCell(0).getStringCellValue().isBlank());
        }
    }

    @Test
    void theRejectedCombinationAnswersWithAClearClientError() throws Exception {
        HttpResponse<byte[]> response = get("/showcase/subtotal-rejected");
        assertEquals(400, response.statusCode());
        assertTrue(new String(response.body()).contains("cannot be combined with table"));
    }

    @Test
    void severalDownloadsAtOnceDoNotMixUp() throws Exception {
        List<Thread> threads = new ArrayList<>();
        List<Throwable> failures = java.util.Collections.synchronizedList(new ArrayList<>());
        for (int i = 0; i < 8; i++) {
            String endpoint = SUBTOTAL_ENDPOINTS[i % SUBTOTAL_ENDPOINTS.length];
            Thread t = new Thread(() -> {
                try (XSSFWorkbook wb = download("/showcase/" + endpoint)) {
                    boolean any = false;
                    for (Row row : wb.getSheetAt(0)) any |= isSubtotalRow(row);
                    assertTrue(any, endpoint);
                } catch (Throwable e) {
                    failures.add(e);
                }
            });
            threads.add(t);
            t.start();
        }
        for (Thread t : threads) t.join();
        assertTrue(failures.isEmpty(), failures.toString());
    }
}
