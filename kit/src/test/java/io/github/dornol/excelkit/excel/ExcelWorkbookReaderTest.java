package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ExcelWorkbookReaderTest {

    record User(String name) {}

    @Test
    void readsDifferentSheetsAndAggregatesReports() {
        byte[] workbook = workbook();
        var users = new java.util.ArrayList<String>();
        var orders = new java.util.ArrayList<String>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet("Users", ExcelReader.mapping(row -> row.get("Name").asString()),
                        result -> { if (result.success()) users.add(result.data()); })
                .sheet("Orders", ExcelReader.mapping(row -> row.get("Order").asString()),
                        result -> { if (result.success()) orders.add(result.data()); })
                .read(new ByteArrayInputStream(workbook));

        assertTrue(report.successful());
        assertEquals(List.of("Alice"), users);
        assertEquals(List.of("1001"), orders);
        assertEquals(1, report.sheet("Users").summary().successRows());
        assertEquals(1, report.sheet("Orders").summary().successRows());
    }

    @Test
    void continuesAfterStructuralSheetFailureByDefault() {
        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet("Users", ExcelReader.<String>setter(String::new)
                                .columnAt(5, (value, cell) -> {}).strictHeaders(), result -> {})
                .sheet("Orders", ExcelReader.mapping(row -> row.get("Order").asString()), result -> {})
                .read(new ByteArrayInputStream(workbook()));

        assertNotNull(report.sheet("Users").failure());
        assertEquals(1, report.sheet("Orders").summary().successRows());
    }

    @Test
    void aMissingNamedSheetIsAPerSheetFailureLikeAMissingIndex() {
        var orders = new java.util.ArrayList<String>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet("Missing", ExcelReader.forMap(), result -> {})
                .sheet("Orders", ExcelReader.mapping(row -> row.get("Order").asString()),
                        result -> { if (result.success()) orders.add(result.data()); })
                .read(new ByteArrayInputStream(workbook()));

        assertFalse(report.successful());
        assertEquals(2, report.sheets().size());
        assertTrue(report.sheet("Missing").failure().getMessage().contains("'Missing' is not present"));
        assertEquals(List.of("1001"), orders, "later sheets are still read under the default policy");

        var stopped = new java.util.ArrayList<String>();
        WorkbookReadReport stop = ExcelWorkbookReader.create()
                .failurePolicy(SheetFailurePolicy.STOP)
                .sheet("Missing", ExcelReader.forMap(), result -> {})
                .sheet("Orders", ExcelReader.mapping(row -> row.get("Order").asString()),
                        result -> stopped.add(result.data()))
                .read(new ByteArrayInputStream(workbook()));

        assertEquals(1, stop.sheets().size());
        assertTrue(stopped.isEmpty());
    }

    @Test
    void stopPolicyDoesNotProcessLaterSheets() {
        WorkbookReadReport report = ExcelWorkbookReader.create()
                .failurePolicy(SheetFailurePolicy.STOP)
                .sheet(0, ExcelReader.<String>setter(String::new)
                                .columnAt(5, (value, cell) -> {}).strictHeaders(), result -> {})
                .sheet(1, ExcelReader.mapping(row -> row.get("Order").asString()), result -> {})
                .read(new ByteArrayInputStream(workbook()));

        assertEquals(1, report.sheets().size());
        assertNotNull(report.sheets().get(0).failure());
    }

    private static byte[] workbook() {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<User>sheet("Users").column("Name", User::name).write(List.of(new User("Alice")));
            workbook.<String>sheet("Orders").column("Order", value -> value).write(List.of("1001"));
            workbook.finish().writeTo(output);
        }
        return output.toByteArray();
    }
}
