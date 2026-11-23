package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class RequiredMarkTest {

    private static final String MARK = "@\" *\"";

    static class Person {
        String name;
        int age;
        String getName() { return name; }
        int getAge() { return age; }
    }

    private static Person person(String name, int age) {
        var p = new Person();
        p.name = name;
        p.age = age;
        return p;
    }

    private static byte[] bytes(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return out.toByteArray();
    }

    @Nested
    class Header {
        @Test
        void showsAStarButKeepsThePlainNameAsTheCellValue() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .column("Name", Person::getName, c -> c.requiredMark())
                    .column("Age", Person::getAge)
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var header = wb.getSheetAt(0).getRow(0);
                assertEquals("Name", header.getCell(0).getStringCellValue());
                assertEquals(MARK, header.getCell(0).getCellStyle().getDataFormatString());
                assertEquals("Age", header.getCell(1).getStringCellValue());
                assertNotEquals(MARK, header.getCell(1).getCellStyle().getDataFormatString());
            }
        }

        @Test
        void aFileWithTheMarkReadsBackByTheSameHeaderNames() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .column("Name", Person::getName, c -> c.requiredMark())
                    .column("Age", Person::getAge)
                    .write(Stream.of(person("Ann", 30))));
            var read = new ArrayList<Person>();
            ExcelReader.setter(Person::new)
                    .column("Name", (p, c) -> p.name = c.asString())
                    .column("Age", (p, c) -> p.age = c.asInt())
                    .read(new ByteArrayInputStream(file), r -> read.add(r.data()));
            assertEquals(1, read.size());
            assertEquals("Ann", read.get(0).name);
            assertEquals(30, read.get(0).age);
        }

        @Test
        void keepsAColumnsHeaderColorsAndWorksWithGroupHeaders() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .column("Name", Person::getName, c -> c.group("People").requiredMark()
                            .headerBackgroundColor(io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW))
                    .column("Age", Person::getAge, c -> c.group("People"))
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var cell = wb.getSheetAt(0).getRow(1).getCell(0);
                assertEquals("Name", cell.getStringCellValue());
                assertEquals(MARK, cell.getCellStyle().getDataFormatString());
                assertNotNull(cell.getCellStyle().getFillForegroundColorColor());
            }
        }

        @Test
        void canBeTurnedOffAgain() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .column("Name", Person::getName, c -> c.requiredMark().requiredMark(false))
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                assertNotEquals(MARK, wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getDataFormatString());
            }
        }

        @Test
        void aDefaultColumnConfigCanMarkEveryColumn() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .defaultStyle(d -> d.requiredMark())
                    .column("Name", Person::getName)
                    .column("Age", Person::getAge)
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var header = wb.getSheetAt(0).getRow(0);
                assertEquals(MARK, header.getCell(0).getCellStyle().getDataFormatString());
                assertEquals(MARK, header.getCell(1).getCellStyle().getDataFormatString());
            }
        }
    }

    @Nested
    class InASchema {
        @Test
        void aRequiredColumnShowsTheMarkOnItsHeader() throws IOException {
            var schema = ExcelKitSchema.<Person>builder()
                    .requiredColumn("Name", Person::getName, (p, c) -> p.name = c.asString())
                    .column("Age", Person::getAge, (p, c) -> p.age = c.asInt())
                    .build();
            var file = bytes(schema.excelWriter().write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var header = wb.getSheetAt(0).getRow(0);
                assertEquals("Name", header.getCell(0).getStringCellValue());
                assertEquals(MARK, header.getCell(0).getCellStyle().getDataFormatString());
                assertNotEquals(MARK, header.getCell(1).getCellStyle().getDataFormatString());
            }
        }

        @Test
        void theSchemasOwnWriteConfigurationMayTurnTheMarkOff() throws IOException {
            var schema = ExcelKitSchema.<Person>builder()
                    .requiredColumn("Name", Person::getName, (p, c) -> p.name = c.asString(), b -> b.requiredMark(false))
                    .build();
            var file = bytes(schema.excelWriter().write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                assertNotEquals(MARK, wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getDataFormatString());
            }
        }
    }

    @Nested
    class MoreCases {
        @Test
        void aTableHeaderKeepsTheColumnNameAndTheMark() throws IOException {
            var file = bytes(ExcelWriter.<Person>create().table("People")
                    .column("Name", Person::getName, c -> c.requiredMark())
                    .column("Age", Person::getAge)
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals(MARK, sheet.getRow(0).getCell(0).getCellStyle().getDataFormatString());
                var table = sheet.getTables().get(0);
                assertEquals("Name", table.getCTTable().getTableColumns().getTableColumnArray(0).getName(),
                        "the table column name is the plain name, as Excel requires of the header cell");
            }
        }

        @Test
        void aColumnWithItsOwnHeaderColorsKeepsThemWithTheMark() throws IOException {
            var file = bytes(ExcelWriter.<Person>create()
                    .column("Name", Person::getName, c -> c.requiredMark()
                            .headerBackgroundColor(io.github.dornol.excelkit.excel.style.ExcelColor.LIGHT_YELLOW)
                            .headerFontColor(io.github.dornol.excelkit.excel.style.ExcelColor.RED))
                    .write(Stream.of(person("Ann", 30))));
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(file))) {
                var style = wb.getSheetAt(0).getRow(0).getCell(0).getCellStyle();
                assertEquals(MARK, style.getDataFormatString());
                assertEquals(255, style.getFillForegroundColorColor().getRGB()[0] & 0xFF);
                assertEquals(255, ((org.apache.poi.xssf.usermodel.XSSFFont) style.getFont()).getXSSFColor().getRGB()[0] & 0xFF);
            }
        }
    }
}
