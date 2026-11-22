package io.github.dornol.excelkit.core.detect;

import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class TabularDetectionTest {

    @Test void detectorRecognizesCsvAndXlsxWithoutConsuming() {
        ByteArrayInputStream csv = csv("A,B\n1,2\n");
        assertEquals(TabularFileType.CSV, TabularFileDetector.detect(csv));
        assertEquals('A', csv.read());

        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create().column("Name", Person::name)
                .write(List.of(new Person("A"))).writeTo(output);
        assertEquals(TabularFileType.XLSX,
                TabularFileDetector.detect(new ByteArrayInputStream(output.toByteArray())));
    }

    @Test void detailedDetectionReportsCharsetDelimiterAndReadSupport() throws Exception {
        byte[] utf16 = "A;B\n1;2\n".getBytes(StandardCharsets.UTF_16LE);
        byte[] bom = new byte[utf16.length + 2];
        bom[0] = (byte) 0xff; bom[1] = (byte) 0xfe;
        System.arraycopy(utf16, 0, bom, 2, utf16.length);
        var result = TabularFileDetector.detectDetailed(new ByteArrayInputStream(bom));
        assertEquals(StandardCharsets.UTF_16LE, result.charset());
        assertEquals(';', result.delimiter());
        assertTrue(TabularFileType.XLSX.isReadable());
        assertFalse(TabularFileType.XLS.isReadable());
    }

    record Person(String name) {}

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
