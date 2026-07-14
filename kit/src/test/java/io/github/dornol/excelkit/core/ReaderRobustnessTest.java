package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

class ReaderRobustnessTest {
    @Test void deterministicCsvFuzzNeverHangsAndPreservesDetectorPosition() {
        assertTimeoutPreemptively(Duration.ofSeconds(5), () -> {
            Random random = new Random(20260711L);
            for (int seed = 0; seed < 250; seed++) {
                byte[] bytes = new byte[random.nextInt(512) + 1];
                random.nextBytes(bytes);
                ByteArrayInputStream detectorInput = new ByteArrayInputStream(bytes);
                TabularFileDetector.detectDetailed(detectorInput);
                assertEquals(bytes[0] & 0xff, detectorInput.read());

                try {
                    CsvReader.forMap().limits(new ReadLimits(1024, -1, 64, 256))
                            .read(new ByteArrayInputStream(bytes), result -> {});
                } catch (ExcelKitException | IllegalArgumentException ignored) {
                    // Malformed random input is expected; termination and bounded handling are the contract.
                }
            }
        });
    }

    @Test void truncatedZipFailsWithoutLeakingOrHanging() {
        assertTimeout(Duration.ofSeconds(2), () -> assertThrows(RuntimeException.class, () ->
                ExcelReader.forMap().read(new ByteArrayInputStream(new byte[]{'P','K',3,4,1,2,3}), r -> {})));
    }

    @Test void strictSecurityRejectsExtremeCompressionRatioBeforePoiParsing() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(output)) {
            zip.putNextEntry(new ZipEntry("xl/worksheets/sheet1.xml"));
            zip.write("<worksheet>".getBytes(StandardCharsets.UTF_8));
            zip.write(new byte[2_000_000]);
            zip.write("</worksheet>".getBytes(StandardCharsets.UTF_8));
            zip.closeEntry();
        }
        ReadSecurityException error = assertThrows(ReadSecurityException.class, () ->
                ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT)
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> {}));
        assertTrue(error.reason() == ReadSecurityException.Reason.COMPRESSION_RATIO
                || error.reason() == ReadSecurityException.Reason.ENTRY_SIZE);
    }
}
