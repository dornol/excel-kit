package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.detect.TabularFileDetector;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.read.ExcelReader;
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

    private static byte[] zipWithEntry(String name, int zeroBytes) throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(output)) {
            zip.putNextEntry(new ZipEntry(name));
            zip.write(new byte[zeroBytes]);
            zip.closeEntry();
        }
        return output.toByteArray();
    }

    @Test void theDefaultPolicyAlsoRejectsAZipBombInAPartItDoesNotScan() throws Exception {
        // Not a worksheet and no formula or link scan is on: only the ratio check can notice it.
        byte[] bomb = zipWithEntry("xl/media/image1.bin", 20_000_000);

        ReadSecurityException error = assertThrows(ReadSecurityException.class, () ->
                ExcelReader.forMap().read(new ByteArrayInputStream(bomb), result -> {}));

        assertEquals(ReadSecurityException.Reason.COMPRESSION_RATIO, error.reason());
        assertTrue(error.getMessage().contains("xl/media/image1.bin"), error.getMessage());
    }

    @Test void aHigherRatioLimitLetsTheSameEntryThroughTheScan() throws Exception {
        byte[] bomb = zipWithEntry("xl/media/image1.bin", 20_000_000);
        ReadSecurityPolicy lenient = new ReadSecurityPolicy(true, true,
                ReadSecurityPolicy.DEFAULT.maxScannedEntryBytes(), ReadSecurityPolicy.DEFAULT.maxTotalScannedBytes(),
                1_000_000.0);

        // The scan lets it pass; what follows is POI failing on a package that is no workbook, not a security error.
        Exception error = assertThrows(Exception.class, () ->
                ExcelReader.forMap().securityPolicy(lenient).read(new ByteArrayInputStream(bomb), result -> {}));

        assertFalse(error instanceof ReadSecurityException, String.valueOf(error));
    }

    @Test void aWorkbookThisLibraryWritesIsNotMistakenForAZipBomb() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        io.github.dornol.excelkit.excel.ExcelWriter.<Integer>create().column("A", n -> null).column("B", n -> n)
                .write(java.util.stream.IntStream.range(0, 200_000).boxed().toList()).writeTo(out);

        long rows = ExcelReader.forMap().readAll(new ByteArrayInputStream(out.toByteArray())).size();

        assertEquals(200_000, rows);
    }
}
