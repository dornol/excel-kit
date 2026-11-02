package io.github.dornol.excelkit.benchmark;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.Files;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.*;

class BenchmarkProbeTest {
    @TempDir Path root;

    @Test
    void preservesTemporaryFilePeakAfterCleanupAndExcludesInputAndOutput() throws Exception {
        Path excluded = Files.createDirectory(root.resolve("input-output"));
        Files.write(excluded.resolve("input.xlsx"), new byte[8192]);
        Files.write(root.resolve("existing.tmp"), new byte[64]);
        try (var probe = new BenchmarkProbe(root, excluded, 10, false)) {
            Path spool = Files.createDirectory(root.resolve("poifiles")).resolve("sheet.xml");
            Files.write(spool, new byte[2048]);
            probe.sampleNow();
            Files.delete(spool);
            var result = probe.stop();
            assertEquals(64, result.tempBaselineBytes());
            assertEquals(2112, result.tempPeakBytes());
            assertEquals(64, result.tempRemainingBytes());
            assertTrue(result.heapPeakBytes() >= result.heapBaselineBytes());
            assertEquals(3, result.samples());
            assertSame(result, probe.stop());
        }
    }

    @Test
    void surfacesFilesystemErrorsInsteadOfReportingAZeroPeak() throws Exception {
        try (var probe = new BenchmarkProbe(root, root.resolve("excluded"), 10, false)) {
            Files.delete(root);
            assertThrows(java.io.IOException.class, probe::stop);
            Files.createDirectory(root); // Allow close to complete after verifying the error.
        }
    }

    @Test
    void rejectsInvalidSamplingInterval() {
        assertThrows(IllegalArgumentException.class, () -> new BenchmarkProbe(root, root, 0, false));
    }
}
