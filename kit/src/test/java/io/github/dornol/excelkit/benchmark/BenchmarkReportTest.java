package io.github.dornol.excelkit.benchmark;

import com.opencsv.CSVReader;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.Files;
import java.nio.file.Path;
import static org.junit.jupiter.api.Assertions.*;

class BenchmarkReportTest {
    @TempDir Path directory;

    @Test
    void reportCanBeParsedWithEscapedFieldsAndASingleHeader() throws Exception {
        Path report = directory.resolve("nested/resources.csv");
        var snapshot = new BenchmarkProbe.Snapshot(100, 500, 0, 2048, 0, 12, 10);
        String scenario = "quoted,\"scenario\"\nnext line";
        BenchmarkReport.append(report, scenario, 1000, 123456, 512, snapshot);
        BenchmarkReport.append(report, "second", 1000, 123456, 512, snapshot);
        try (var reader = new CSVReader(Files.newBufferedReader(report))) {
            var rows = reader.readAll();
            assertEquals(3, rows.size());
            assertEquals("schema_version", rows.get(0)[0]);
            assertEquals(rows.get(0).length, rows.get(1).length);
            assertEquals(rows.get(0).length, rows.get(2).length);
            assertEquals(scenario, rows.get(1)[1]);
            assertEquals("123456", rows.get(1)[3]);
            assertEquals("500", rows.get(1)[6]);
            assertEquals("2048", rows.get(1)[8]);
            assertEquals("second", rows.get(2)[1]);
        }
    }
}
