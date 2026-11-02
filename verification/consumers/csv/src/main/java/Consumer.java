import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

public class Consumer {
    public static void main(String[] args) throws Exception {
        absent("org.apache.poi.xssf.usermodel.XSSFWorkbook");
        absent("jakarta.validation.Validator");
        var file = Files.createTempFile("excel-kit-consumer", ".csv");
        try {
            CsvWriter.<String>create().column("Name", value -> value)
                .write(Stream.of("Alice")).writeTo(file);
            var values = new ArrayList<String>();
            CsvReader.<String>mapping(row -> row.get("Name").asString()).read(file, result -> {
                if (!result.success()) throw new AssertionError(result.messages());
                values.add(result.data());
            });
            if (!values.equals(List.of("Alice"))) throw new AssertionError(values);
        } finally {
            Files.deleteIfExists(file);
        }
        System.out.println("CSV-only round trip passed without POI or Validation; JVM " + Runtime.version());
    }
    static void absent(String name) throws Exception {
        try { Class.forName(name); } catch (ClassNotFoundException expected) { return; }
        throw new AssertionError("Unexpected dependency: " + name);
    }
}
