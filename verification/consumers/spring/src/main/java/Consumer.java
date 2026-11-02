import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.spring.ExcelKitResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Stream;

public class Consumer {
    public static void main(String[] args) throws Exception {
        try {
            Class.forName("jakarta.validation.Validator");
            throw new AssertionError("Validation must remain optional for Spring downloads");
        } catch (ClassNotFoundException expected) { }
        var handler = ExcelWriter.<String>create().column("Name", value -> value).write(Stream.of("Alice"));
        var response = ExcelKitResponse.excel(handler, "users.xlsx");
        if (!response.getStatusCode().is2xxSuccessful()) throw new AssertionError(response);
        if (!response.getHeaders().getContentDisposition().getFilename().equals("users.xlsx")) {
            throw new AssertionError(response.getHeaders());
        }
        var output = new ByteArrayOutputStream();
        response.getBody().writeTo(output);
        var count = new AtomicInteger();
        ExcelReader.<String>mapping(row -> row.get("Name").asString())
            .read(new ByteArrayInputStream(output.toByteArray()), result -> {
                if (!result.success() || !"Alice".equals(result.data())) throw new AssertionError(result);
                count.incrementAndGet();
            });
        if (count.get() != 1) throw new AssertionError(count);
        System.out.println("Spring consumer passed with transitive excel-kit artifact; JVM " + Runtime.version());
    }
}
