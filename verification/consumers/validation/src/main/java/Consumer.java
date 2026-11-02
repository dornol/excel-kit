import io.github.dornol.excelkit.csv.CsvReader;
import jakarta.validation.Validation;
import jakarta.validation.constraints.NotBlank;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.concurrent.atomic.AtomicInteger;

public class Consumer {
    public record Person(@NotBlank String name) {}
    public static void main(String[] args) throws Exception {
        try (var factory = Validation.byDefaultProvider().configure()
                .messageInterpolator(new ParameterMessageInterpolator()).buildValidatorFactory()) {
            var failures = new AtomicInteger();
            CsvReader.<Person>mapping(row -> new Person(""))
                .validator(factory.getValidator())
                .read(new ByteArrayInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8)), result -> {
                    if (!result.success()) failures.incrementAndGet();
                });
            if (failures.get() != 1) throw new AssertionError("Validation did not reject the blank name");
        }
        System.out.println("Optional Validation integration passed; JVM " + Runtime.version());
    }
}
