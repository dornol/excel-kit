package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.ExcelKitSchema;
import io.github.dornol.excelkit.core.InputStreamSource;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.ReadSummary;
import io.github.dornol.excelkit.core.RowError;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Supplier;

/** Reads multiple workbook sheets with independent readers and result aggregation. */
public final class ExcelWorkbookReader {
    private static final int DEFAULT_MAX_ERRORS_PER_SHEET = 100;

    private final List<SheetPlan> plans = new ArrayList<>();
    private SheetFailurePolicy failurePolicy = SheetFailurePolicy.CONTINUE;
    private int maxErrorsPerSheet = DEFAULT_MAX_ERRORS_PER_SHEET;

    private ExcelWorkbookReader() {
    }

    public static ExcelWorkbookReader create() {
        return new ExcelWorkbookReader();
    }

    public ExcelWorkbookReader failurePolicy(SheetFailurePolicy policy) {
        this.failurePolicy = Objects.requireNonNull(policy, "policy cannot be null");
        return this;
    }

    public ExcelWorkbookReader maxErrorsPerSheet(int maximum) {
        if (maximum < 0) throw new IllegalArgumentException("maximum must be non-negative");
        this.maxErrorsPerSheet = maximum;
        return this;
    }

    public <T> ExcelWorkbookReader sheet(String name, ExcelReader<T> reader,
                                         Consumer<ReadResult<T>> consumer) {
        Objects.requireNonNull(name, "name cannot be null");
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        plans.add(new SheetPlan(name, null, (path, index, errors) -> execute(reader, consumer, path, index, errors)));
        return this;
    }

    public <T> ExcelWorkbookReader sheet(int index, ExcelReader<T> reader,
                                         Consumer<ReadResult<T>> consumer) {
        if (index < 0) throw new IllegalArgumentException("index must be non-negative");
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        plans.add(new SheetPlan(null, index, (path, resolved, errors) ->
                execute(reader, consumer, path, resolved, errors)));
        return this;
    }

    public <T> ExcelWorkbookReader sheet(String name, ExcelKitSchema<T> schema,
                                         Supplier<T> supplier, @Nullable Validator validator,
                                         Consumer<ReadResult<T>> consumer) {
        Objects.requireNonNull(schema, "schema cannot be null");
        return sheet(name, schema.excelReader(supplier, validator), consumer);
    }

    public <T> ExcelWorkbookReader sheet(int index, ExcelKitSchema<T> schema,
                                         Supplier<T> supplier, @Nullable Validator validator,
                                         Consumer<ReadResult<T>> consumer) {
        Objects.requireNonNull(schema, "schema cannot be null");
        return sheet(index, schema.excelReader(supplier, validator), consumer);
    }

    public WorkbookReadReport read(InputStream input) {
        Objects.requireNonNull(input, "input cannot be null");
        try {
            Path temp = Files.createTempFile("excel-kit-workbook-", ".xlsx");
            try {
                Files.copy(input, temp, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                return read(temp);
            } finally {
                Files.deleteIfExists(temp);
            }
        } catch (IOException e) {
            throw new ExcelKitException("Failed to prepare workbook input", e);
        }
    }

    public WorkbookReadReport read(InputStreamSource source) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return read(input);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to open workbook input", e);
        }
    }

    public WorkbookReadReport read(Path path) {
        Objects.requireNonNull(path, "path cannot be null");
        List<ExcelSheetInfo> sheets = ExcelReader.getSheetNames(path);
        List<SheetReadReport> reports = new ArrayList<>();
        for (SheetPlan plan : plans) {
            int index = plan.index != null ? plan.index : resolveSheetIndex(plan.name, sheets);
            String name = sheets.stream().filter(sheet -> sheet.index() == index)
                    .map(ExcelSheetInfo::name).findFirst().orElse("#" + index);
            if (index >= sheets.size()) {
                SheetReadReport report = failure(name, index,
                        new ExcelKitException("Sheet index " + index + " is not present"));
                reports.add(report);
                if (failurePolicy == SheetFailurePolicy.STOP) break;
                continue;
            }
            try {
                reports.add(plan.read(path, index, name, maxErrorsPerSheet));
            } catch (RuntimeException e) {
                reports.add(failure(name, index, e));
                if (failurePolicy == SheetFailurePolicy.STOP) break;
            }
        }
        return new WorkbookReadReport(reports);
    }

    private int resolveSheetIndex(@Nullable String name, List<ExcelSheetInfo> sheets) {
        return sheets.stream().filter(sheet -> sheet.name().equals(name)).map(ExcelSheetInfo::index)
                .findFirst().orElseThrow(() -> new ExcelKitException("Sheet '" + name + "' is not present"));
    }

    private SheetReadReport failure(String name, int index, RuntimeException failure) {
        return new SheetReadReport(name, index, new ReadSummary(0, 0, 0, false, Duration.ZERO),
                List.of(), false, failure);
    }

    private static <T> SheetReadReport execute(ExcelReader<T> reader, Consumer<ReadResult<T>> consumer,
                                                Path path, int index, int maximumErrors) {
        reader.sheetIndex(index);
        List<RowError> errors = new ArrayList<>();
        boolean[] errorsTruncated = {false};
        long[] rowOrdinal = {0};
        ReadSummary summary = reader.readWithSummary(path, result -> {
            rowOrdinal[0]++;
            consumer.accept(result);
            if (!result.success()) {
                if (errors.size() < maximumErrors) {
                    errors.add(new RowError(rowOrdinal[0], result.fileRowNum(), result.errorType(), result.messages(),
                            result.cause(), result.cellErrors(), result.rawValues()));
                } else {
                    errorsTruncated[0] = true;
                }
            }
        });
        return new SheetReadReport("#" + index, index, summary, errors, errorsTruncated[0], null);
    }

    @FunctionalInterface
    private interface SheetExecution {
        SheetReadReport read(Path path, int index, int maximumErrors);
    }

    private static final class SheetPlan {
        private final @Nullable String name;
        private final @Nullable Integer index;
        private final SheetExecution execution;

        private SheetPlan(@Nullable String name, @Nullable Integer index, SheetExecution execution) {
            this.name = name;
            this.index = index;
            this.execution = execution;
        }

        private SheetReadReport read(Path path, int index, String name, int maximumErrors) {
            SheetReadReport result = execution.read(path, index, maximumErrors);
            return new SheetReadReport(name, index, result.summary(), result.errors(),
                    result.errorsTruncated(), result.failure());
        }
    }
}
