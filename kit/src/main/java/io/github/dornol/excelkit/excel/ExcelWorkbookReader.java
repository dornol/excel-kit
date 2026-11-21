package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.InputStreamSource;
import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.ReadSummary;
import io.github.dornol.excelkit.core.RowError;
import org.jspecify.annotations.Nullable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

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
        plans.add(new SheetPlan(name, null, new ReaderSheetExecution<>(reader, consumer)));
        return this;
    }

    public <T> ExcelWorkbookReader sheet(int index, ExcelReader<T> reader,
                                         Consumer<ReadResult<T>> consumer) {
        if (index < 0) throw new IllegalArgumentException("index must be non-negative");
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        plans.add(new SheetPlan(null, index, new ReaderSheetExecution<>(reader, consumer)));
        return this;
    }

    public WorkbookReadReport read(InputStream input) {
        Objects.requireNonNull(input, "input cannot be null");
        try {
            Path temp = Files.createTempFile("excel-kit-workbook-", ".xlsx");
            try {
                Files.copy(input, temp, StandardCopyOption.REPLACE_EXISTING);
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
        List<ExcelSheetInfo> sheets = sheetNames(path);
        List<SheetReadReport> reports = new ArrayList<>();
        boolean sharePackage = !plans.isEmpty() && plans.stream().allMatch(SheetPlan::canSharePackage);
        try (OPCPackage pkg = sharePackage ? OPCPackage.open(path.toFile(), PackageAccess.READ) : null) {
            for (SheetPlan plan : plans) {
                Integer resolved = plan.index != null ? plan.index : findSheetIndex(plan.name, sheets);
                if (resolved == null) {
                    // A missing named sheet is a per-sheet failure, like a missing index.
                    reports.add(failure(String.valueOf(plan.name), -1,
                            new ExcelKitException("Sheet '" + plan.name + "' is not present")));
                    if (failurePolicy == SheetFailurePolicy.STOP) break;
                    continue;
                }
                int index = resolved;
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
                    reports.add(sharePackage
                            ? plan.readShared(path, index, name, maxErrorsPerSheet, pkg)
                            : plan.read(path, index, name, maxErrorsPerSheet));
                } catch (RuntimeException e) {
                    reports.add(failure(name, index, e));
                    if (failurePolicy == SheetFailurePolicy.STOP) break;
                }
            }
        } catch (Exception e) {
            throw new ExcelKitException("Failed to open workbook", e);
        }
        return new WorkbookReadReport(reports);
    }

    /**
     * The sheets of the workbook. An encrypted workbook is listed with the password of the readers that were
     * registered for it (all sheets of one file share one password).
     */
    private List<ExcelSheetInfo> sheetNames(Path path) {
        String password = plans.stream().map(plan -> plan.execution.password()).filter(Objects::nonNull)
                .findFirst().orElse(null);
        return password == null ? ExcelReader.getSheetNames(path) : ExcelReader.getSheetNames(path, password);
    }

    private @Nullable Integer findSheetIndex(@Nullable String name, List<ExcelSheetInfo> sheets) {
        return sheets.stream().filter(sheet -> sheet.name().equals(name)).map(ExcelSheetInfo::index)
                .findFirst().orElse(null);
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

    private static <T> SheetReadReport executeShared(ExcelReader<T> reader, Consumer<ReadResult<T>> consumer,
                                                      Path path, int index, int maximumErrors, OPCPackage pkg) {
        reader.sheetIndex(index);
        List<RowError> errors = new ArrayList<>();
        boolean[] errorsTruncated = {false};
        long[] rowOrdinal = {0};
        ReadSummary summary = reader.readSharedPackage(path, pkg, result -> {
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

        default boolean canSharePackage() { return false; }

        default @Nullable String password() { return null; }

        default SheetReadReport readShared(Path path, int index, int maximumErrors, OPCPackage pkg) {
            return read(path, index, maximumErrors);
        }
    }

    private static final class ReaderSheetExecution<T> implements SheetExecution {
        private final ExcelReader<T> reader;
        private final Consumer<ReadResult<T>> consumer;

        private ReaderSheetExecution(ExcelReader<T> reader, Consumer<ReadResult<T>> consumer) {
            this.reader = reader;
            this.consumer = consumer;
        }

        @Override
        public SheetReadReport read(Path path, int index, int maximumErrors) {
            return execute(reader, consumer, path, index, maximumErrors);
        }

        @Override
        public boolean canSharePackage() {
            return reader.canSharePackage();
        }

        @Override
        public @Nullable String password() {
            return reader.password();
        }

        @Override
        public SheetReadReport readShared(Path path, int index, int maximumErrors, OPCPackage pkg) {
            return executeShared(reader, consumer, path, index, maximumErrors, pkg);
        }
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

        private boolean canSharePackage() { return execution.canSharePackage(); }

        private SheetReadReport readShared(Path path, int index, String name, int maximumErrors, OPCPackage pkg) {
            SheetReadReport result = execution.readShared(path, index, maximumErrors, pkg);
            return new SheetReadReport(name, index, result.summary(), result.errors(),
                    result.errorsTruncated(), result.failure());
        }
    }
}
