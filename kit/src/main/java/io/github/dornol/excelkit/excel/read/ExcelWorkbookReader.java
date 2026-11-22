package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.core.io.TempWorkspace;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.ReadSummary;
import io.github.dornol.excelkit.core.read.RowError;
import org.jspecify.annotations.Nullable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

/**
 * Reads multiple workbook sheets with independent readers and result aggregation.
 *
 * @since 0.23.0
 */
public final class ExcelWorkbookReader {
    private static final int DEFAULT_MAX_ERRORS_PER_SHEET = 100;

    private final List<SheetPlan> plans = new ArrayList<>();
    private SheetFailurePolicy failurePolicy = SheetFailurePolicy.CONTINUE;
    private int maxErrorsPerSheet = DEFAULT_MAX_ERRORS_PER_SHEET;

    private ExcelWorkbookReader() {
    }

    /** Creates a reader with no sheets registered. */
    public static ExcelWorkbookReader create() {
        return new ExcelWorkbookReader();
    }

    /**
     * Sets what happens after a sheet fails; the default is {@link SheetFailurePolicy#CONTINUE}.
     *
     * @param policy whether to carry on with the remaining sheets or stop
     * @return this reader for chaining
     */
    public ExcelWorkbookReader failurePolicy(SheetFailurePolicy policy) {
        this.failurePolicy = Objects.requireNonNull(policy, "policy cannot be null");
        return this;
    }

    /**
     * Sets how many failed rows each {@link SheetReadReport} keeps; the default is 100 and 0 keeps
     * none.
     *
     * @param maximum the most failed rows kept per sheet
     * @return this reader for chaining
     */
    public ExcelWorkbookReader maxErrorsPerSheet(int maximum) {
        if (maximum < 0) throw new IllegalArgumentException("maximum must be non-negative");
        this.maxErrorsPerSheet = maximum;
        return this;
    }

    /**
     * Registers a sheet by name. A name that is not in the workbook is reported as that sheet's
     * failure.
     *
     * @param name     the sheet name
     * @param reader   the reader that maps the sheet's rows
     * @param consumer receives every row's {@link ReadResult}
     * @param <T>      the row type
     * @return this reader for chaining
     */
    public <T> ExcelWorkbookReader sheet(String name, ExcelReader<T> reader,
                                         Consumer<ReadResult<T>> consumer) {
        Objects.requireNonNull(name, "name cannot be null");
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        plans.add(new SheetPlan(name, null, new ReaderSheetExecution<>(reader, consumer)));
        return this;
    }

    /**
     * Registers a sheet by its 0-based index.
     *
     * @param index    the sheet index
     * @param reader   the reader that maps the sheet's rows
     * @param consumer receives every row's {@link ReadResult}
     * @param <T>      the row type
     * @return this reader for chaining
     */
    public <T> ExcelWorkbookReader sheet(int index, ExcelReader<T> reader,
                                         Consumer<ReadResult<T>> consumer) {
        if (index < 0) throw new IllegalArgumentException("index must be non-negative");
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        plans.add(new SheetPlan(null, index, new ReaderSheetExecution<>(reader, consumer)));
        return this;
    }

    /**
     * Reads every registered sheet from a stream, which is copied to a temporary file first and is not
     * closed.
     *
     * @param input the workbook
     * @return one {@link SheetReadReport} per registered sheet
     */
    public WorkbookReadReport read(InputStream input) {
        Objects.requireNonNull(input, "input cannot be null");
        try (TempWorkspace workspace = TempWorkspace.create()) {
            Path copy = workspace.newFile("excel-kit-workbook-", ".xlsx");
            ExcelInputCopy.copy(input, copy, inputByteLimit());
            return read(copy);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to prepare workbook input", e);
        }
    }

    /**
     * The most input bytes any registered reader accepts, or -1 when one of them is unlimited. A file
     * above it would be refused by every sheet anyway, so the copy stops there instead of filling the disk.
     */
    private long inputByteLimit() {
        long limit = 0;
        for (SheetPlan plan : plans) {
            long maximum = plan.execution.maxInputBytes();
            if (maximum < 0) return -1;
            limit = Math.max(limit, maximum);
        }
        return plans.isEmpty() ? -1 : limit;
    }

    /** Like {@link #read(InputStream)}, for a source whose stream is opened and closed here. */
    public WorkbookReadReport read(InputStreamSource source) {
        Objects.requireNonNull(source, "source cannot be null");
        try (InputStream input = source.openStream()) {
            return read(input);
        } catch (IOException e) {
            throw new ExcelKitException("Failed to open workbook input", e);
        }
    }

    /**
     * Reads every registered sheet from a workbook file.
     *
     * @param path the workbook
     * @return one {@link SheetReadReport} per registered sheet
     */
    public WorkbookReadReport read(Path path) {
        Objects.requireNonNull(path, "path cannot be null");
        String password = sharedPassword();
        if (password == null) return readPlain(path);
        // An encrypted workbook is decrypted once here; every sheet then reads the same plain copy.
        try (TempWorkspace workspace = TempWorkspace.create()) {
            Path decrypted = ExcelReadSupport.decryptInto(workspace, path, password);
            plans.forEach(plan -> plan.execution.decryptedInput(true));
            try {
                return readPlain(decrypted);
            } finally {
                plans.forEach(plan -> plan.execution.decryptedInput(false));
            }
        } catch (ExcelKitException e) {
            throw e;
        } catch (Exception e) {
            throw new ExcelKitException("Failed to open workbook", e);
        }
    }

    private WorkbookReadReport readPlain(Path path) {
        List<ExcelSheetInfo> sheets = ExcelWorkbookInspector.create().sheetNames(path);
        List<SheetReadReport> reports = new ArrayList<>();
        // One security scan per policy for the whole file, however many sheets are read from it.
        SecurityScanCache scans = new SecurityScanCache();
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
                            ? plan.readShared(path, index, name, maxErrorsPerSheet, pkg, scans)
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

    /** The password of the registered readers (all sheets of one file share one password), or {@code null}. */
    private @Nullable String sharedPassword() {
        return plans.stream().map(plan -> plan.execution.password()).filter(Objects::nonNull)
                .findFirst().orElse(null);
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
                                                      Path path, int index, int maximumErrors, OPCPackage pkg,
                                                      SecurityScanCache scans) {
        reader.sheetIndex(index);
        List<RowError> errors = new ArrayList<>();
        boolean[] errorsTruncated = {false};
        long[] rowOrdinal = {0};
        ReadSummary summary = reader.readSharedPackage(path, pkg, scans, result -> {
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

        default long maxInputBytes() { return -1; }

        /** Tells the reader that the file it is given is already decrypted, so it must not decrypt it again. */
        default void decryptedInput(boolean decrypted) { }

        default @Nullable String password() { return null; }

        default SheetReadReport readShared(Path path, int index, int maximumErrors, OPCPackage pkg,
                                             SecurityScanCache scans) {
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
        public long maxInputBytes() {
            return reader.maxInputBytes();
        }

        @Override
        public void decryptedInput(boolean decrypted) {
            reader.decryptedInput(decrypted);
        }

        @Override
        public @Nullable String password() {
            return reader.password();
        }

        @Override
        public SheetReadReport readShared(Path path, int index, int maximumErrors, OPCPackage pkg,
                                          SecurityScanCache scans) {
            return executeShared(reader, consumer, path, index, maximumErrors, pkg, scans);
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

        private SheetReadReport readShared(Path path, int index, String name, int maximumErrors, OPCPackage pkg,
                                           SecurityScanCache scans) {
            SheetReadReport result = execution.readShared(path, index, maximumErrors, pkg, scans);
            return new SheetReadReport(name, index, result.summary(), result.errors(),
                    result.errorsTruncated(), result.failure());
        }
    }
}
