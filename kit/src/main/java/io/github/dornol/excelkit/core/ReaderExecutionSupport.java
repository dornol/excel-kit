package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.BiFunction;
import java.util.function.BooleanSupplier;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/** Internal execution and result aggregation support shared by all readers. */
final class ReaderExecutionSupport {
    private ReaderExecutionSupport() {
    }

    static <R> R executeFromSource(InputStreamSource source, Function<InputStream, R> execution,
                                   BiFunction<String, IOException, ? extends RuntimeException> exceptionFactory) {
        Objects.requireNonNull(source, "source cannot be null");
        Objects.requireNonNull(execution, "execution cannot be null");
        Objects.requireNonNull(exceptionFactory, "exceptionFactory cannot be null");
        try (InputStream input = source.openStream()) {
            return execution.apply(input);
        } catch (IOException e) {
            throw exceptionFactory.apply("Failed to open input", e);
        }
    }

    static <H, R> R executeWithInputStream(InputStream inputStream,
                                           Function<InputStream, H> handlerFactory,
                                           Function<H, R> execution) {
        Objects.requireNonNull(inputStream, "inputStream cannot be null");
        Objects.requireNonNull(handlerFactory, "handlerFactory cannot be null");
        Objects.requireNonNull(execution, "execution cannot be null");
        return execution.apply(handlerFactory.apply(inputStream));
    }

    static <H> void executeInputStream(InputStream inputStream,
                                       Function<InputStream, H> handlerFactory,
                                       Consumer<H> execution) {
        executeWithInputStream(inputStream, handlerFactory, handler -> {
            execution.accept(handler);
            return null;
        });
    }

    static <H, R> R executeWithPath(Path path, Function<Path, H> handlerFactory,
                                    Function<H, R> execution) {
        Objects.requireNonNull(path, "path cannot be null");
        Objects.requireNonNull(handlerFactory, "handlerFactory cannot be null");
        Objects.requireNonNull(execution, "execution cannot be null");
        return execution.apply(handlerFactory.apply(path));
    }

    static <H> void executePath(Path path, Function<Path, H> handlerFactory,
                                Consumer<H> execution) {
        executeWithPath(path, handlerFactory, handler -> {
            execution.accept(handler);
            return null;
        });
    }

    static <T> ReadSummary summarizeRead(Consumer<Consumer<ReadResult<T>>> execution,
                                         BooleanSupplier stoppedEarly,
                                         Consumer<ReadResult<T>> consumer) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        execution.accept(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            consumer.accept(result);
        });
        return summary(counts, stoppedEarly, started);
    }

    static <T> ReadReport collectReadReport(Consumer<Consumer<ReadResult<T>>> execution,
                                            BooleanSupplier stoppedEarly, int maximum) {
        if (maximum < 0) throw new IllegalArgumentException("maxCollectedErrors must be non-negative");
        List<RowError> errors = new ArrayList<>();
        long[] row = {0};
        ReadSummary summary = summarizeRead(execution, stoppedEarly, result -> {
            row[0]++;
            if (!result.success() && errors.size() < maximum) errors.add(new RowError(row[0],
                    result.fileRowNum(), result.errorType(), result.messages(), result.cause(),
                    result.cellErrors(), result.rawValues()));
        });
        return new ReadReport(summary, errors, summary.errorRows() > errors.size());
    }

    static <T> ReadSummary summarizeReadWhile(Consumer<Predicate<ReadResult<T>>> execution,
                                              BooleanSupplier stoppedEarly,
                                              Predicate<ReadResult<T>> predicate) {
        long started = System.nanoTime();
        long[] counts = new long[3];
        execution.accept(result -> {
            counts[0]++;
            if (result.success()) counts[1]++; else counts[2]++;
            return predicate.test(result);
        });
        return summary(counts, stoppedEarly, started);
    }

    private static ReadSummary summary(long[] counts, BooleanSupplier stoppedEarly, long started) {
        return new ReadSummary(counts[0], counts[1], counts[2], stoppedEarly.getAsBoolean(),
                Duration.ofNanos(System.nanoTime() - started));
    }
}
