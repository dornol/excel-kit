package io.github.dornol.excelkit.core;

import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

/** Adapts row-result callbacks to the convenience callback forms exposed by readers. */
final class ReadResultCallbacks {
    private ReadResultCallbacks() {
    }

    static <T> void strict(Consumer<Consumer<ReadResult<T>>> reader, Consumer<T> consumer) {
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(consumer, "consumer cannot be null");
        AtomicLong rowNumber = new AtomicLong();
        reader.accept(result -> {
            long row = rowNumber.incrementAndGet();
            if (!result.success()) {
                String detail = result.messages().isEmpty()
                        ? "Unknown error" : String.join("; ", result.messages());
                throw new ReadAbortException("Row " + row + " read failed: " + detail);
            }
            consumer.accept(result.data());
        });
    }

    static <T> void route(Consumer<Consumer<ReadResult<T>>> reader,
                          Consumer<T> onSuccess, Consumer<RowError> onError) {
        Objects.requireNonNull(reader, "reader cannot be null");
        Objects.requireNonNull(onSuccess, "onSuccess cannot be null");
        Objects.requireNonNull(onError, "onError cannot be null");
        AtomicLong rowNumber = new AtomicLong();
        reader.accept(result -> {
            long row = rowNumber.incrementAndGet();
            if (result.success()) {
                onSuccess.accept(result.data());
                return;
            }
            List<String> messages = result.messages();
            onError.accept(new RowError(row, result.fileRowNum(), result.errorType(), messages,
                    result.cause(), result.cellErrors(), result.rawValues()));
        });
    }
}
