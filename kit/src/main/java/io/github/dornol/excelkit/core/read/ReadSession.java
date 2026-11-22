package io.github.dornol.excelkit.core.read;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.Predicate;

/**
 * One executable reader session. A session owns resources created while opening a path or source.
 * Caller-owned input streams passed to {@code open(InputStream)} remain caller-owned.
 *
 * @since 0.23.0
 */
public interface ReadSession<T> extends AutoCloseable {
    void read(Consumer<ReadResult<T>> consumer);

    void readStrict(Consumer<T> consumer);

    void readWhile(Predicate<ReadResult<T>> predicate);

    default List<T> readAll() {
        List<T> rows = new ArrayList<>();
        readStrict(rows::add);
        return List.copyOf(rows);
    }

    @Override
    void close();
}
