package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

public class CsvReader<T> {
    private final List<CsvReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;

    public CsvReader(Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = instanceSupplier;
        this.validator = validator;
    }

    void addColumn(CsvReadColumn<T> column) {
        columns.add(column);
    }

    public CsvReadColumn.CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new CsvReadColumn.CsvReadColumnBuilder<>(this, setter);
    }

    public CsvReadHandler<T> build(InputStream inputStream) {
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator);
    }
}
