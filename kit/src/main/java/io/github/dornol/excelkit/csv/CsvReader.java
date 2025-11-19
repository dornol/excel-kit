package io.github.dornol.excelkit.csv;


import io.github.dornol.excelkit.shared.CellData;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

/**
 * CSV Reader. 컬럼 정의와 객체 생성 방법, Validator를 가지고 CSV를 읽어 객체로 변환.
 */
public class CsvReader<T> {
    private final List<CsvReadColumn<T>> columns = new ArrayList<>();
    private final Supplier<T> instanceSupplier;
    private final Validator validator;

    public CsvReader(Supplier<T> instanceSupplier, Validator validator) {
        this.instanceSupplier = instanceSupplier;
        this.validator = validator;
    }

    /**
     * 컬럼 추가
     */
    void addColumn(CsvReadColumn<T> column) {
        columns.add(column);
    }

    /**
     * CSV 컬럼 빌더 시작
     */
    public CsvReadColumn.CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
        return new CsvReadColumn.CsvReadColumnBuilder<>(this, setter);
    }

    /**
     * CsvReadHandler 생성
     */
    public CsvReadHandler<T> build(InputStream inputStream) {
        return new CsvReadHandler<>(inputStream, columns, instanceSupplier, validator);
    }
}
