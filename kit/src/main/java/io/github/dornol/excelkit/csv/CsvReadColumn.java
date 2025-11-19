package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.shared.CellData;

import java.io.InputStream;
import java.util.function.BiConsumer;

/**
 * CSV 컬럼 정의를 나타내는 클래스.
 * @param <T> CSV 데이터를 매핑할 객체 타입
 */
public record CsvReadColumn<T>(BiConsumer<T, CellData> setter) {

    /**
     * CsvReader에서 컬럼을 체이닝 방식으로 추가하기 위한 빌더 클래스
     */
    public static class CsvReadColumnBuilder<T> {
        private final CsvReader<T> reader;
        private final BiConsumer<T, CellData> setter;

        CsvReadColumnBuilder(CsvReader<T> reader, BiConsumer<T, CellData> setter) {
            this.reader = reader;
            this.setter = setter;
        }

        /**
         * 다음 컬럼 추가
         * @param setter 객체에 값을 매핑할 함수
         * @return 새로운 컬럼 빌더
         */
        public CsvReadColumnBuilder<T> column(BiConsumer<T, CellData> setter) {
            buildCurrentAndAddToReader();
            return new CsvReadColumn.CsvReadColumnBuilder<>(reader, setter);
        }

        /**
         * CSV 읽기 핸들러 생성
         * @param inputStream CSV 입력 스트림
         * @return CsvReadHandler
         */
        public CsvReadHandler<T> build(InputStream inputStream) {
            buildCurrentAndAddToReader();
            return this.reader.build(inputStream);
        }

        /**
         * 내부: 현재 컬럼 정의를 CsvReader에 추가
         */
        private void buildCurrentAndAddToReader() {
            this.reader.addColumn(new CsvReadColumn<>(this.setter));
        }
    }


}
