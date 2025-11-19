package io.github.dornol.excelkit.csv;

import com.opencsv.CSVReader;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import io.github.dornol.excelkit.shared.TempResourceContainer;
import io.github.dornol.excelkit.shared.TempResourceCreator;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.*;
import java.util.function.Consumer;
import java.util.function.Supplier;

/**
 * CSV를 실제로 읽고 객체로 매핑하는 핸들러
 * 임시 파일에 저장 후 OpenCSV로 읽음
 */
public class CsvReadHandler<T> extends TempResourceContainer {
    private static final Logger log = LoggerFactory.getLogger(CsvReadHandler.class);
    private final List<String> headerNames = new ArrayList<>();
    private final List<CsvReadColumn<T>> columns;
    private final Supplier<T> instanceSupplier;
    private final Validator validator;

    CsvReadHandler(InputStream inputStream, List<CsvReadColumn<T>> columns, Supplier<T> instanceSupplier, Validator validator) {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        if (columns == null || columns.isEmpty()) {
            throw new IllegalArgumentException("Columns cannot be null or empty");
        }
        if (instanceSupplier == null) {
            throw new IllegalArgumentException("Instance supplier cannot be null");
        }
        this.columns = columns;
        this.instanceSupplier = instanceSupplier;
        this.validator = validator;
        try {
            setTempDir(TempResourceCreator.createTempDirectory());
            setTempFile(TempResourceCreator.createTempFile(getTempDir(), UUID.randomUUID().toString(), ".csv"));
            try (InputStream is = inputStream) {
                Files.copy(is, getTempFile(), StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }
    }

    /**
     * 헤더 이름 초기화
     */
    private void prepareColumnHeaders(String[] line) {
        Collections.addAll(headerNames, line);
    }

    /**
     * Validator가 존재하면 객체 유효성 검사 수행
     */
    private boolean validateIfNeeded(T currentInstance, List<String> messages) {
        if (validator == null) {
            return true;
        }

        Set<ConstraintViolation<T>> violations = validator.validate(currentInstance);
        if (violations.isEmpty()) return true;

        violations.stream()
                .map(ConstraintViolation::getMessage)
                .forEach(messages::add);

        return false;
    }

    /**
     * CSV 읽기
     * @param consumer ReadResult를 받아 처리하는 콜백
     */
    public void read(Consumer<ReadResult<T>> consumer) {
        try (CSVReader reader = new CSVReader(new FileReader(getTempFile().toFile()))) {
            String[] line;

            T currentInstance;

            prepareColumnHeaders(reader.readNext());

            while ((line = reader.readNext()) != null) {
                currentInstance = instanceSupplier.get();
                boolean success = true;
                List<String> messages = new ArrayList<>();

                for (int i = 0; i < columns.size(); i++) {
                    try {
                        String columnValue = null;

                        if (i < line.length) {
                            columnValue = line[i];
                        }

                        columns.get(i).setter().accept(currentInstance, new CellData(i, columnValue));
                    } catch (Exception e) {
                        success = false;
                        String header = (i < headerNames.size()) ? headerNames.get(i) : "column#" + i;
                        messages.add("Failed to set column: " + header);
                        log.warn("Column mapping failed", e);
                    }
                }
                boolean validationSuccess = success && validateIfNeeded(currentInstance, messages);

                consumer.accept(new ReadResult<>(currentInstance, validationSuccess, messages));
            }
        } catch (Exception e) {
            throw new IllegalStateException("Failed to read excel", e);
        } finally {
            close();
        }
    }

}
