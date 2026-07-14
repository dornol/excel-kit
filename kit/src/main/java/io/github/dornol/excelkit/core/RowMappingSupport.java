package io.github.dornol.excelkit.core;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validator;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.function.BiConsumer;
import java.util.function.Function;

/** Setter, row-mapper, and Bean Validation execution shared by format handlers. */
final class RowMappingSupport {
    private static final Logger log = LoggerFactory.getLogger(RowMappingSupport.class);

    private RowMappingSupport() { }

    static <T> boolean validate(T instance, @Nullable Validator validator, List<String> messages) {
        if (validator == null) return true;
        Set<ConstraintViolation<T>> violations = validator.validate(instance);
        violations.stream().map(ConstraintViolation::getMessage).forEach(messages::add);
        return violations.isEmpty();
    }

    static <T> boolean map(BiConsumer<T, CellData> setter, T instance, CellData cell,
                           int index, List<String> headers, List<String> messages,
                           @Nullable List<CellError> cellErrors) {
        try {
            setter.accept(instance, cell);
            return true;
        } catch (Exception exception) {
            String header = header(index, headers);
            String message = "Failed to set column '" + header + "': value='"
                    + cell.formattedValue() + "', reason=" + exception.getMessage();
            messages.add(message);
            if (cellErrors != null) cellErrors.add(new CellError(index, header, cell.formattedValue(), message));
            log.warn("Column mapping failed for '{}': value='{}'", header, cell.formattedValue(), exception);
            return false;
        }
    }

    static <T> boolean map(ReadColumn<T> column, T instance, CellData cell, int index,
                           List<String> headers, List<String> messages,
                           @Nullable List<CellError> cellErrors) {
        if (column.isRequired() && cell.isEmpty()) {
            String header = header(index, headers);
            String message = "Required column '" + header + "' is empty";
            messages.add(message);
            if (cellErrors != null) cellErrors.add(new CellError(index, header, cell.formattedValue(), message));
            return false;
        }
        return map(column.setter(), instance, cell, index, headers, messages, cellErrors);
    }

    static <T> ReadResult<T> mapRow(Function<RowData, T> mapper, RowData rowData,
                                    @Nullable Validator validator, long fileRow,
                                    List<String> rawValues) {
        T instance;
        try {
            instance = mapper.apply(rowData);
        } catch (Exception exception) {
            log.warn("Row mapping failed", exception);
            return new ReadResult<>(null, false,
                    List.of("Row mapping failed: " + exception.getMessage()), exception,
                    fileRow, List.of(), rawValues);
        }
        List<String> messages = new ArrayList<>();
        boolean valid = validate(instance, validator, messages);
        return new ReadResult<>(instance, valid, messages.isEmpty() ? null : messages,
                null, fileRow, List.of(), rawValues);
    }

    private static String header(int index, List<String> headers) {
        return index < headers.size() ? headers.get(index) : "column#" + index;
    }
}
