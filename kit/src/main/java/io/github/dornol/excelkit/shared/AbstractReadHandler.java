package io.github.dornol.excelkit.shared;

import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.Set;
import java.util.UUID;
import java.util.function.Supplier;

/**
 * Abstract base class for file read handlers (Excel, CSV).
 * <p>
 * Provides common functionality including:
 * <ul>
 *     <li>Constructor parameter validation</li>
 *     <li>Temporary file initialization from an InputStream</li>
 *     <li>Optional Bean Validation support</li>
 * </ul>
 *
 * @param <T> The target row data type to map each row into
 * @author dhkim
 * @since 2025-07-19
 */
public abstract class AbstractReadHandler<T> extends TempResourceContainer {
    private static final Logger log = LoggerFactory.getLogger(AbstractReadHandler.class);

    protected final Supplier<T> instanceSupplier;
    protected final Validator validator;

    /**
     * Constructs a read handler by validating inputs and initializing a temporary file.
     *
     * @param inputStream      The input stream of the uploaded file
     * @param instanceSupplier A supplier to instantiate new row objects
     * @param validator        Optional bean validator for validating mapped instances
     * @param extension        File extension for the temporary file (e.g., ".xlsx", ".csv")
     */
    protected AbstractReadHandler(InputStream inputStream, Supplier<T> instanceSupplier, Validator validator, String extension) {
        if (inputStream == null) {
            throw new IllegalArgumentException("InputStream cannot be null");
        }
        if (instanceSupplier == null) {
            throw new IllegalArgumentException("Instance supplier cannot be null");
        }
        this.instanceSupplier = instanceSupplier;
        this.validator = validator;
        initTempFile(inputStream, extension);
    }

    private void initTempFile(InputStream inputStream, String extension) {
        try {
            setTempDir(TempResourceCreator.createTempDirectory());
            setTempFile(TempResourceCreator.createTempFile(getTempDir(), UUID.randomUUID().toString(), extension));
            try (InputStream is = inputStream) {
                Files.copy(is, getTempFile(), StandardCopyOption.REPLACE_EXISTING);
            }
        } catch (IOException e) {
            throw new ExcelKitException("Failed to initialize temporary file", e);
        }
    }

    /**
     * Validates the given instance using Bean Validation (if a validator is configured).
     *
     * @param instance The object to validate
     * @param messages A mutable list to collect violation messages
     * @return {@code true} if valid or no validator is configured, {@code false} if violations exist
     */
    protected boolean validateIfNeeded(T instance, List<String> messages) {
        if (validator == null) {
            return true;
        }

        Set<ConstraintViolation<T>> violations = validator.validate(instance);
        if (violations.isEmpty()) return true;

        violations.stream()
                .map(ConstraintViolation::getMessage)
                .forEach(messages::add);

        return false;
    }

    /**
     * Maps a single column value to the instance, handling exceptions.
     *
     * @param setter      The setter to apply
     * @param instance    The target object
     * @param cellData    The cell data to set
     * @param columnIndex The column index (for error reporting)
     * @param headerNames The header names (for error reporting)
     * @param messages    A mutable list to collect error messages
     * @return {@code true} if mapping succeeded, {@code false} if an exception occurred
     */
    protected boolean mapColumn(java.util.function.BiConsumer<T, CellData> setter, T instance, CellData cellData,
                                int columnIndex, List<String> headerNames, List<String> messages) {
        try {
            setter.accept(instance, cellData);
            return true;
        } catch (Exception e) {
            String header = (columnIndex < headerNames.size()) ? headerNames.get(columnIndex) : "column#" + columnIndex;
            messages.add("Failed to set column '" + header + "': value='" + cellData.formattedValue() + "', reason=" + e.getMessage());
            log.warn("Column mapping failed for '{}': value='{}'", header, cellData.formattedValue(), e);
            return false;
        }
    }
}
