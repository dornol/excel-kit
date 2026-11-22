package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.schema.ExcelKitSchema;
import io.github.dornol.excelkit.core.read.ReadLimits;
import io.github.dornol.excelkit.core.detect.TabularFileDetector;
import io.github.dornol.excelkit.core.detect.TabularFileType;
import jakarta.validation.Validator;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.util.function.Supplier;

/**
 * Spring upload helpers that open {@link MultipartFile}s and collect read results.
 *
 * @since 0.18.4
 */
public final class ExcelKitUpload {

    private ExcelKitUpload() {
    }

    /**
     * Opens an upload, passes its stream to the reader, and collects the rows and errors it reports.
     *
     * @param type   a label stored in the result, such as {@code "Excel"}
     * @param file   the upload
     * @param reader reads the stream
     * @param <T>    the row type
     * @return the collected result
     * @throws ExcelKitUploadException if the file is empty or cannot be opened or closed
     */
    public static <T> UploadResult<T> read(
            String type,
            MultipartFile file,
            UploadReader<T> reader) {
        try (InputStream inputStream = ExcelKitMultipartFile.open(file)) {
            return UploadResult.read(type, consumer -> reader.read(inputStream, consumer),
                    ExcelKitMultipartFile.safeOriginalFilename(file), file.getSize());
        } catch (IOException e) {
            throw new ExcelKitUploadException(
                    "Failed to close upload file: " + ExcelKitMultipartFile.safeOriginalFilename(file), e);
        }
    }

    /**
     * Like {@link #read(String, MultipartFile, UploadReader)}, keeping at most the rows and errors the
     * limits allow; the counts still include the rest.
     *
     * @param type             a label stored in the result, such as {@code "Excel"}
     * @param file             the upload
     * @param reader           reads the stream
     * @param collectionLimits how many rows and errors to keep
     * @param <T>              the row type
     * @return the collected result
     * @throws ExcelKitUploadException if the file is empty or cannot be opened or closed
     */
    public static <T> UploadResult<T> read(String type, MultipartFile file, UploadReader<T> reader,
            UploadCollectionLimits collectionLimits) {
        try (InputStream inputStream = ExcelKitMultipartFile.open(file)) {
            return UploadResult.read(type, consumer -> reader.read(inputStream, consumer),
                    ExcelKitMultipartFile.safeOriginalFilename(file), file.getSize(), collectionLimits);
        } catch (IOException e) {
            throw new ExcelKitUploadException("Failed to close upload file: "
                    + ExcelKitMultipartFile.safeOriginalFilename(file), e);
        }
    }

    /** Detects content by signature and rejects a mismatch before invoking the reader. */
    public static <T> UploadResult<T> readDetected(String type, TabularFileType expected,
            MultipartFile file, UploadReader<T> reader) {
        try (BufferedInputStream input = new BufferedInputStream(ExcelKitMultipartFile.open(file))) {
            TabularFileType detected = TabularFileDetector.detect(input);
            if (detected != expected) {
                throw new ExcelKitUploadException("Expected " + expected + " content but detected " + detected);
            }
            return UploadResult.read(type, consumer -> reader.read(input, consumer),
                    ExcelKitMultipartFile.safeOriginalFilename(file), file.getSize());
        } catch (IOException e) {
            throw new ExcelKitUploadException("Failed to close upload file: "
                    + ExcelKitMultipartFile.safeOriginalFilename(file), e);
        }
    }

    /**
     * Reads an Excel upload with a reader of your own.
     *
     * @param file   the upload
     * @param reader reads the stream
     * @param <T>    the row type
     * @return the collected result
     */
    public static <T> UploadResult<T> excel(
            MultipartFile file,
            UploadReader<T> reader) {
        return read("Excel", file, reader);
    }

    /**
     * Reads an XLSX upload with a schema, a Bean Validation provider and read limits. The content
     * signature is checked first, so a file that is not XLSX is rejected with {@link
     * ExcelKitUploadException} whatever its name.
     *
     * @param file      the upload
     * @param schema    the columns
     * @param supplier  creates a row object
     * @param validator validates each row; may be null
     * @param limits    bounds on the read
     * @param <T>       the row type
     * @return the collected result
     */
    public static <T> UploadResult<T> excel(MultipartFile file, ExcelKitSchema<T> schema,
            Supplier<T> supplier, Validator validator, ReadLimits limits) {
        return readDetected("Excel", TabularFileType.XLSX, file, (input, consumer) ->
                schema.excelReader(supplier, validator).limits(limits).read(input, consumer));
    }

    /**
     * Reads an XLSX upload with a schema. The content signature is checked first, so a file
     * whose content is not XLSX is rejected with {@link ExcelKitUploadException} regardless of its name.
     * The read is bounded by {@link ReadLimits#UPLOAD}; pass limits of your own to the overload that
     * takes them, or {@link ReadLimits#UNLIMITED} to lift the bounds.
     */
    public static <T> UploadResult<T> excel(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier) {
        return excel(file, schema, supplier, null);
    }

    /** Like the three-argument overload, with a Bean Validation provider; bounded by {@link ReadLimits#UPLOAD}. */
    public static <T> UploadResult<T> excel(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier,
            Validator validator) {
        return excel(file, schema, supplier, validator, ReadLimits.UPLOAD);
    }

    /**
     * Reads a CSV upload with a reader of your own.
     *
     * @param file   the upload
     * @param reader reads the stream
     * @param <T>    the row type
     * @return the collected result
     */
    public static <T> UploadResult<T> csv(
            MultipartFile file,
            UploadReader<T> reader) {
        return read("CSV", file, reader);
    }

    /**
     * Reads a CSV upload with a schema, a Bean Validation provider and read limits. The content
     * signature is checked first, so a file that is not CSV is rejected with {@link
     * ExcelKitUploadException} whatever its name.
     *
     * @param file      the upload
     * @param schema    the columns
     * @param supplier  creates a row object
     * @param validator validates each row; may be null
     * @param limits    bounds on the read
     * @param <T>       the row type
     * @return the collected result
     */
    public static <T> UploadResult<T> csv(MultipartFile file, ExcelKitSchema<T> schema,
            Supplier<T> supplier, Validator validator, ReadLimits limits) {
        return readDetected("CSV", TabularFileType.CSV, file, (input, consumer) ->
                schema.csvReader(supplier, validator).limits(limits).read(input, consumer));
    }

    /**
     * Reads a CSV upload with a schema. The content signature is checked first, so a file
     * whose content is not CSV is rejected with {@link ExcelKitUploadException} regardless of its name.
     * The read is bounded by {@link ReadLimits#UPLOAD}; pass limits of your own to the overload that
     * takes them, or {@link ReadLimits#UNLIMITED} to lift the bounds.
     */
    public static <T> UploadResult<T> csv(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier) {
        return csv(file, schema, supplier, null);
    }

    /** Like the three-argument overload, with a Bean Validation provider; bounded by {@link ReadLimits#UPLOAD}. */
    public static <T> UploadResult<T> csv(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier,
            Validator validator) {
        return csv(file, schema, supplier, validator, ReadLimits.UPLOAD);
    }
}
