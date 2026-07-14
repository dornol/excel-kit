package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitSchema;
import io.github.dornol.excelkit.core.ReadLimits;
import io.github.dornol.excelkit.core.TabularFileDetector;
import io.github.dornol.excelkit.core.TabularFileType;
import jakarta.validation.Validator;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.BufferedInputStream;
import java.util.function.Supplier;

/**
 * Spring upload helpers that open {@link MultipartFile}s and collect read results.
 */
public final class ExcelKitUpload {

    private ExcelKitUpload() {
    }

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

    public static <T> UploadResult<T> validateExcel(
            MultipartFile file,
            UploadReader<T> reader) {
        return excel(file, reader);
    }

    public static <T> UploadResult<T> validateCsv(
            MultipartFile file,
            UploadReader<T> reader) {
        return csv(file, reader);
    }

    public static <T> UploadResult<T> excel(
            MultipartFile file,
            UploadReader<T> reader) {
        return read("Excel", file, reader);
    }

    public static <T> UploadResult<T> excel(MultipartFile file, ExcelKitSchema<T> schema,
            Supplier<T> supplier, Validator validator, ReadLimits limits) {
        return readDetected("Excel", TabularFileType.XLSX, file, (input, consumer) ->
                schema.excelReader(supplier, validator).limits(limits).read(input, consumer));
    }

    public static <T> UploadResult<T> excel(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier) {
        return excel(file, schema, supplier, null);
    }

    public static <T> UploadResult<T> excel(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier,
            Validator validator) {
        return excel(file, (inputStream, consumer) -> schema.excelReader(supplier, validator).read(inputStream, consumer));
    }

    public static <T> UploadResult<T> csv(
            MultipartFile file,
            UploadReader<T> reader) {
        return read("CSV", file, reader);
    }

    public static <T> UploadResult<T> csv(MultipartFile file, ExcelKitSchema<T> schema,
            Supplier<T> supplier, Validator validator, ReadLimits limits) {
        return readDetected("CSV", TabularFileType.CSV, file, (input, consumer) ->
                schema.csvReader(supplier, validator).limits(limits).read(input, consumer));
    }

    public static <T> UploadResult<T> csv(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier) {
        return csv(file, schema, supplier, null);
    }

    public static <T> UploadResult<T> csv(
            MultipartFile file,
            ExcelKitSchema<T> schema,
            Supplier<T> supplier,
            Validator validator) {
        return csv(file, (inputStream, consumer) -> schema.csvReader(supplier, validator).read(inputStream, consumer));
    }
}
