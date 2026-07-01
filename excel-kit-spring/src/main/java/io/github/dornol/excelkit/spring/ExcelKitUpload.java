package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.AbstractReadHandler;
import io.github.dornol.excelkit.core.ExcelKitSchema;
import jakarta.validation.Validator;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.function.Function;
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
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        try (InputStream inputStream = ExcelKitMultipartFile.open(file)) {
            return UploadResult.read(type, handlerFactory.apply(inputStream),
                    ExcelKitMultipartFile.safeOriginalFilename(file), file.getSize());
        } catch (IOException e) {
            throw new ExcelKitUploadException(
                    "Failed to close upload file: " + ExcelKitMultipartFile.safeOriginalFilename(file), e);
        }
    }

    public static <T> UploadResult<T> validateExcel(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return excel(file, handlerFactory);
    }

    public static <T> UploadResult<T> validateCsv(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return csv(file, handlerFactory);
    }

    public static <T> UploadResult<T> excel(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return read("Excel", file, handlerFactory);
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
        return excel(file, inputStream -> schema.excelReader(supplier, validator).build(inputStream));
    }

    public static <T> UploadResult<T> csv(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return read("CSV", file, handlerFactory);
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
        return csv(file, inputStream -> schema.csvReader(supplier, validator).build(inputStream));
    }
}
