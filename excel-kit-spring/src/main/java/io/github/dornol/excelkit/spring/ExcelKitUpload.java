package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.AbstractReadHandler;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.function.Function;

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
            return UploadResult.read(type, handlerFactory.apply(inputStream));
        } catch (IOException e) {
            throw new ExcelKitUploadException(
                    "Failed to close upload file: " + ExcelKitMultipartFile.safeOriginalFilename(file), e);
        }
    }

    public static <T> UploadResult<T> excel(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return read("Excel", file, handlerFactory);
    }

    public static <T> UploadResult<T> csv(
            MultipartFile file,
            Function<InputStream, AbstractReadHandler<T>> handlerFactory) {
        return read("CSV", file, handlerFactory);
    }
}
