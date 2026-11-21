package io.github.dornol.excelkit.spring;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;

/**
 * Runs a streaming response body, as Spring MVC does after the controller returns.
 * <p>
 * The helpers build their handler when the response is created and only write it when the body runs,
 * so a test that looks at the headers alone leaves the handler (and its temp files) unconsumed.
 */
final class StreamedBody {
    private StreamedBody() {
    }

    static byte[] bytes(ResponseEntity<StreamingResponseBody> response) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        response.getBody().writeTo(out);
        return out.toByteArray();
    }

    static String text(ResponseEntity<StreamingResponseBody> response) throws IOException {
        return new String(bytes(response), StandardCharsets.UTF_8);
    }

    static XSSFWorkbook workbook(ResponseEntity<StreamingResponseBody> response) throws IOException {
        return new XSSFWorkbook(new ByteArrayInputStream(bytes(response)));
    }
}
