package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class DownloadFileTypeTest {

    @Test
    void contentDisposition_doesNotDuplicateExtension() {
        String header = DownloadFileType.EXCEL.contentDisposition("report.xlsx");

        assertTrue(header.contains("filename=\"report.xlsx\""));
        assertTrue(header.contains("filename*=UTF-8''report.xlsx"));
    }

    @Test
    void contentDisposition_appendsMissingExtension() {
        String header = DownloadFileType.CSV.contentDisposition("report");

        assertTrue(header.contains("filename=\"report.csv\""));
        assertTrue(header.contains("filename*=UTF-8''report.csv"));
    }

    @Test
    void contentDisposition_encodesNonAsciiFilenameWithAsciiFallback() {
        String header = DownloadFileType.EXCEL.contentDisposition("도서 목록");

        assertTrue(header.contains("filename=\"__ __.xlsx\""));
        assertTrue(header.contains("filename*=UTF-8''%EB%8F%84%EC%84%9C%20%EB%AA%A9%EB%A1%9D.xlsx"));
    }

    @Test
    void contentType_exposesExpectedMimeTypes() {
        assertEquals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                DownloadFileType.EXCEL.getContentType());
        assertEquals("text/csv; charset=UTF-8", DownloadFileType.CSV.getContentType());
    }
}
