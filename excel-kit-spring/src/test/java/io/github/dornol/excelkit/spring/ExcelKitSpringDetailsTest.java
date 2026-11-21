package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.RowError;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayOutputStream;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Details of the response and upload helpers found by probing: content checks, filenames, caching, CSV injection. */
class ExcelKitSpringDetailsTest {

    private static MockMultipartFile file(String name, byte[] bytes) {
        return new MockMultipartFile("f", name, null, bytes);
    }

    private static void accepts(String name, byte[] bytes) {
        assertDoesNotThrow(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file(name, bytes)), name);
    }

    private static void rejects(String name, byte[] bytes) {
        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(file(name, bytes)), name);
    }

    @Test
    void aWorkbookNeedsTheZipSignatureNotJustTheLettersPk() {
        rejects("a.xlsx", "PK this is a text file, not a workbook".getBytes(StandardCharsets.UTF_8));
        accepts("a.xlsx", new byte[]{'P', 'K', 3, 4, 20, 0, 0, 0});
        accepts("a.xlsm", new byte[]{'P', 'K', 3, 4, 20, 0, 0, 0});
    }

    @Test
    void csvMustBeTextInUtf8() {
        accepts("a.csv", "a,b\n1,2\n".getBytes(StandardCharsets.UTF_8));
        accepts("A.CSV", "﻿a,b\n".getBytes(StandardCharsets.UTF_8));
        rejects("a.csv", "이름,도시\n홍길동,서울\n".getBytes(Charset.forName("EUC-KR")));
        rejects("a.csv", new byte[]{(byte) 0x89, 'P', 'N', 'G', 0, 1});
        rejects("a.csv", new byte[]{'P', 'K', 3, 4, 20, 0, 0, 0});
        rejects("a.csv", new byte[0]);
    }

    @Test
    void theNameDecidesWhichChecksApply() {
        rejects("data", "a\n".getBytes(StandardCharsets.UTF_8));
        rejects("a.csv.exe", "a\n".getBytes(StandardCharsets.UTF_8));
    }

    @Test
    void theAllowedExtensionsAreListedInOneOrder() {
        for (int i = 0; i < 3; i++) {
            var error = assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.requireExtension(file("a.txt", new byte[]{1}), "xlsx", "csv", "xlsm"));
            assertEquals("Unsupported upload file extension. Allowed extensions: csv, xlsm, xlsx", error.getMessage());
        }
    }

    @Test
    void anAsteriskInAFilenameIsEscapedInTheExtendedParameter() {
        String header = DownloadFileType.EXCEL.contentDisposition("a*b");
        assertTrue(header.endsWith("filename*=UTF-8''a%2Ab.xlsx"), header);
    }

    @Test
    void aFilenameCannotInjectHeadersOrBreakTheQuoting() {
        for (String name : List.of("a\r\nX-Evil: 1", "a\"b", "..", "a/b\\c", "한글 보고서", "😀", "x".repeat(300))) {
            String header = DownloadFileType.CSV.contentDisposition(name);
            assertTrue(header.startsWith("attachment; filename=\""), header);
            assertTrue(header.chars().noneMatch(c -> c == '\r' || c == '\n'), header);
            String fallback = header.substring("attachment; filename=\"".length(), header.indexOf("\"; filename*="));
            assertTrue(fallback.chars().allMatch(c -> c >= 0x20 && c < 0x7f && c != '"' && c != '\\' && c != '/'), header);
        }
    }

    @Test
    void downloadsAreCachedPrivately() {
        assertEquals("private, max-age=10",
                ExcelKitResponse.excel("a").build().getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
        assertEquals("private, max-age=0",
                ExcelKitResponse.csv("a", java.time.Duration.ZERO).build().getHeaders().getFirst(HttpHeaders.CACHE_CONTROL));
    }

    @Test
    void theCsvErrorReportDefusesFormulasInEveryColumn() throws Exception {
        var error = new UploadError(1, 2, RowError.Type.MAPPING, List.of("=cmd|' /C calc'!A0"),
                List.of(new CellErrorResponse(0, "=HYPERLINK(\"http://evil\")", "=1+1", "@SUM(1)")),
                List.of("=1+1", "+2", "-3", "@x"));
        var out = new ByteArrayOutputStream();
        ExcelKitErrorResponse.csv(List.of(error), "errors").getBody().writeTo(out);
        String csv = out.toString(StandardCharsets.UTF_8);
        for (String line : csv.split("\r?\n")) {
            for (String cell : line.split(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)")) {
                assertTrue(cell.isEmpty() || !"=+-@".contains(cell.substring(0, 1)) || cell.matches("-?\\d+(\\.\\d+)?"), cell);
            }
        }
        assertTrue(csv.contains("'=1+1"), csv);
    }
}
