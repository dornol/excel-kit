package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.asyncDispatch;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.header;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.request;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/** Exercises the helpers through real MVC request handling, including streaming bodies. */
class ExcelKitMockMvcTest {

    public static class Product {
        public String name;
        public Integer price;
    }

    static final ExcelKitSchema<Product> SCHEMA = ExcelKitSchema.<Product>builder()
            .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
            .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt())
            .build();

    @RestController
    static class TestController {
        private static List<Product> products() {
            Product a = new Product();
            a.name = "Notebook";
            a.price = 1200;
            Product b = new Product();
            b.name = "한글 상품";
            b.price = 300;
            return List.of(a, b);
        }

        @GetMapping("/excel")
        ResponseEntity<StreamingResponseBody> excel() {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(products()), "상품 목록");
        }

        @GetMapping("/excel-encrypted")
        ResponseEntity<StreamingResponseBody> encrypted() {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(products()), "secret", "pw1234");
        }

        @GetMapping("/csv")
        ResponseEntity<StreamingResponseBody> csv() {
            return ExcelKitResponse.csv(SCHEMA.csvWriter().write(products()), "products");
        }

        @PostMapping("/upload/excel")
        UploadResult<Product> uploadExcel(@RequestParam("file") MultipartFile file) {
            return ExcelKitUpload.excel(file, SCHEMA, Product::new);
        }

        @PostMapping("/upload/csv")
        UploadResult<Product> uploadCsv(@RequestParam("file") MultipartFile file) {
            return ExcelKitUpload.csv(file, SCHEMA, Product::new);
        }
    }

    private MockMvc mvc;

    @BeforeEach
    void setUp() {
        mvc = MockMvcBuilders.standaloneSetup(new TestController()).build();
    }

    private byte[] download(String path) throws Exception {
        MvcResult started = mvc.perform(get(path)).andExpect(request().asyncStarted()).andReturn();
        return mvc.perform(asyncDispatch(started)).andExpect(status().isOk()).andReturn()
                .getResponse().getContentAsByteArray();
    }

    @Test
    void excelDownload_streamsReadableWorkbookWithDownloadHeaders() throws Exception {
        MvcResult started = mvc.perform(get("/excel")).andExpect(request().asyncStarted()).andReturn();
        MvcResult done = mvc.perform(asyncDispatch(started))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andReturn();

        String disposition = done.getResponse().getHeader(HttpHeaders.CONTENT_DISPOSITION);
        assertTrue(disposition.startsWith("attachment"), disposition);
        assertTrue(disposition.contains("filename*=UTF-8''"), disposition);

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(done.getResponse().getContentAsByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
            assertEquals("Notebook", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("한글 상품", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals(300, (int) sheet.getRow(2).getCell(1).getNumericCellValue());
        }
    }

    @Test
    void encryptedExcelDownload_isNotAPlainWorkbook() throws Exception {
        byte[] body = download("/excel-encrypted");
        assertTrue(body.length > 0);
        // An encrypted workbook is an OLE2 container, not a ZIP (PK) archive.
        assertTrue((body[0] & 0xFF) == 0xD0 && (body[1] & 0xFF) == 0xCF, "expected OLE2 header");
    }

    @Test
    void csvDownload_streamsRowsWithCsvContentType() throws Exception {
        MvcResult started = mvc.perform(get("/csv")).andExpect(request().asyncStarted()).andReturn();
        MvcResult done = mvc.perform(asyncDispatch(started))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE, "text/csv; charset=UTF-8"))
                .andReturn();

        String csv = new String(done.getResponse().getContentAsByteArray(), StandardCharsets.UTF_8);
        assertTrue(csv.contains("Name"), csv);
        assertTrue(csv.contains("Notebook"), csv);
        assertTrue(csv.contains("한글 상품"), csv);
    }

    @Test
    void excelRoundTrip_downloadedWorkbookCanBeUploaded() throws Exception {
        byte[] xlsx = download("/excel");
        MockMultipartFile file = new MockMultipartFile("file", "products.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, xlsx);

        mvc.perform(multipart("/upload/excel").file(file))
                .andExpect(status().isOk())
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.type").value("Excel"))
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.successCount").value(2))
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.errorCount").value(0));
    }

    @Test
    void csvUpload_reportsRowErrorsAsJson() throws Exception {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nNotebook,1200\nPen,bad\n".getBytes(StandardCharsets.UTF_8));

        mvc.perform(multipart("/upload/csv").file(file))
                .andExpect(status().isOk())
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.successCount").value(1))
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.errorCount").value(1))
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers
                        .jsonPath("$.errors[0].type").exists());
    }

    @Test
    void excelUpload_rejectsCsvContentWithXlsxName_evenWithoutExplicitLimits() {
        MockMultipartFile file = new MockMultipartFile("file", "fake.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, "Name,Price\nA,1\n".getBytes(StandardCharsets.UTF_8));

        Exception error = org.junit.jupiter.api.Assertions.assertThrows(Exception.class,
                () -> mvc.perform(multipart("/upload/excel").file(file)));
        assertTrue(error.getCause() instanceof ExcelKitUploadException, String.valueOf(error));
    }

    @Test
    void csvUploadWithAWrongHeaderExplainsWhichHeaderIsMissing() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Cost\nNotebook,1200\n".getBytes(StandardCharsets.UTF_8));

        Exception error = org.junit.jupiter.api.Assertions.assertThrows(Exception.class,
                () -> mvc.perform(multipart("/upload/csv").file(file)));

        // What a controller advice or an error body reports is getMessage() of the library exception.
        Throwable failure = error.getCause();
        assertTrue(failure instanceof io.github.dornol.excelkit.csv.CsvReadException, String.valueOf(failure));
        assertTrue(failure.getMessage().contains("Header aliases [Price] not found"), failure.getMessage());
        assertTrue(failure.getMessage().contains("Available headers: [Name, Cost]"), failure.getMessage());
    }
}
