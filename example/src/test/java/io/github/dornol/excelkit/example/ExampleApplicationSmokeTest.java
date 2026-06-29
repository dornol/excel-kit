package io.github.dornol.excelkit.example;

import org.apache.poi.ss.usermodel.DataFormatter;
import io.github.dornol.excelkit.example.app.book.adapter.out.persistence.BookMyBatisMapper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.webmvc.test.autoconfigure.AutoConfigureMockMvc;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.asyncDispatch;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.content;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.header;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.request;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@AutoConfigureMockMvc
@SpringBootTest(properties = {
        "spring.docker.compose.enabled=false",
        "spring.datasource.url=jdbc:h2:mem:example-smoke;MODE=MariaDB;DB_CLOSE_DELAY=-1;DATABASE_TO_LOWER=TRUE",
        "spring.datasource.driver-class-name=org.h2.Driver",
        "spring.datasource.username=sa",
        "spring.datasource.password=",
        "spring.jpa.hibernate.ddl-auto=create-drop",
        "spring.jpa.show-sql=false",
        "example.dummy-count=0"
})
class ExampleApplicationSmokeTest {

    @Autowired
    private BookMyBatisMapper bookMyBatisMapper;

    @Autowired
    private MockMvc mockMvc;

    @Test
    void contextLoads() {
        assertEquals(0, bookMyBatisMapper.countBooks());
    }

    @Test
    void downloadExcel_returnsWorkbookResponse() throws Exception {
        MvcResult result = mockMvc.perform(get("/download-excel"))
                .andExpect(request().asyncStarted())
                .andReturn();

        MvcResult response = mockMvc.perform(asyncDispatch(result))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andExpect(header().string(HttpHeaders.CONTENT_DISPOSITION,
                        org.hamcrest.Matchers.containsString("book list excel.xlsx")))
                .andReturn();

        byte[] body = response.getResponse().getContentAsByteArray();
        assertTrue(body.length > 4);
        assertEquals('P', body[0]);
        assertEquals('K', body[1]);
    }

    @Test
    void indexPage_containsInlineUploadResultPanel() throws Exception {
        mockMvc.perform(get("/index.html"))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("id=\"upload-result\"")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("data-enhanced-upload=\"true\"")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("href=\"/app.css\"")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("src=\"/app.js\"")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Download Error CSV")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Server Error Excel")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Download Excel Template")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Download CSV Template")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("feature-filter")));
    }

    @Test
    void staticAssets_areServed() throws Exception {
        mockMvc.perform(get("/app.css"))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("#upload-result")));

        mockMvc.perform(get("/app.js"))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("downloadServerErrorReport")));
    }

    @Test
    void uploadShowcaseExcel_readsRows() throws Exception {
        MockMultipartFile file = new MockMultipartFile("file", "products.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", productWorkbook());

        mockMvc.perform(multipart("/showcase/read-by-name-excel").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Success: 2 rows")));
    }

    @Test
    void uploadShowcaseCsv_readsRows() throws Exception {
        String csv = "Name,Category,Price,Quantity,Discount\n"
                + "Notebook,Stationery,1200,3,0.1\n"
                + "Pen,Stationery,500,10,0.0\n";
        MockMultipartFile file = new MockMultipartFile("file", "products.csv",
                "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/showcase/read-by-name-csv").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Success: 2 rows")));
    }

    @Test
    void uploadShowcaseCsv_reportsCellErrors() throws Exception {
        String csv = "Name,Category,Price,Quantity,Discount\n"
                + "Notebook,Stationery,not-a-number,3,0.1\n";
        MockMultipartFile file = new MockMultipartFile("file", "products.csv",
                "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/showcase/read-by-name-csv").file(file))
                .andExpect(status().isOk())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Errors: 1 rows")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("fileRow=2")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("header=Price")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("value=not-a-number")));
    }

    @Test
    void uploadShowcaseCsv_canReturnStructuredJsonErrors() throws Exception {
        String csv = "Name,Category,Price,Quantity,Discount\n"
                + "Notebook,Stationery,not-a-number,3,0.1\n";
        MockMultipartFile file = new MockMultipartFile("file", "products.csv",
                "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/showcase/read-by-name-csv").file(file)
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(content().contentTypeCompatibleWith(MediaType.APPLICATION_JSON))
                .andExpect(jsonPath("$.type").value("CSV"))
                .andExpect(jsonPath("$.successCount").value(0))
                .andExpect(jsonPath("$.errorCount").value(1))
                .andExpect(jsonPath("$.errors[0].fileRowNum").value(2))
                .andExpect(jsonPath("$.errors[0].cellErrors[0].headerName").value("Price"))
                .andExpect(jsonPath("$.errors[0].cellErrors[0].cellValue").value("not-a-number"));
    }

    @Test
    void uploadShowcaseExcel_canReturnHtmlResult() throws Exception {
        MockMultipartFile file = new MockMultipartFile("file", "products.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", productWorkbook());

        mockMvc.perform(multipart("/showcase/read-by-name-excel").file(file)
                        .accept(MediaType.TEXT_HTML))
                .andExpect(status().isOk())
                .andExpect(content().contentTypeCompatibleWith(MediaType.TEXT_HTML))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("<h1>Name-Based Excel Read Result</h1>")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("<td>Notebook</td>")));
    }

    @Test
    void uploadShowcaseErrors_canDownloadServerCsvReport() throws Exception {
        String csv = "Name,Category,Price,Quantity,Discount\n"
                + "Notebook,Stationery,not-a-number,3,0.1\n";
        MockMultipartFile file = new MockMultipartFile("file", "products.csv",
                "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        mockMvc.perform(multipart("/showcase/read-errors-csv").file(file))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE,
                        org.hamcrest.Matchers.containsString("text/csv")))
                .andExpect(header().string(HttpHeaders.CONTENT_DISPOSITION,
                        org.hamcrest.Matchers.containsString("read-errors.csv")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("rowNum,fileRowNum,columnIndex,headerName")))
                .andExpect(content().string(org.hamcrest.Matchers.containsString("not-a-number")));
    }

    @Test
    void uploadShowcaseErrors_canDownloadServerExcelReport() throws Exception {
        String csv = "Name,Category,Price,Quantity,Discount\n"
                + "Notebook,Stationery,not-a-number,3,0.1\n";
        MockMultipartFile file = new MockMultipartFile("file", "products.csv",
                "text/csv", csv.getBytes(StandardCharsets.UTF_8));

        MvcResult result = mockMvc.perform(multipart("/showcase/read-errors-excel").file(file))
                .andExpect(request().asyncStarted())
                .andReturn();

        MvcResult response = mockMvc.perform(asyncDispatch(result))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andExpect(header().string(HttpHeaders.CONTENT_DISPOSITION,
                        org.hamcrest.Matchers.containsString("read-errors.xlsx")))
                .andReturn();

        byte[] body = response.getResponse().getContentAsByteArray();
        assertTrue(body.length > 4);
        assertEquals('P', body[0]);
        assertEquals('K', body[1]);

        try (Workbook workbook = new XSSFWorkbook(new ByteArrayInputStream(body))) {
            DataFormatter formatter = new DataFormatter();
            Sheet sheet = workbook.getSheet("Read Errors");
            assertEquals("rowNum", formatter.formatCellValue(sheet.getRow(0).getCell(0)));
            assertEquals("fileRowNum", formatter.formatCellValue(sheet.getRow(0).getCell(1)));
            assertEquals("columnIndex", formatter.formatCellValue(sheet.getRow(0).getCell(2)));
            assertEquals("headerName", formatter.formatCellValue(sheet.getRow(0).getCell(3)));
            assertEquals("cellValue", formatter.formatCellValue(sheet.getRow(0).getCell(4)));
            assertEquals("message", formatter.formatCellValue(sheet.getRow(0).getCell(5)));
            assertEquals("1", formatter.formatCellValue(sheet.getRow(1).getCell(0)));
            assertEquals("2", formatter.formatCellValue(sheet.getRow(1).getCell(1)));
            assertEquals("Price", formatter.formatCellValue(sheet.getRow(1).getCell(3)));
            assertEquals("not-a-number", formatter.formatCellValue(sheet.getRow(1).getCell(4)));
        }
    }

    private static byte[] productWorkbook() throws Exception {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Products");
            Row header = sheet.createRow(0);
            String[] headers = {"Name", "Category", "Price", "Quantity", "Discount"};
            for (int i = 0; i < headers.length; i++) {
                header.createCell(i).setCellValue(headers[i]);
            }
            Row first = sheet.createRow(1);
            first.createCell(0).setCellValue("Notebook");
            first.createCell(1).setCellValue("Stationery");
            first.createCell(2).setCellValue(1200);
            first.createCell(3).setCellValue(3);
            first.createCell(4).setCellValue(0.1);
            Row second = sheet.createRow(2);
            second.createCell(0).setCellValue("Pen");
            second.createCell(1).setCellValue("Stationery");
            second.createCell(2).setCellValue(500);
            second.createCell(3).setCellValue(10);
            second.createCell(4).setCellValue(0.0);
            workbook.write(out);
            return out.toByteArray();
        }
    }
}
