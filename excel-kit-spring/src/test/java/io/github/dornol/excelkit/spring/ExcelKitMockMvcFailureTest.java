package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelWriteException;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.asyncDispatch;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.request;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

/**
 * What a client and a {@code @ControllerAdvice} actually see when something goes wrong around a download or an
 * upload. The temp-file guard fails any test that leaves a file behind, so each case also proves the cleanup.
 */
class ExcelKitMockMvcFailureTest {

    public static class Item {
        public Integer id;
        public String secret;
    }

    static final ExcelKitSchema<Item> SCHEMA = ExcelKitSchema.<Item>builder()
            .column("Id", i -> i.id, (i, c) -> i.id = c.asInt())
            .column("Secret", i -> i.secret, (i, c) -> i.secret = c.asString())
            .build();

    private static Item item(int id) {
        Item i = new Item();
        i.id = id;
        i.secret = "secret-value-" + id;
        return i;
    }

    /** Set by the "download" endpoints so a test can run the body the way the servlet container would. */
    static ResponseEntity<StreamingResponseBody> lastResponse;

    @RestController
    static class Controller {
        @GetMapping("/fail-while-building")
        ResponseEntity<StreamingResponseBody> failWhileBuilding() {
            Stream<Item> rows = IntStream.range(0, 1000).mapToObj(i -> {
                if (i == 500) throw new IllegalStateException("source went away at row " + i);
                return item(i);
            });
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(rows), "x");
        }

        @GetMapping("/conflict")
        ResponseEntity<StreamingResponseBody> conflict() {
            ExcelHandler handler = SCHEMA.excelWriter().password("writer-pw").write(List.of(item(1)));
            return ExcelKitResponse.excel(handler, "x", "response-pw");
        }

        @GetMapping("/download")
        ResponseEntity<StreamingResponseBody> download(@RequestParam(name = "rows", defaultValue = "10") int rows) {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(
                    IntStream.range(0, rows).mapToObj(ExcelKitMockMvcFailureTest::item)), "x");
        }

        @GetMapping("/download-encrypted")
        ResponseEntity<StreamingResponseBody> downloadEncrypted() {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(
                    IntStream.range(0, 1000).mapToObj(ExcelKitMockMvcFailureTest::item)), "x", "pw");
        }

        @PostMapping("/upload")
        UploadResult<Item> upload(@RequestParam("file") MultipartFile file) {
            return ExcelKitUpload.excel(file, SCHEMA, Item::new);
        }
    }

    @RestControllerAdvice
    static class Advice {
        @ExceptionHandler(IllegalStateException.class)
        ResponseEntity<String> conflict(IllegalStateException e) {
            return ResponseEntity.status(HttpStatus.CONFLICT).body(e.getMessage());
        }

        @ExceptionHandler(ExcelKitException.class)
        ResponseEntity<String> library(ExcelKitException e) {
            return ResponseEntity.status(HttpStatus.UNPROCESSABLE_ENTITY).body(e.getClass().getSimpleName());
        }
    }

    private MockMvc mvc;

    @BeforeEach
    void setUp() {
        mvc = MockMvcBuilders.standaloneSetup(new Controller()).setControllerAdvice(new Advice()).build();
    }

    private byte[] download(MockMvc mvc, String path) throws Exception {
        MvcResult started = mvc.perform(get(path)).andExpect(request().asyncStarted()).andReturn();
        return mvc.perform(asyncDispatch(started)).andExpect(status().isOk()).andReturn()
                .getResponse().getContentAsByteArray();
    }

    // ---- failures before the response starts -------------------------------------------------------------

    @Test
    void aSourceThatFailsWhileTheWorkbookIsBuiltBecomesAnOrdinaryErrorResponse() throws Exception {
        // The handler is built in the controller, so the failure happens before any byte is sent: the advice
        // can still answer with a status instead of a half-written file.
        MvcResult result = mvc.perform(get("/fail-while-building")).andReturn();

        assertFalse(result.getRequest().isAsyncStarted(), "no streaming should have started");
        // The library wraps the source's exception; the advice answers for the wrapper, with no row data in it.
        assertEquals(422, result.getResponse().getStatus());
        assertEquals("ExcelWriteException", result.getResponse().getContentAsString());
    }

    @Test
    void theWrappedSourceExceptionKeepsItsCauseForTheServerLog() throws Exception {
        ExcelWriteException failure = assertThrows(ExcelWriteException.class,
                () -> new Controller().failWhileBuilding());

        assertInstanceOf(IllegalStateException.class, failure.getCause());
        assertEquals("source went away at row 500", failure.getCause().getMessage());
    }

    @Test
    void aPasswordSetTwiceIsRejectedBeforeTheDownloadStarts() throws Exception {
        MvcResult result = mvc.perform(get("/conflict")).andReturn();

        assertFalse(result.getRequest().isAsyncStarted(), "the conflict must surface before streaming");
        assertEquals(409, result.getResponse().getStatus());
        assertTrue(result.getResponse().getContentAsString().contains("password"),
                result.getResponse().getContentAsString());
    }

    // ---- failures while the body streams (the client went away) ------------------------------------------

    private static final class FailingOutput extends OutputStream {
        private final int allowed;
        private int written;

        FailingOutput(int allowed) {
            this.allowed = allowed;
        }

        @Override
        public void write(int b) throws IOException {
            write(new byte[]{(byte) b}, 0, 1);
        }

        @Override
        public void write(byte[] b, int off, int len) throws IOException {
            if (written + len > allowed) throw new IOException("Broken pipe");
            written += len;
        }
    }

    @Test
    void aClientThatDisconnectsMidDownloadSurfacesAsAnExcelWriteException() throws Exception {
        ResponseEntity<StreamingResponseBody> response = new Controller()
                .download(5_000);

        ExcelWriteException failure = assertThrows(ExcelWriteException.class,
                () -> response.getBody().writeTo(new FailingOutput(2_000)));

        assertInstanceOf(IOException.class, failure.getCause());
        assertEquals("Broken pipe", failure.getCause().getMessage());
        // The disconnect is the container's problem, but the temp files must not outlive it (the guard checks).
    }

    @Test
    void aClientThatDisconnectsMidEncryptedDownloadSurfacesTheSameWay() throws Exception {
        ResponseEntity<StreamingResponseBody> response = new Controller()
                .downloadEncrypted();

        ExcelWriteException failure = assertThrows(ExcelWriteException.class,
                () -> response.getBody().writeTo(new FailingOutput(2_000)));

        assertInstanceOf(IOException.class, failure.getCause());
    }

    @Test
    void aBodyCannotBeStreamedTwice() throws Exception {
        ResponseEntity<StreamingResponseBody> response = new Controller()
                .download(10);
        response.getBody().writeTo(OutputStream.nullOutputStream());

        // e.g. a container retry after the first attempt: it must fail loudly, not send an empty/partial file
        assertThrows(RuntimeException.class, () -> response.getBody().writeTo(new ByteArrayOutputStream()));
    }

    // ---- uploads that must be refused --------------------------------------------------------------------

    @Test
    void aZipBombIsRefusedAsAnUploadAndDoesNotReachTheController() throws Exception {
        ByteArrayOutputStream bomb = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(bomb)) {
            zip.putNextEntry(new ZipEntry("[Content_Types].xml"));
            zip.write("<Types/>".getBytes(StandardCharsets.UTF_8));
            zip.closeEntry();
            zip.putNextEntry(new ZipEntry("xl/worksheets/sheet1.xml"));
            byte[] chunk = new byte[1 << 20];
            java.util.Arrays.fill(chunk, (byte) 'a');
            for (int i = 0; i < 300; i++) zip.write(chunk);
            zip.closeEntry();
        }
        MockMultipartFile file = new MockMultipartFile("file", "bomb.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, bomb.toByteArray());

        MvcResult result = mvc.perform(multipart("/upload").file(file)).andReturn();

        assertEquals(422, result.getResponse().getStatus(), result.getResponse().getContentAsString());
        assertFalse(result.getResponse().getContentAsString().contains("secret"));
    }

    @Test
    void aTruncatedWorkbookIsRefusedWithALibraryException() throws Exception {
        byte[] whole = download(mvc, "/download?rows=200");
        byte[] cut = java.util.Arrays.copyOf(whole, whole.length / 2);
        MockMultipartFile file = new MockMultipartFile("file", "cut.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, cut);

        MvcResult result = mvc.perform(multipart("/upload").file(file)).andReturn();

        assertEquals(422, result.getResponse().getStatus(), result.getResponse().getContentAsString());
    }

    @Test
    void anEmptyUploadIsRefusedWithALibraryException() throws Exception {
        MockMultipartFile file = new MockMultipartFile("file", "empty.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, new byte[0]);

        MvcResult result = mvc.perform(multipart("/upload").file(file)).andReturn();

        assertEquals(422, result.getResponse().getStatus(), result.getResponse().getContentAsString());
    }

    // ---- volume and concurrency --------------------------------------------------------------------------

    @Test
    void aLargeDownloadCanBeUploadedAgainWithEveryRowCounted() throws Exception {
        byte[] xlsx = download(mvc, "/download?rows=60000");
        MockMultipartFile file = new MockMultipartFile("file", "big.xlsx",
                MediaType.APPLICATION_OCTET_STREAM_VALUE, xlsx);

        mvc.perform(multipart("/upload").file(file))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.successCount").value(60000))
                .andExpect(jsonPath("$.errorCount").value(0));
    }

    @Test
    void concurrentDownloadsDoNotMixTheirRows() throws Exception {
        int clients = 8;
        ExecutorService pool = Executors.newFixedThreadPool(clients);
        try {
            List<Future<Integer>> futures = new ArrayList<>();
            for (int c = 0; c < clients; c++) {
                final int rows = 100 + c * 37;
                Callable<Integer> job = () -> {
                    try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(
                            download(mvc, "/download?rows=" + rows)))) {
                        var sheet = wb.getSheetAt(0);
                        for (int i = 0; i < rows; i++) {
                            assertEquals(i, (int) sheet.getRow(i + 1).getCell(0).getNumericCellValue());
                            assertEquals("secret-value-" + i, sheet.getRow(i + 1).getCell(1).getStringCellValue());
                        }
                        return sheet.getLastRowNum();
                    }
                };
                futures.add(pool.submit(job));
            }
            for (int c = 0; c < clients; c++) {
                assertEquals(100 + c * 37, futures.get(c).get());
            }
        } finally {
            pool.shutdownNow();
        }
    }
}
