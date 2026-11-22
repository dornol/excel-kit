package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.excel.ExcelWriteException;
import io.github.dornol.excelkit.schema.ExcelKitSchema;
import jakarta.servlet.MultipartConfigElement;
import org.apache.catalina.Context;
import org.apache.catalina.Wrapper;
import org.apache.catalina.startup.Tomcat;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.context.support.AnnotationConfigWebApplicationContext;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartResolver;
import org.springframework.web.multipart.support.StandardServletMultipartResolver;
import org.springframework.web.servlet.DispatcherServlet;
import org.springframework.web.servlet.config.annotation.EnableWebMvc;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.InetAddress;
import java.net.Socket;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.IntStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The same helpers behind a real servlet container (embedded Tomcat) and a real socket, for what MockMvc cannot
 * show: a client that goes away while the body is streamed, and a multipart upload that has to be parsed by the
 * container. The temp-file guard fails any test that leaves a file behind.
 */
class RealContainerTest {

    public static class Item {
        public Integer id;
        public String text;
    }

    static final ExcelKitSchema<Item> SCHEMA = ExcelKitSchema.<Item>builder()
            .column("Id", i -> i.id, (i, c) -> i.id = c.asInt())
            .column("Text", i -> i.text, (i, c) -> i.text = c.asString())
            .build();

    /** What happened to the last streamed body: filled when the container is done with it. */
    static final AtomicReference<Throwable> BODY_OUTCOME = new AtomicReference<>();
    static volatile CountDownLatch bodyDone = new CountDownLatch(1);

    private static Item item(int id) {
        Item i = new Item();
        i.id = id;
        // hard to compress, so a big download really is big on the wire
        i.text = Long.toHexString(id * 0x9E3779B97F4A7C15L) + Long.toHexString(id * 0xC2B2AE3D27D4EB4FL);
        return i;
    }

    @RestController
    static class Controller {
        @GetMapping("/download")
        ResponseEntity<StreamingResponseBody> download(@RequestParam("rows") int rows) {
            var response = ExcelKitResponse.excel(SCHEMA.excelWriter().write(
                    IntStream.range(0, rows).mapToObj(RealContainerTest::item)), "export");
            StreamingResponseBody original = response.getBody();
            StreamingResponseBody observed = out -> {
                try {
                    original.writeTo(out);
                    BODY_OUTCOME.set(null);
                } catch (Throwable t) {
                    BODY_OUTCOME.set(t);
                    throw t;
                } finally {
                    bodyDone.countDown();
                }
            };
            return ResponseEntity.status(response.getStatusCode()).headers(response.getHeaders()).body(observed);
        }

        @GetMapping("/build-fails")
        ResponseEntity<StreamingResponseBody> buildFails() {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().write(IntStream.range(0, 100).mapToObj(i -> {
                if (i == 50) throw new IllegalStateException("source failed at " + i);
                return item(i);
            })), "x");
        }

        @GetMapping("/password-twice")
        ResponseEntity<StreamingResponseBody> passwordTwice() {
            return ExcelKitResponse.excel(SCHEMA.excelWriter().password("a").write(java.util.List.of(item(1))), "x", "b");
        }

        @PostMapping("/upload")
        String upload(@RequestParam("file") MultipartFile file) {
            UploadResult<Item> result = ExcelKitUpload.excel(file, SCHEMA, Item::new);
            return result.successCount() + "/" + result.errorCount();
        }
    }

    @RestControllerAdvice
    static class Advice {
        @ExceptionHandler(IllegalStateException.class)
        ResponseEntity<String> conflict(IllegalStateException e) {
            return ResponseEntity.status(HttpStatus.CONFLICT).body("conflict");
        }

        @ExceptionHandler(ExcelKitException.class)
        ResponseEntity<String> library(ExcelKitException e) {
            return ResponseEntity.status(HttpStatus.UNPROCESSABLE_ENTITY).body(e.getClass().getSimpleName());
        }
    }

    @Configuration
    @EnableWebMvc
    static class Config {
        @Bean Controller controller() { return new Controller(); }
        @Bean Advice advice() { return new Advice(); }
        @Bean MultipartResolver multipartResolver() { return new StandardServletMultipartResolver(); }
    }

    @TempDir
    static Path baseDir;
    static Tomcat tomcat;
    static int port;
    static final HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(10)).build();

    @BeforeAll
    static void startContainer() throws Exception {
        tomcat = new Tomcat();
        tomcat.setBaseDir(baseDir.toString());
        tomcat.setPort(0);
        tomcat.getConnector();
        Context context = tomcat.addContext("", baseDir.toString());
        AnnotationConfigWebApplicationContext spring = new AnnotationConfigWebApplicationContext();
        spring.register(Config.class);
        Wrapper wrapper = Tomcat.addServlet(context, "dispatcher", new DispatcherServlet(spring));
        wrapper.setAsyncSupported(true);
        wrapper.setLoadOnStartup(1);
        wrapper.setMultipartConfigElement(new MultipartConfigElement(baseDir.toString()));
        context.addServletMappingDecoded("/", "dispatcher");
        tomcat.start();
        port = tomcat.getConnector().getLocalPort();
    }

    @AfterAll
    static void stopContainer() throws Exception {
        if (tomcat != null) {
            tomcat.stop();
            tomcat.destroy();
        }
    }

    private static HttpResponse<byte[]> get(String path) throws Exception {
        return client.send(HttpRequest.newBuilder(URI.create("http://localhost:" + port + path)).GET().build(),
                HttpResponse.BodyHandlers.ofByteArray());
    }

    // ---- a normal download over a real connection ---------------------------------------------------------

    @Test
    void aDownloadArrivesCompleteAndReadableOverARealConnection() throws Exception {
        HttpResponse<byte[]> response = get("/download?rows=3000");

        assertEquals(200, response.statusCode());
        assertTrue(response.headers().firstValue("Content-Disposition").orElse("").contains("export"));
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(response.body()))) {
            assertEquals(3000, wb.getSheetAt(0).getLastRowNum());
            assertEquals(2999, (int) wb.getSheetAt(0).getRow(3000).getCell(0).getNumericCellValue());
        }
    }

    // ---- failures before the response starts, through the container's error handling ------------------------

    @Test
    void aFailureWhileBuildingIsAnOrdinaryErrorStatusOnTheWire() throws Exception {
        HttpResponse<byte[]> response = get("/build-fails");

        assertEquals(422, response.statusCode());
        assertEquals("ExcelWriteException", new String(response.body(), StandardCharsets.UTF_8));
    }

    @Test
    void aPasswordSetTwiceIsAConflictOnTheWireNotABrokenDownload() throws Exception {
        HttpResponse<byte[]> response = get("/password-twice");

        assertEquals(409, response.statusCode());
        assertEquals("conflict", new String(response.body(), StandardCharsets.UTF_8));
    }

    // ---- the client goes away while the file is streamed ----------------------------------------------------

    @Test
    void aClientThatClosesTheConnectionMidDownloadEndsTheBodyWithAnIoErrorAndTheServerKeepsServing() throws Exception {
        bodyDone = new CountDownLatch(1);
        BODY_OUTCOME.set(null);

        try (Socket socket = new Socket(InetAddress.getLoopbackAddress(), port)) {
            socket.setReceiveBufferSize(4096); // make the server's send block early
            OutputStream out = socket.getOutputStream();
            out.write(("GET /download?rows=400000 HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n")
                    .getBytes(StandardCharsets.US_ASCII));
            out.flush();
            InputStream in = socket.getInputStream();
            byte[] firstBytes = in.readNBytes(2048);
            assertTrue(new String(firstBytes, StandardCharsets.US_ASCII).startsWith("HTTP/1.1 200"),
                    "the download started");
            socket.setSoLinger(true, 0); // reset instead of a polite close, as a crashed browser tab does
        }

        assertTrue(bodyDone.await(60, TimeUnit.SECONDS), "the container finished with the body");
        Throwable outcome = BODY_OUTCOME.get();
        assertNotNull(outcome, "the body must not look successful when the client left");
        assertInstanceOf(ExcelWriteException.class, outcome);
        assertInstanceOf(java.io.IOException.class, outcome.getCause(), String.valueOf(outcome.getCause()));

        // the container and the library are fine afterwards
        HttpResponse<byte[]> next = get("/download?rows=20");
        assertEquals(200, next.statusCode());
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(next.body()))) {
            assertEquals(20, wb.getSheetAt(0).getLastRowNum());
        }
    }

    // ---- multipart uploads parsed by the container ----------------------------------------------------------

    private static HttpResponse<byte[]> upload(byte[] file, String filename) throws Exception {
        String boundary = "----excelkit" + System.nanoTime();
        ByteArrayOutputStream body = new ByteArrayOutputStream();
        body.write(("--" + boundary + "\r\nContent-Disposition: form-data; name=\"file\"; filename=\"" + filename
                + "\"\r\nContent-Type: application/octet-stream\r\n\r\n").getBytes(StandardCharsets.UTF_8));
        body.write(file);
        body.write(("\r\n--" + boundary + "--\r\n").getBytes(StandardCharsets.UTF_8));
        return client.send(HttpRequest.newBuilder(URI.create("http://localhost:" + port + "/upload"))
                        .header("Content-Type", "multipart/form-data; boundary=" + boundary)
                        .POST(HttpRequest.BodyPublishers.ofByteArray(body.toByteArray())).build(),
                HttpResponse.BodyHandlers.ofByteArray());
    }

    @Test
    void aDownloadedWorkbookCanBeUploadedBackThroughTheContainer() throws Exception {
        byte[] xlsx = get("/download?rows=5000").body();

        HttpResponse<byte[]> response = upload(xlsx, "back.xlsx");

        assertEquals(200, response.statusCode());
        assertEquals("5000/0", new String(response.body(), StandardCharsets.UTF_8));
    }

    @Test
    void aZipBombUploadIsRefusedWithAnErrorStatus() throws Exception {
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

        HttpResponse<byte[]> response = upload(bomb.toByteArray(), "bomb.xlsx");

        assertEquals(422, response.statusCode(), new String(response.body(), StandardCharsets.UTF_8));
    }

    @Test
    void aTextFileCalledXlsxIsRefusedWithAnErrorStatus() throws Exception {
        HttpResponse<byte[]> response = upload("Id,Text\n1,a\n".getBytes(StandardCharsets.UTF_8), "fake.xlsx");

        assertEquals(422, response.statusCode(), new String(response.body(), StandardCharsets.UTF_8));
        Files.list(baseDir).forEach(p -> { /* the container's own scratch files are not ours to judge */ });
    }
}
