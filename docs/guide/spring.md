# Spring Integration

> [Back to Index](index.md)

## Spring MVC

Add the optional Spring helper module:

```kotlin
implementation("io.github.dornol:excel-kit-spring:<version>")
```

`ExcelKitResponse` wraps handlers into `ResponseEntity<StreamingResponseBody>`
with proper Content-Type, Content-Disposition (including RFC 5987 Korean
filename encoding), and Cache-Control.

```java
@GetMapping("/download")
public ResponseEntity<StreamingResponseBody> download() {
    ExcelHandler handler = writer.write(dataStream);
    return ExcelKitResponse.excel(handler, "report");
}

@GetMapping("/download-csv")
public ResponseEntity<StreamingResponseBody> downloadCsv() {
    CsvHandler handler = csvWriter.write(dataStream);
    return ExcelKitResponse.csv(handler, "report");
}

// Password-encrypted
@GetMapping("/download-encrypted")
public ResponseEntity<StreamingResponseBody> downloadEncrypted() {
    ExcelHandler handler = writer.password("P@ssw0rd!").write(dataStream);
    return ExcelKitResponse.excel(handler, "secret");
}
```

For upload endpoints, return structured read errors when the client asks for
JSON and a readable HTML/text summary for manual testing:

```java
@PostMapping("/upload")
public ResponseEntity<UploadResult<User>> upload(MultipartFile file) {
    UploadResult<User> result = ExcelKitUpload.excel(
        file, (in, consumer) -> userReader.read(in, consumer));
    return ResponseEntity.ok(result);
}
```

Validate size and filename extension before opening uploads:

```java
@PostMapping("/upload")
public ResponseEntity<UploadResult<User>> upload(MultipartFile file) {
    MultipartFile checked = ExcelKitMultipartFile.requireSizeAtMost(
        ExcelKitMultipartFile.requireExcelOrCsvContent(file), 5 * 1024 * 1024);

    return ResponseEntity.ok(ExcelKitUpload.excel(
        checked, (in, consumer) -> userReader.read(in, consumer)));
}
```

For schema-backed uploads, pass the schema directly:

```java
@PostMapping("/upload")
public ResponseEntity<UploadResult<User>> upload(MultipartFile file) {
    MultipartFile checked = ExcelKitMultipartFile.requireExcelOrCsvContent(file);
    return ResponseEntity.ok(ExcelKitUpload.excel(checked, userSchema, User::new, validator));
}
```

`UploadResult.summary()` includes total rows, success/error counts, duration,
filename, and file size. `UploadError.rawValues()` carries the source row values
for correction screens.

When users need a downloadable correction report, reuse the same upload parse
path and convert the structured errors to CSV or Excel:

```java
@PostMapping("/upload/errors.csv")
public ResponseEntity<StreamingResponseBody> errorReport(MultipartFile file) {
    UploadResult<User> result = ExcelKitUpload.excel(
        file, (in, consumer) -> userReader.read(in, consumer));
    return ExcelKitErrorResponse.csv(result, "read-errors");
}
```

Schema-based upload templates can be streamed empty or with sample rows:

```java
@GetMapping("/template.xlsx")
public ResponseEntity<StreamingResponseBody> excelTemplate() {
    return ExcelKitTemplateResponse.excel(
        userSchema, "users-template", List.of(new User("Alice", 30)));
}

@GetMapping("/template-guided.xlsx")
public ResponseEntity<StreamingResponseBody> guidedTemplate() {
    return ExcelKitTemplateResponse.excelWithGuidance(userSchema, "users-template");
}

@GetMapping("/template.csv")
public ResponseEntity<StreamingResponseBody> csvTemplate() {
    return ExcelKitTemplateResponse.csv(userSchema, "users-template");
}
```

### Late-Binding Password

When the service layer builds the handler but the password is only known
at the presentation layer:

```java
// Service
public ExcelHandler buildReport() {
    return writer.column(...).write(dataStream);
}

// Controller
ExcelHandler handler = service.buildReport();
handler.writeTo(outputStream, "P@ssw0rd!");
```

## Spring WebFlux

Apache POI is blocking I/O — wrap on `boundedElastic`:

```java
@GetMapping("/download")
public Mono<Void> download(ServerHttpResponse response) {
    response.getHeaders().setContentType(
        MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
    response.getHeaders().set(HttpHeaders.CONTENT_DISPOSITION,
        "attachment; filename=\"report.xlsx\"");

    return response.writeWith(DataBufferUtils.readInputStream(
        () -> {
            PipedInputStream pis = new PipedInputStream();
            PipedOutputStream pos = new PipedOutputStream(pis);
            Schedulers.boundedElastic().schedule(() -> {
                try {
                    writer.write(dataStream).writeTo(pos);
                    pos.close();
                } catch (IOException e) {
                    throw new UncheckedIOException(e);
                }
            });
            return pis;
        },
        response.bufferFactory(), 8192));
}
```

With reactive repositories:
```java
Flux<MyData> flux = repository.findAll();
ExcelHandler handler = writer.write(flux.toStream());
// Flux.toStream() handles backpressure — not loaded entirely into memory.
```
