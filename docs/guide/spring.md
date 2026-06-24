# Spring Integration

> [Back to Index](index.md)

## Spring MVC

The `example` module includes `ExcelResponse` and `CsvResponse` helpers that wrap
handlers into `ResponseEntity<StreamingResponseBody>` with proper Content-Type,
Content-Disposition (including RFC 5987 Korean filename encoding), and Cache-Control.

```java
@GetMapping("/download")
public ResponseEntity<StreamingResponseBody> download() {
    ExcelHandler handler = writer.write(dataStream);
    return ExcelResponse.of(handler, "report");
}

@GetMapping("/download-csv")
public ResponseEntity<StreamingResponseBody> downloadCsv() {
    CsvHandler handler = csvWriter.write(dataStream);
    return CsvResponse.of(handler, "report");
}

// Password-encrypted
@GetMapping("/download-encrypted")
public ResponseEntity<StreamingResponseBody> downloadEncrypted() {
    ExcelHandler handler = writer.password("P@ssw0rd!").write(dataStream);
    return ExcelResponse.of(handler, "secret");
}
```

For upload endpoints, return structured read errors when the client asks for
JSON and a readable HTML/text summary for manual testing:

```java
@PostMapping("/upload")
public ResponseEntity<?> upload(MultipartFile file, @RequestHeader(HttpHeaders.ACCEPT) String accept)
        throws IOException {
    List<RowError> errors = new ArrayList<>();
    List<User> rows = new ArrayList<>();

    try (InputStream in = file.getInputStream()) {
        userReader.build(in).read(rows::add, errors::add);
    }

    if (accept.contains(MediaType.APPLICATION_JSON_VALUE)) {
        return ResponseEntity.ok(Map.of("rows", rows, "errors", errors));
    }
    return ResponseEntity.ok("Success: %d rows, Errors: %d rows".formatted(rows.size(), errors.size()));
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
