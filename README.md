# excel-kit

Apache POI 기반의 대용량 Excel(.xlsx) 및 CSV 생성/파싱을 간단하고 안전하게 처리하기 위한 경량 유틸리티입니다. 스트리밍 방식으로 동작하여 메모리 사용을 최소화하고, 컬럼 정의를 플루언트 DSL 스타일로 작성할 수 있습니다. 또한 Excel 비밀번호 암호화 출력과 Bean Validation(선택) 기반의 읽기 검증을 지원합니다.

- 그룹/아티팩트: `io.github.dornol:excel-kit`
- 라이선스: MIT

## 주요 기능
- Excel 쓰기 (SXSSFWorkbook 스트리밍 사용)
  - 컬럼별 타입/포맷/정렬 지정
  - 시트 행 수 자동 분할(최대 행수 지정)
  - 헤더 스타일 색상 지정
  - 데이터 행 높이 설정
  - 결과물을 OutputStream으로 한 번에 소비(consume-once)
  - Excel 파일 비밀번호 암호화 출력 지원
- Excel 읽기 (SAX 기반 스트리밍 파싱)
  - 헤더 자동 인식 및 컬럼 매핑 DSL
  - 헤더 행 인덱스 지정 (메타데이터가 있는 파일 지원)
  - Bean Validation(선택) 통합으로 행 단위 유효성 검증 결과 제공
  - 행별 파싱 결과/메시지 전달
  - 멀티시트 읽기 지원 (시트 인덱스 지정)
- CSV 쓰기
  - 임시 파일로 스트리밍 작성 후 OutputStream으로 전달
  - CSV 기본 이스케이프 처리(인용부호/콤마/개행)
  - UTF-8 BOM 포함으로 Excel에서 한글 깨짐 방지
  - 커서(Cursor) 제공: 현재 행/총 행 등 정보 활용 가능
- CSV 읽기 (OpenCSV 기반)
  - 헤더 자동 인식 및 컬럼 매핑 DSL (BOM 자동 제거)
  - Bean Validation(선택) 통합으로 행 단위 유효성 검증 결과 제공

## 설치
Gradle(Kotlin DSL)
```kotlin
dependencies {
    implementation("io.github.dornol:excel-kit:<latest-version>")
}
```

Maven
```xml
<dependency>
  <groupId>io.github.dornol</groupId>
  <artifactId>excel-kit</artifactId>
  <version><!-- latest-version --></version>
</dependency>
```

### 런타임 의존성 안내
`build.gradle.kts`는 다음을 `compileOnly`로 선언합니다. 실제 실행 시에는 환경에 맞게 의존성을 제공해야 합니다.
- Apache POI: `org.apache.poi:poi-ooxml`
- SLF4J API: `org.slf4j:slf4j-api`
- Jakarta Bean Validation API(선택): `jakarta.validation:jakarta.validation-api`
- OpenCSV(CSV 읽기 시 필요): `com.opencsv:opencsv`

테스트 또는 실행 예제에서처럼 동작시키려면 적절한 구현체를 추가하세요.
- SLF4J 구현: 예) `org.slf4j:slf4j-simple`
- Bean Validation 구현: 예) `org.hibernate:hibernate-validator`

## 빠른 시작

### 1) Excel 쓰기
```java
import io.github.dornol.excelkit.excel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

// 예시 DTO
record Person(long id, String name, int age) {}

// 데이터 스트림 (예시)
var stream = java.util.stream.Stream.of(
        new Person(1, "Alice", 30),
        new Person(2, "Bob", 28)
);

try (ExcelWriter<Person> writer = new ExcelWriter<>(255, 255, 255, 1_000_000)) {
    ExcelHandler handler = writer
            .rowHeight(25) // 데이터 행 높이 설정 (기본 20pt)
            .column("ID", p -> p.id())
                .type(ExcelDataType.LONG)
                .alignment(HorizontalAlignment.RIGHT)
            .column("Name", p -> p.name())
                .type(ExcelDataType.STRING)
            .column("Age", (p, c) -> p.age()) // 커서(c)도 전달 가능
                .type(ExcelDataType.INTEGER)
                .format("0")
            .write(stream); // 마지막에 write 호출

    // 출력(한 번만 가능)
    try (var os = java.nio.file.Files.newOutputStream(java.nio.file.Path.of("people.xlsx"))) {
        handler.consumeOutputStream(os);
    }
}
```

- `ExcelWriter`는 `AutoCloseable`을 구현합니다. try-with-resources 사용을 권장합니다.
- `ExcelWriter#column(name, function)`으로 컬럼을 정의하고, 체이닝으로 `type/format/alignment/style` 등을 지정합니다.
- `write(Stream<T>)` 또는 `write(Stream<T>, ExcelConsumer<T>)`를 호출하면 `ExcelHandler`가 반환됩니다.
- `ExcelHandler#consumeOutputStream(OutputStream)`은 한 번만 호출할 수 있습니다. 재호출 시 `ExcelWriteException`이 발생합니다.

헤더 색상, 시트 최대 행수, 디폴트 생성자 등 다양한 생성자를 제공합니다.

#### 비밀번호 암호화로 Excel 내보내기
```java
try (var os = java.nio.file.Files.newOutputStream(java.nio.file.Path.of("secret.xlsx"))) {
    handler.consumeOutputStreamWithPassword(os, "P@ssw0rd!");
}
```

### 2) Excel 읽기 (Bean Validation 선택 지원)
```java
import io.github.dornol.excelkit.excel.*;
import jakarta.validation.Validation;
import jakarta.validation.Validator;

class User {
    // 예: Bean Validation 어노테이션을 여기에 부착할 수 있습니다.
    String name;
    Integer age;
}

Validator validator = Validation.buildDefaultValidatorFactory().getValidator();

// 인스턴스 공급자와(필수) Validator(선택)를 전달
ExcelReader<User> reader = new ExcelReader<>(User::new, validator);

ExcelReadHandler<User> rh = reader
        .column((u, cell) -> u.name = cell.asString())
        .column((u, cell) -> u.age = cell.asInt())
        .build(java.nio.file.Files.newInputStream(java.nio.file.Path.of("users.xlsx")));

// 행 단위 결과 소비
rh.read(result -> {
    if (result.success()) {
        User u = result.data();
        // 성공 처리
    } else {
        // 유효성 오류 메시지 등 확인
        System.out.println(result.messages());
    }
});
```

- 헤더는 기본적으로 첫 번째 행(인덱스 0)으로 인식합니다.
- `CellData`는 `asString()/asInt()/asLong()/asBigDecimal()/asLocalDate()` 등 변환 메서드를 제공합니다.
- Validator를 전달하지 않으면 유효성 검증 없이 파싱만 수행합니다.

#### 헤더 행 인덱스 지정
메타데이터나 타이틀이 있는 Excel 파일에서 헤더 행의 위치를 지정할 수 있습니다. 지정된 행 이전의 행은 무시됩니다.
```java
ExcelReadHandler<User> rh = new ExcelReader<>(User::new, null)
        .headerRowIndex(2) // 세 번째 행을 헤더로 사용 (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .column((u, cell) -> u.age = cell.asInt())
        .build(inputStream);
```

#### 특정 시트 읽기
```java
ExcelReadHandler<User> rh = new ExcelReader<>(User::new, null)
        .sheetIndex(1) // 두 번째 시트 (0-based)
        .column((u, cell) -> u.name = cell.asString())
        .column((u, cell) -> u.age = cell.asInt())
        .build(inputStream);
```

#### 대용량 파일 읽기 설정
```java
// 애플리케이션 시작 시 한 번만 호출
ExcelReader.configureLargeFileSupport();
```

### 3) CSV 쓰기
```java
import io.github.dornol.excelkit.csv.*;

record Row(long id, String name) {}

var rows = java.util.stream.Stream.of(new Row(1, "Alice"), new Row(2, "Bob"));

CsvWriter<Row> csv = new CsvWriter<>();
CsvHandler ch = csv
        .column("ID", r -> r.id())
        .column("Name", (r, c) -> r.name()) // c는 Cursor (현재 총/행 카운터 등)
        .constColumn("Const", "fixed")
        .write(rows);

try (var os = java.nio.file.Files.newOutputStream(java.nio.file.Path.of("rows.csv"))) {
    ch.consumeOutputStream(os); // 한 번만 호출 가능
}
```

CSV 파일은 UTF-8 BOM을 포함하여 작성되므로 Excel에서 직접 열어도 한글이 깨지지 않습니다.

### 4) CSV 읽기 (Bean Validation 선택 지원)
```java
import io.github.dornol.excelkit.csv.*;
import io.github.dornol.excelkit.shared.CellData;

class Product {
    String name;
    Integer price;
}

CsvReader<Product> csvReader = new CsvReader<>(Product::new, null);

CsvReadHandler<Product> crh = csvReader
        .column((p, cell) -> p.name = cell.asString())
        .column((p, cell) -> p.price = cell.asInt())
        .build(java.nio.file.Files.newInputStream(java.nio.file.Path.of("products.csv")));

crh.read(result -> {
    if (result.success()) {
        Product p = result.data();
        // 성공 처리
    } else {
        System.out.println(result.messages());
    }
});
```

## 고급 기능 및 팁
- 커서(Cursor)
  - `Cursor`를 통해 현재 행/총 행 등 내부 상태를 활용할 수 있습니다.
  - `getCurrentTotal()`은 `long` 타입을 반환하여 대용량 데이터셋에서도 안전합니다.
  - Excel/CSV 쓰기 모두 동일한 `Cursor`를 사용합니다.
- 자동 열 너비
  - Excel은 값의 논리 길이를 기준으로 열 너비를 계산하여 적용합니다(ASCII 1폭, 비ASCII 2폭 가중치).
  - 처음 100행의 데이터를 샘플링하여 최적 너비를 결정합니다.
- 시트 분할
  - `ExcelWriter`는 설정한 최대 행수에 도달하면 자동으로 새 시트를 생성합니다.
- 임시 리소스 관리
  - CSV/Excel 처리 중 생성되는 임시 파일·디렉토리는 `TempResourceContainer`를 통해 안전하게 정리됩니다.
- 단일 소비(consume-once)
  - `CsvHandler#consumeOutputStream`, `ExcelHandler#consumeOutputStream(WithPassword)`는 한 번만 호출할 수 있습니다. 재사용 시 예외가 발생합니다.
- Locale 설정
  - `CellData.setDefaultLocale(Locale)`로 숫자 파싱 시 기본 Locale을 변경할 수 있습니다. 기본값은 `Locale.KOREA`입니다.
- 날짜 포맷 관리
  - `CellData.addDateFormat(pattern)` / `CellData.addDateTimeFormat(pattern)`으로 커스텀 날짜 포맷을 추가할 수 있습니다.
  - `CellData.resetDateFormats()` / `CellData.resetDateTimeFormats()`로 기본 포맷으로 초기화할 수 있습니다.

## 예외 처리
- `ExcelKitException` — 모든 라이브러리 예외의 기본 클래스
  - `ExcelWriteException` — Excel 쓰기 관련 오류
  - `ExcelReadException` — Excel 읽기 관련 오류
  - `CsvWriteException` — CSV 쓰기 관련 오류
  - `CsvReadException` — CSV 읽기 관련 오류
- 컬럼 매핑 함수 내부 예외는 안전하게 로깅되며 기본 문자열 처리로 대체될 수 있습니다(Excel 쓰기).
- 잘못된 비밀번호(빈 값)로 암호화 호출 시 `IllegalArgumentException`이 발생합니다.
- 이미 소비된 핸들러를 다시 소비 시도하면 해당 `WriteException`이 발생합니다.

## 0.2.0에서 변경 사항 (0.2.1)

### Breaking Changes
- `Cursor.getCurrentTotal()` 반환 타입이 `int`에서 `long`으로 변경되었습니다. `ExcelDataType.INTEGER`와 함께 사용하던 경우 `ExcelDataType.LONG`으로 변경해야 합니다.
- `ExcelDataType.TIME`이 `LocalDateTime` 대신 `LocalTime`을 직접 받도록 변경되었습니다. 기존에 `LocalDateTime`으로 감싸던 코드를 제거하세요.

### New Features
- `ExcelWriter.rowHeight(float)` — 데이터 행 높이 설정 (기본 20pt)
- `ExcelReader.headerRowIndex(int)` — 헤더 행 인덱스 지정 (메타데이터가 있는 파일 지원)
- `CellData.resetDateFormats()` / `CellData.resetDateTimeFormats()` — 날짜 포맷 기본값 초기화

### Improvements
- CSV 파일에 UTF-8 BOM을 포함하여 Excel에서 한글 깨짐 방지
- CSV 읽기 시 BOM 자동 제거
- `CsvReadHandler` 예외 이중 래핑 수정

## 요구 사항
- JDK 버전: 프로젝트 설정에 따르며, 일반적으로 LTS(예: JDK 17+) 환경을 권장합니다.
- 메모리: 대용량을 다루지만, 매우 큰 데이터의 경우에도 적절한 배치/스트리밍 사용을 권장합니다.

## 개발 & 빌드
- Gradle 빌드: `./gradlew build`
- 테스트: `./gradlew test`

## 라이선스
MIT License. 자세한 내용은 [LICENSE](./LICENSE) 파일을 참고하세요.
