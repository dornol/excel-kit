# excel-kit

Apache POI 기반의 대용량 Excel(.xlsx) 및 CSV 생성/파싱을 간단하고 안전하게 처리하기 위한 경량 유틸리티입니다. 스트리밍 방식으로 동작하여 메모리 사용을 최소화하고, 컬럼 정의를 플루언트 DSL 스타일로 작성할 수 있습니다. 또한 Excel 비밀번호 암호화 출력과 Bean Validation(선택) 기반의 읽기 검증을 지원합니다.

- 그룹/아티팩트: `io.github.dornol:excel-kit`
- 라이선스: MIT

## 주요 기능
- Excel 쓰기 (SXSSFWorkbook 스트리밍 사용)
  - 컬럼별 타입/포맷/정렬 지정
  - 시트 행 수 자동 분할(최대 행수 지정)
  - 헤더 스타일 색상 지정
  - 결과물을 OutputStream으로 한 번에 소비(consume-once)
  - Excel 파일 비밀번호 암호화 출력 지원
- Excel 읽기 (SAX 기반 스트리밍 파싱)
  - 헤더 자동 인식 및 컬럼 매핑 DSL
  - Bean Validation(선택) 통합으로 행 단위 유효성 검증 결과 제공
  - 행별 파싱 결과/메시지 전달
- CSV 쓰기
  - 임시 파일로 스트리밍 작성 후 OutputStream으로 전달
  - CSV 기본 이스케이프 처리(인용부호/콤마/개행)
  - 커서(cursor) 제공: 현재 행/총 행 등 정보 활용 가능

> 주의: 현재 CSV 읽기(파싱) 기능은 제공하지 않습니다. CSV는 쓰기만 지원합니다.

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

ExcelWriter<Person> writer = new ExcelWriter<>(255, 255, 255, 1_000_000); // 헤더색(흰색), 시트당 최대행
ExcelHandler handler = writer
        .column("ID", p -> p.id())
            .type(ExcelDataType.NUMERIC)
            .alignment(HorizontalAlignment.RIGHT)
        .column("Name", p -> p.name())
            .type(ExcelDataType.STRING)
        .column("Age", (p, c) -> p.age()) // 커서(c)도 전달 가능
            .type(ExcelDataType.NUMERIC)
            .format("0")
        .write(stream); // 마지막에 write 호출

// 출력(한 번만 가능)
try (var os = java.nio.file.Files.newOutputStream(java.nio.file.Path.of("people.xlsx"))) {
    handler.consumeOutputStream(os);
}
```

- `ExcelWriter#column(name, function)`으로 컬럼을 정의하고, 체이닝으로 `type/format/alignment/style` 등을 지정합니다.
- `write(Stream<T>)` 또는 `write(Stream<T>, ExcelConsumer<T>)`를 호출하면 `ExcelHandler`가 반환됩니다.
- `ExcelHandler#consumeOutputStream(OutputStream)`은 한 번만 호출할 수 있습니다. 재호출 시 예외가 발생합니다.

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
        .column((u, cell) -> u.age = cell.asInteger())
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

- 헤더는 첫 번째 행으로 가정하고 자동 인식합니다.
- `ExcelCellData`는 `asString()/asInteger()/asLong()/asBigDecimal()/asLocalDate()` 등 변환 메서드를 제공합니다.
- Validator를 전달하지 않으면 유효성 검증 없이 파싱만 수행합니다.

### 3) CSV 쓰기
```java
import io.github.dornol.excelkit.csv.*;

record Row(long id, String name) {}

var rows = java.util.stream.Stream.of(new Row(1, "Alice"), new Row(2, "Bob"));

CsvWriter<Row> csv = new CsvWriter<>();
CsvHandler ch = csv
        .column("ID", r -> r.id())
        .column("Name", (r, c) -> r.name()) // c는 CsvCursor (현재 총/행 카운터 등)
        .constColumn("Const", "fixed")
        .write(rows);

try (var os = java.nio.file.Files.newOutputStream(java.nio.file.Path.of("rows.csv"))) {
    ch.consumeOutputStream(os); // 한 번만 호출 가능
}
```

## 고급 기능 및 팁
- 커서(Cursor)
  - Excel: `ExcelCursor`를 통해 현재 행/열 등 내부 상태를 활용할 수 있습니다.
  - CSV: `CsvCursor`는 `plusRow/plusTotal` 등으로 내부 카운팅을 제공하며, 컬럼 함수에 전달됩니다.
- 자동 열 너비
  - Excel은 값의 논리 길이를 기준으로 열 너비를 계산하여 적용합니다(ASCII 1폭, 비ASCII 2폭 가중치).
- 시트 분할
  - `ExcelWriter`는 설정한 최대 행수에 도달하면 자동으로 새 시트를 생성합니다.
- 임시 리소스 관리
  - CSV/Excel 처리 중 생성되는 임시 파일·디렉토리는 `TempResourceContainer`를 통해 안전하게 정리됩니다.
- 단일 소비(consume-once)
  - `CsvHandler#consumeOutputStream`, `ExcelHandler#consumeOutputStream(WithPassword)`는 한 번만 호출할 수 있습니다. 재사용 시 `IllegalStateException`이 발생합니다.

## 요구 사항
- JDK 버전: 프로젝트 설정에 따르며, 일반적으로 LTS(예: JDK 17+) 환경을 권장합니다.
- 메모리: 대용량을 다루지만, 매우 큰 데이터의 경우에도 적절한 배치/스트리밍 사용을 권장합니다.

## 예외 및 에러 처리
- 컬럼 매핑 함수 내부 예외는 안전하게 로깅되며 기본 문자열 처리로 대체될 수 있습니다(Excel 쓰기).
- 잘못된 비밀번호(빈 값)로 암호화 호출 시 `IllegalArgumentException`이 발생합니다.
- 이미 소비된 핸들러를 다시 소비 시도하면 `IllegalStateException`이 발생합니다.

## 개발 & 빌드
- Gradle 빌드: `./gradlew build`
- 테스트: `./gradlew test`

## 라이선스
MIT License. 자세한 내용은 [LICENSE](./LICENSE) 파일을 참고하세요.
