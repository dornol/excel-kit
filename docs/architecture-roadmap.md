# Architecture and Usability Roadmap

이 문서는 excel-kit의 다음 주요 버전에서 검토할 구조 개선과 기능 후보를 정리한다.
현재 기능 목록을 늘리는 것보다 설치 경험, API 일관성, 리소스 수명, 확장성을 개선하는 데 초점을 둔다.

아직 확정된 릴리스 계획이나 호환성 약속은 아니다. 각 제안은 독립적으로 채택할 수 있지만,
아래의 우선순위와 의존 관계를 따르면 중복 작업을 줄일 수 있다.

## 현재 상태

excel-kit은 다음 영역을 이미 폭넓게 지원한다.

- Excel/CSV 스트리밍 읽기와 쓰기
- setter, mapping, map 기반 읽기
- 스타일, 수식, 이미지, 차트, validation, protection
- 다중 sheet와 대용량 sheet rollover
- header alias, strict header, duplicate header 처리
- 행 제한, 빈 행 처리, 진행률 callback
- schema 기반 Excel/CSV 공통 매핑
- Spring MVC 다운로드 및 업로드 지원

따라서 다음 버전의 가장 큰 개선 여지는 새로운 Excel 서식 기능보다 다음 경계에 있다.

- 어떤 의존성을 사용자가 직접 설치해야 하는가
- reader와 handler 중 누가 실행과 리소스를 소유하는가
- Excel과 CSV 설정이 얼마나 일관되게 확장되는가
- mutable 객체와 immutable 객체가 schema를 같은 수준으로 활용할 수 있는가
- 오류와 실행 결과를 애플리케이션이 얼마나 쉽게 관찰할 수 있는가

## 목표와 비목표

### 목표

- Excel만 또는 CSV만 사용하는 애플리케이션이 필요한 의존성만 설치할 수 있게 한다.
- fluent configuration과 실제 읽기 실행의 경계를 명확히 한다.
- `InputStream`, 임시 파일, background thread의 소유권과 수명을 API에서 드러낸다.
- 공통 옵션을 immutable value object로 묶어 Excel/CSV/schema 간 설정 누락을 방지한다.
- record와 immutable DTO도 schema를 완전하게 재사용할 수 있게 한다.
- 오류가 과도하게 많은 입력을 제한해 서버 자원과 처리 시간을 보호한다.

### 비목표

- Apache POI 또는 OpenCSV 자체를 대체하지 않는다.
- annotation 기반 domain mapping을 기본 설계로 도입하지 않는다.
- 모든 변경을 한 릴리스에 동시에 포함할 필요는 없다.
- 성능 근거 없이 비동기 또는 병렬 처리를 기본값으로 만들지 않는다.

## 제안 1: 선택적 런타임 의존성 정책 공식화

### 문제

현재 `kit` 모듈은 POI, OpenCSV, SLF4J, Jakarta Validation, JSpecify를 모두
`compileOnly`로 선언한다. 이는 누락이 아니라 애플리케이션이 이미 사용하는 라이브러리
버전을 excel-kit이 강제로 바꾸지 않기 위한 의도적인 정책이다.

이 방식은 의존성 버전 통제권을 애플리케이션에 준다는 장점이 있지만, 현재는 정책과
지원 범위가 충분히 명문화되어 있지 않다.

- 특정 버전이 필수인지 단순한 테스트 기준 버전인지 README만으로 구별하기 어렵다.
- Excel, CSV, Validation별로 실제 필요한 의존성을 사용자가 구분해야 한다.
- 필수 runtime dependency를 빠뜨리면 컴파일 이후에 `ClassNotFoundException`이 발생할 수 있다.
- 현재 CI는 version catalog에 지정된 단일 조합만 검증한다.
- 사용자가 선택한 이전/새 버전을 어디까지 지원하는지 객관적인 근거가 없다.

### 유지할 구조와 정책

```text
excel-kit
└── Excel/CSV/core API, integration libraries는 compileOnly

excel-kit-spring
└── Spring MVC/WebFlux adapter
```

다음 원칙을 유지한다.

1. `excel-kit`을 core/excel/csv로 분리하지 않는다.
2. POI, OpenCSV, SLF4J, Jakarta Validation의 `compileOnly` 정책을 유지한다.
3. SLF4J 구현체나 Jakarta Validation 구현체를 전이 의존성으로 제공하지 않는다.
4. 사용자는 자신이 사용하는 기능에 필요한 integration과 버전만 선택한다.
5. excel-kit은 문서화된 호환 범위를 CI에서 검증한다.

Gradle 설치 예제는 기능별로 분리한다.

```kotlin
// Excel
implementation("io.github.dornol:excel-kit:<version>")
implementation("org.apache.poi:poi-ooxml:<application-version>")
implementation("org.slf4j:slf4j-api:<application-version>")
```

```kotlin
// CSV
implementation("io.github.dornol:excel-kit:<version>")
implementation("com.opencsv:opencsv:<application-version>")
implementation("org.slf4j:slf4j-api:<application-version>")
```

```kotlin
// Optional Bean Validation
implementation("jakarta.validation:jakarta.validation-api:<application-version>")
```

### 호환성 표와 CI matrix

README에는 `Minimum supported`와 `Currently tested`를 구분한 표를 둔다.

| Dependency | Minimum supported | Currently tested | Required for |
|------------|-------------------|------------------|--------------|
| Apache POI | CI 검증 후 결정 | version catalog 값 | Excel |
| OpenCSV | CI 검증 후 결정 | version catalog 값 | CSV read |
| SLF4J API | CI 검증 후 결정 | version catalog 값 | Logging |
| Jakarta Validation API | CI 검증 후 결정 | version catalog 값 | Optional validation |

최소 지원 버전은 추측으로 정하지 않는다. 후보 버전으로 전체 테스트를 실행하고 통과한 조합만
지원 범위로 문서화한다. 최소한 다음 축을 CI에서 검증한다.

```text
JDK:     17 / 21 / current LTS
POI:     minimum-supported / currently-tested
OpenCSV: minimum-supported / currently-tested
```

Gradle property로 테스트 의존성 버전을 덮어쓸 수 있게 하면 동일 테스트 suite를 재사용할 수 있다.

```bash
./gradlew test \
  -PpoiVersion=<version> \
  -PopencsvVersion=<version>
```

matrix의 모든 가능한 조합을 매번 실행할 필요는 없다. 최소 조합과 현재 조합을 필수로 두고,
추가 조합은 주기적 workflow 또는 release 검증에서 실행할 수 있다.

### 누락된 의존성의 오류 경험

가능한 진입점에서는 일반적인 `NoClassDefFoundError`보다 필요한 artifact를 알려주는 메시지를
제공한다.

```text
Apache POI is required to use ExcelReader.
Add org.apache.poi:poi-ooxml to the application's runtime dependencies.
```

단, public class의 signature나 static initialization 과정에서 외부 타입을 먼저 해석하면 library가
오류를 변환하기 전에 class loading이 실패할 수 있다. 따라서 실제 적용 범위를 bytecode와 class
loading 테스트로 확인해야 하며, 불가능한 경우 설치 문서와 troubleshooting 항목을 우선한다.

### 호환성 영향

- module, package, public API 변경이 없으므로 원칙적으로 breaking change가 아니다.
- `compileOnly`를 유지하므로 기존 애플리케이션의 dependency resolution도 바뀌지 않는다.
- 실제 검증 결과에 따라 지원 최소 버전을 올리면 이는 지원 정책상의 변경이므로 changelog에 남긴다.
- runtime dependency 누락 오류를 감싸더라도 기존 `NoClassDefFoundError`에 의존한 코드는 거의 없겠지만
  예외 종류가 달라질 수 있으므로 release note에 기록한다.

### 완료 기준

- README가 Excel, CSV, Validation 설치를 각각 설명한다.
- `compileOnly`가 의도된 버전 비강제 정책임을 명시한다.
- 최소 지원 버전과 현재 테스트 버전을 구분해 게시한다.
- 최소/현재 dependency 조합을 CI가 실제로 검증한다.
- Excel-only와 CSV-only sample이 각각 필요한 의존성만으로 실행된다.
- 지원 범위 밖 버전에 대한 보장 여부와 issue reporting 기준을 문서화한다.

## 제안 2: Reader 실행 API와 리소스 소유권 개편

### 문제

현재 일반적인 읽기 흐름은 다음과 같다.

```java
ExcelReader.setter(User::new)
    .column("Name", (user, cell) -> user.setName(cell.asString()))
    .build(inputStream)
    .read(result -> handle(result));
```

`build(inputStream)`은 이름상 설정 완료처럼 보이지만 실제로는 다음 작업도 수행한다.

- 임시 디렉터리와 임시 파일 생성
- 입력 스트림 전체를 임시 파일에 복사
- 전달받은 입력 스트림 닫기
- 한 번만 소비할 수 있는 handler 생성

즉 `Reader`는 builder, `ReadHandler`는 실행 session 역할을 하지만 그 경계가 API에서
충분히 명확하지 않다. 특히 caller가 생성한 `InputStream`을 library가 닫는다는 사실은
메서드 이름만으로 예측하기 어렵다.

### 권장 API

Reader는 재사용 가능한 configuration으로 만들고 입력은 실행 메서드에서 받는다.

```java
ExcelReader<User> reader = ExcelReader.setter(User::new)
    .column("Name", (user, cell) -> user.setName(cell.asString()));

reader.read(inputStream, result -> handle(result));
```

`Path`, `InputStream`, stream supplier를 일급 입력으로 지원한다.

```java
reader.read(Path.of("users.xlsx"), consumer);
reader.read(inputStream, consumer);
reader.read(file::getInputStream, consumer);
```

리소스를 연 쪽이 닫는다는 규칙을 고정한다.

- caller가 전달한 `InputStream`은 caller가 닫는다.
- Reader가 `Path` 또는 supplier를 통해 연 stream은 Reader가 닫는다.
- Reader가 만든 parser, package, 임시 파일, background thread는 Reader가 정리한다.

따라서 `InputOwnership` 같은 선택 enum은 제공하지 않는다. 동일한 parameter type에서 호출마다
소유권을 바꾸는 것보다 입력 형태에 따라 하나의 예측 가능한 규칙을 적용한다.

Spring Web에서는 `MultipartFile`을 core Reader에 직접 전달하지 않는다. Controller나 Spring
adapter가 stream을 열어 try-with-resources로 닫거나, opener를 supplier 형태로 Reader에 전달한다.

```java
try (InputStream input = file.getInputStream()) {
    reader.read(input, consumer);
}

// Reader가 열고 닫게 하는 경우
reader.read(file::getInputStream, consumer);
```

내부적으로 실행별 `ReadSession`을 둘 수 있지만 public API에서 반드시 노출할 필요는 없다.

### 대안

- `open(input)`이 `AutoCloseable ReadSession`을 반환하고 session에서 `read()`를 호출한다.

두 번째 대안은 수명은 가장 명확하지만 사용 단계가 현재와 비슷하게 많다. 일반 사용성은
직접 실행 API가 더 낫고, 고급 제어가 필요할 때만 session API를 추가하는 것이 적절하다.

### 호환성 영향

- `build(InputStream)`과 public handler 타입 제거 시 큰 source breaking change가 된다.
- 기존 handler가 제공하던 overload를 Reader로 이동해야 한다.
- 입력 스트림을 자동으로 닫던 동작을 바꾸면 리소스 정책도 behavioral breaking change가 된다.

### migration 예시

```java
// Before
reader.build(input).read(consumer);

// After
try (InputStream input = source.openStream()) {
    reader.read(input, consumer);
}
```

### 완료 기준

- public method마다 입력과 출력 리소스의 소유권이 Javadoc에 명시된다.
- 전달받은 `InputStream`은 성공과 실패 모두에서 Reader가 닫지 않는다.
- `Path`와 supplier로 Reader가 연 stream은 성공과 실패 모두에서 Reader가 닫는다.
- 동일 Reader 설정으로 여러 입력을 순차 실행할 수 있다.
- 성공, 실패, callback 예외, `readWhile()` 조기 종료 모든 경로에서 임시 파일이 제거된다.
- `Path` 입력은 불필요하게 별도 임시 파일로 한 번 더 복사하지 않는다.

## 제안 3: Reader/Writer 설정 snapshot 패턴 통일

### 문제

Reader와 Writer는 외부에서 모두 fluent API를 사용하지만 내부 설정 저장 방식은 다르다.

- Excel/CSV Reader는 설정 대부분을 자신의 mutable field에 저장하고 handler의 긴 생성자
  인자로 하나씩 전달한다.
- Excel Writer는 writer field, `SheetConfig`, `HeaderStyleConfig`, `ColumnStyleConfig`,
  `InitOptions` 등에 설정이 목적별로 나뉜다.
- CSV Writer는 설정 대부분을 자신의 mutable field에 저장한다.
- `ExcelKitSchema`는 기본값을 별도 field에 저장한 뒤 Reader/Writer에 fluent setter로 복사한다.

Writer의 설정 계층이 Reader보다 복잡한 것은 workbook, sheet, header, column, row라는 출력 구조상
자연스럽다. 문제는 외부 문법의 차이가 아니라 실행 계층으로 설정을 전달할 때 고정된 snapshot이
없고, 특히 Reader handler 생성자에 개별 값이 길게 나열된다는 점이다.

새로운 공통 옵션 하나를 추가하려면 일반적으로 다음 위치를 함께 수정해야 한다.

- `AbstractReader`
- `ExcelReader.build()`
- `CsvReader.build()`
- Excel/CSV handler 생성자와 field
- `ExcelKitSchema`와 `applyReaderDefaults()`
- Spring adapter에서 별도 노출하는 경우 해당 DTO/helper

이는 컴파일로 일부 누락을 잡을 수 있어도 Excel/CSV 중 한쪽에 값을 전달하지 않는 실수를
만들기 쉽다.

### 외부 사용 흐름

Reader와 Writer 모두 다음 흐름으로 맞춘다.

```text
factory → fluent configuration/columns → execution
```

```java
// Read
ExcelReader.setter(User::new)
    .column("Name", (user, cell) -> user.setName(cell.asString()))
    .strictHeaders()
    .maxRows(10_000)
    .read(input, consumer);

// Write
ExcelWriter.<User>create()
    .column("Name", User::name)
    .maxRows(10_000)
    .write(users)
    .writeTo(output);
```

기존 fluent method는 유지한다. Reader에 전체 설정용 `create(options -> ...)`를 추가하거나,
Reader와 Writer 모두에 동일한 `configure(options -> ...)`를 추가하지 않는다. 같은 옵션을 fluent
method와 config lambda 두 방식으로 설정하게 만들면 API 표면적과 우선순위 규칙만 늘어난다.

Writer의 `create(InitOptions)`는 SXSSF workbook 생성 전에 반드시 필요한 설정에만 사용한다.
이는 Reader/Writer의 일반 configuration 문법이 아니라 생성 시점의 제약을 표현하는 API다.

컬럼과 스타일 설정 lambda도 특정 하위 객체를 구성하기 위한 DSL로 유지한다.

```java
writer.column("Price", Product::price, column -> column
    .type(ExcelDataType.INTEGER)
    .format("#,##0"));
```

### 내부 snapshot 구조

fluent builder는 mutable하게 유지하되 `read()` 또는 `write()`가 시작되는 순간 immutable options
snapshot을 만든다. 실행 session은 원본 Reader/Writer field가 아니라 snapshot만 참조한다.

```text
Mutable fluent Reader/Writer
            │
            │ snapshot at execution start
            ▼
Immutable ReadOptions/WriteOptions
            │
            ▼
Per-input ReadSession/WriteSession
```

Reader용 타입은 다음처럼 구성할 수 있다.

```java
public record HeaderOptions(
    int headerRowIndex,
    boolean strict,
    DuplicateHeaderPolicy duplicatePolicy
) {}

public record RowLimitOptions(
    long maxRows,
    boolean skipBlankRows,
    int stopAtBlankRows
) {}

public record ProgressOptions(
    int interval,
    ProgressCallback callback
) {}

public record ReadOptions(
    HeaderOptions header,
    RowLimitOptions limits,
    CellConversionConfig conversion,
    ProgressOptions progress,
    long maxErrors
) {}

public record ExcelReadOptions(
    ReadOptions common,
    int sheetIndex,
    int headerRows,
    boolean countRows
) {}

public record CsvReadOptions(
    ReadOptions common,
    CsvDialect dialect,
    CsvParserOptions parser
) {}
```

Writer는 Reader와 동일한 단일 options 타입을 공유하지 않는다. 출력 구조에 맞춰 workbook/sheet/header
설정 객체를 조합한 `ExcelWriteOptions`와 단순한 `CsvWriteOptions`를 별도로 둔다.

```java
ExcelReadOptions readOptions = reader.snapshotOptions();
ExcelWriteOptions writeOptions = writer.snapshotOptions();
```

`snapshotOptions()`는 우선 package-private 내부 API로 둔다. logging이나 테스트 fixture에 실제 수요가
생기기 전에는 public configuration API로 노출하지 않는다.

### 설계 원칙

- option record 생성 시 모든 범위 검증을 끝낸다.
- disabled 상태는 가능하면 magic number `-1` 대신 명시적인 타입으로 표현한다.
- password 같은 secret은 `toString()`에 포함되지 않게 별도 credentials 타입으로 둔다.
- callback을 포함한 options와 순수 value options를 나눌지 검토한다.
- Excel/CSV 공통 설정은 반드시 `ReadOptions`에만 둔다.
- 실행 시작 후 원본 Reader/Writer의 fluent 설정을 변경해도 진행 중인 session에 영향을 주지 않는다.
- Reader와 Writer의 options 타입을 억지로 상속하거나 하나의 거대한 공통 타입으로 합치지 않는다.
- fluent builder 자체를 immutable 객체로 바꾸지 않는다. 실행 snapshot만 immutable하게 만든다.

### 완료 기준

- Reader handler/session 생성자에 개별 설정 primitive가 나열되지 않는다.
- Writer 실행 지원 코드도 실행 중 mutable writer field 대신 snapshot을 참조한다.
- schema가 Reader setter를 연속 호출하지 않고 options를 직접 전달한다.
- Excel과 CSV 공통 옵션 parity를 단일 parameterized test로 검증한다.
- Reader와 Writer의 public 흐름이 모두 `factory → configuration/columns → execution`을 따른다.

## 제안 4: Schema의 immutable object mapping 개선 — 보류

### 검토한 문제

현재 `ExcelKitSchema`는 하나의 column에 다음 정보를 저장할 수 있다.

- 출력 header 이름
- 객체에서 값을 꺼내는 write function
- mutable 객체에 값을 넣는 read setter
- Excel write configurer
- 읽을 때 허용할 header alias
- required 여부

이 설계는 JavaBean에는 잘 맞지만 record와 immutable DTO에는 setter가 없으므로 완전하게
적용되지 않는다. Mapping reader를 사용할 수는 있지만, mapping mode에서는 schema column의
read definition을 사용하지 않고 `RowData`에서 header 이름으로 값을 다시 조회해야 한다.

결과적으로 immutable 사용자는 이미 schema에 정의한 header 이름을 mapping lambda에서 다시 적을
수 있다.

### 보류 결정

라이브러리 수준의 `ColumnKey<V>`를 추가하는 안은 채택하지 않는다. 작은 schema에도 key 선행 선언이
필요해지고 generic, nullability, equality, read/write 방향 같은 새로운 개념을 함께 설계해야 한다.
현재 문제 크기에 비해 API 복잡도가 커질 가능성이 높다.

header 이름 재사용이 필요한 사용자는 상수나 enum으로 직접 관리할 수 있다.

```java
private static final String NAME = "Name";
private static final String AGE = "Age";

ExcelReader.mapping(row -> new Person(
    row.get(NAME).asString(),
    row.get(AGE).asInt()
));
```

record와 immutable DTO는 현재의 `ExcelReader.mapping(...)`, `CsvReader.mapping(...)`과 `RowData`를
계속 사용한다. Schema의 mapper 보관, write-only column, constructor mapping 같은 작은 대안도 실제
사용 사례가 확인되기 전에는 추가하지 않는다.

### 재검토 조건

다음 중 하나가 반복적으로 확인될 때 별도 제안으로 다시 검토한다.

- 동일 immutable mapper를 Excel과 CSV에서 중복 정의하는 사례가 많다.
- Schema의 setter 필수 구조 때문에 write-only schema를 만들 수 없는 문제가 실제로 발생한다.
- header 상수만으로 해결되지 않는 converter/alias/required 설정 중복이 다수 발생한다.
- 사용자 요청이나 example 코드에서 record 기반 schema가 주요 사용 방식이 된다.

보류 상태에서는 public API, `SchemaColumn`, migration 계획을 변경하지 않는다.

## 제안 5: 최대 읽기 오류 수 제한

### 문제

현재 오류 처리 방식은 호출한 메서드에 의해 명확하게 구분되어 있다.

- `read(Consumer<ReadResult<T>>)`는 성공과 실패를 값으로 전달한다.
- `read(onSuccess, onError)`는 두 callback으로 분리한다.
- `readStrict()`는 첫 실패에 `ReadAbortException`을 던진다.
- callback에서 예외를 던져 읽기를 중단할 수도 있다.

기존 API만으로 오류 보고, 계속 진행, 즉시 중단을 모두 표현할 수 있으므로 별도의
`ReadErrorPolicy` enum은 추가하지 않는다. `ReadSummary`도 일반 core 사용자의 수요가 확인되지 않은
상태에서 Spring의 `UploadSummary`와 별개로 추가하지 않는다.

남아 있는 문제는 잘못된 대용량 파일에서 변환 또는 검증 오류가 지나치게 많이 발생해도 파일 끝까지
계속 읽는다는 점이다. callback이 오류를 저장하거나 error report를 만들면 메모리, 디스크, 처리 시간이
불필요하게 증가할 수 있다.

### 제안 API

```java
reader
    .maxErrors(100)
    .read(input, resultConsumer);
```

의미는 다음처럼 정의한다.

| 설정 | 동작 |
|------|------|
| 설정하지 않음 | 기존처럼 오류 개수 제한 없이 계속 읽음 |
| `maxErrors(0)` | 첫 오류가 발생하면 즉시 중단 |
| `maxErrors(100)` | 오류 100개까지 callback으로 전달하고 101번째 오류에서 중단 |

오류 수는 library가 count만 추적한다. 오류 객체 전체를 library 내부에 보관하지 않으며 제한 안의
오류는 기존 `ReadResult` 또는 error callback으로 그대로 전달한다.

### 중단 예외

제한을 초과하면 기존 `ReadAbortException`을 사용한다. 호출자가 일반 strict 오류와 최대 오류 수 초과를
구별할 필요가 있으므로 reason을 추가한다.

```java
public enum ReadAbortReason {
    STRICT_FAILURE,
    MAX_ERRORS_EXCEEDED
}
```

예외에는 최소한 다음 정보를 보존한다.

- 중단 reason
- 설정된 최대 오류 수와 현재 오류 수
- 마지막 logical row와 physical file row
- 마지막 `RowError` 또는 `ReadResult`의 구조화된 오류 정보

기존 `readStrict()`는 편의 API로 유지한다. `maxErrors(0)`과 중단 시점은 유사하지만 성공 callback만
받는 strict API의 의미가 명확하므로 제거하거나 `maxErrors(0)`로 대체하지 않는다.

### 완료 기준

- Excel과 CSV Reader가 동일한 `maxErrors(long)` API와 경계값 의미를 제공한다.
- 음수 값은 설정 시점에 거부한다.
- 제한 안의 오류는 기존 callback으로 모두 전달된다.
- 제한을 초과한 오류에서 정확히 한 번 중단되고 내부 리소스가 정리된다.
- `ReadAbortException`으로 strict failure와 최대 오류 수 초과를 구별할 수 있다.
- `read()`, 분리 callback `read()`, `readWhile()`에서 동일한 오류 count 규칙을 사용한다.

## 제안 6: Java Stream 읽기 API 제거

### 문제

Excel의 현재 `readAsStream()`은 SAX parser의 push callback을 Java Stream의 pull 소비 방식에
연결하기 위해 다음 구조를 사용한다.

```text
POI parser producer thread
        │
        ▼
bounded BlockingQueue (1024)
        │
        ▼
Java Stream consumer thread
```

callback 기반 `read()`도 파일 전체를 메모리에 올리지 않고 한 행씩 처리하는 실제 streaming API다.
따라서 대용량 처리를 위해 Java `Stream` 반환형이 반드시 필요한 것은 아니다. 반면 현재 bridge는
다음 복잡성을 만든다.

- stream 호출마다 platform thread가 생성된다.
- buffer 크기가 고정되어 있다.
- consumer가 `limit()`, `findFirst()` 등으로 조기에 멈출 때 producer를 interrupt해야 한다.
- 사용자가 stream을 닫지 않으면 thread와 임시 파일 수명이 길어질 수 있다.
- producer exception을 consumer thread로 전달하기 위한 별도 channel이 필요하다.
- Spring의 transaction, MDC, SecurityContext 등 thread-local context가 producer와 consumer 경계에서
  이해하기 어려워질 수 있다.

### 결정

`ExcelReadHandler.readAsStream()`과 `CsvReadHandler.readAsStream()`을 제거한다. 2번의 새 Reader API에도
`reader.stream(input)`은 추가하지 않는다. callback 기반 `read()`를 유일한 일반 스트리밍 API로 둔다.

```java
reader.read(input, result -> process(result));
```

이를 통해 다음 내부 구성요소를 제거할 수 있다.

- per-read producer thread
- bounded blocking queue와 sentinel
- producer exception 전달용 atomic state
- interrupt와 join 기반 cancellation
- 반환 Stream의 `onClose` cleanup bridge

전체 결과 수집 API는 추가하지 않는다. 필요한 사용자는 callback에서 직접 collection에 추가할 수
있지만, 대용량 라이브러리가 `toList()`를 편의 기능으로 권장하지는 않는다.

### 정상적인 조기 종료

Stream의 `limit()`과 `findFirst()`를 대체할 수 있도록 `readWhile()`을 제공한다. predicate가 `true`를
반환하면 계속 읽고 `false`를 반환하면 정상 종료한다.

```java
reader.readWhile(input, result -> {
    process(result);
    return shouldContinue(result);
});
```

기존 `read()`의 `Consumer` overload와 lambda resolution이 혼동되지 않도록 별도 이름을 사용한다.
`boolean` 의미는 Javadoc에 `true = continue`, `false = stop`으로 명시한다. 별도 `ReadAction` enum은
두 값만 필요한 현재 범위에서는 추가하지 않는다.

`readWhile()`의 `false`는 오류나 abort가 아닌 정상 종료다. 내부 parser, package와 임시 파일은 즉시
정리하며, `ReadAbortException`을 던지지 않는다.

### Spring Web 관점

callback은 파일을 읽는 caller thread에서 동기적으로 실행한다.

```java
try (InputStream input = file.getInputStream()) {
    reader.read(input, result -> importRow(result));
}
```

별도 producer thread가 없어 transaction context, MDC, SecurityContext와 callback 예외가 자연스럽게
같은 thread에 유지된다. 요청 취소와 입력 stream lifecycle도 단순해진다.

### 호환성 및 migration

이 변경은 `readAsStream()` 사용자의 source breaking change다.

```java
// Before
try (Stream<ReadResult<User>> rows = handler.readAsStream()) {
    rows.filter(ReadResult::success).forEach(this::save);
}

// After
reader.read(input, result -> {
    if (result.success()) {
        save(result.data());
    }
});
```

`limit()`이나 `findFirst()`를 사용하던 코드는 `readWhile()`과 caller-owned holder로 이동한다.

### 검증 항목

- 전체 소비
- `readWhile()` 첫 행 종료와 중간 종료
- mapper와 callback 예외
- parser 예외
- `maxErrors()`에 의한 중단
- caller thread interrupt
- 반복 실행 후 임시 파일 누수
- callback이 caller thread에서 실행되는지 확인

### 완료 기준

- public API와 문서에서 `readAsStream()` 사용 예제가 제거된다.
- producer thread, blocking queue와 관련 lifecycle 코드가 제거된다.
- callback 기반 전체 읽기와 `readWhile()` 조기 종료가 동일한 session cleanup 경로를 사용한다.
- Excel과 CSV가 동일한 `readWhile()` 동작을 제공한다.
- 제거 전후 throughput과 peak memory를 비교하고 허용 regression 기준을 통과한다.

## 제안 7: Writer의 `Iterable` 입력과 Stream 소유권 정리

### 문제

현재 주요 쓰기 API는 `Stream<T>`를 입력으로 받고 내부에서 try-with-resources로 닫는다.
편리하지만 호출자가 만든 stream, 특히 DB cursor나 네트워크 리소스를 감싼 stream의 소유권을
library가 가져간다는 점이 메서드 signature에 나타나지 않는다.

또한 이미 `List`, `Set` 같은 `Iterable`을 가진 사용자가 간단한 쓰기에도 직접 stream을 만들어야 한다.

### 결정된 API

```java
ExcelHandler write(Iterable<T> rows);
ExcelHandler write(Stream<T> rows);
```

CSV Writer에도 동일한 overload를 제공한다.

가장 흔한 collection 사용은 다음처럼 단순해진다.

```java
writer.write(users);
```

`Iterable` overload는 collection 전체를 복사하거나 크기를 미리 계산하지 않고 iterator를 직접
순회한다. lazy `Iterable`도 허용하며 `size()`를 요구하지 않는다.

`Iterator<T>` 전용 overload는 추가하지 않는다. 실제 사용 사례가 확인되기 전에는 API 수를 늘리지
않고, 일반 collection과 사용자 정의 source는 `Iterable<T>`로 받는다.

Writer의 Stream 입력은 유지한다. Reader가 반환하던 Stream과 달리 Writer는 전달받은 Stream을 같은
thread에서 동기적으로 소비할 수 있어 producer thread나 queue가 필요하지 않으며, DB cursor가
Stream으로 제공되는 경우도 많다.

### 소유권 규칙

2번의 Reader와 같은 원칙으로 리소스를 연 쪽이 닫는다. `InputOwnership` enum은 추가하지 않는다.

- caller가 전달한 `Stream<T>`은 caller가 닫는다.
- Writer는 Stream을 소비하지만 `close()`하지 않는다.
- Writer가 생성한 workbook, 임시 파일과 output 관련 내부 리소스는 Writer/Handler가 정리한다.
- `Iterable<T>`은 close 대상이 없는 일반 반복 입력으로 취급한다.

리소스를 가진 DB Stream은 caller가 try-with-resources로 관리한다.

```java
try (Stream<User> users = repository.streamUsers()) {
    writer.write(users).writeTo(output);
}
```

이는 현재 Writer가 전달받은 Stream을 내부에서 닫는 동작을 바꾸는 behavioral breaking change다.

```java
// Before: Writer가 stream을 닫음
writer.write(repository.streamUsers()).writeTo(output);

// After: stream을 연 caller가 닫음
try (Stream<User> users = repository.streamUsers()) {
    writer.write(users).writeTo(output);
}
```

### 추가 검토

Java의 `Flow.Publisher<T>` 또는 Reactor `Flux<T>` 지원은 core에 직접 넣지 않고 integration 모듈로
두는 편이 적절하다. POI 쓰기는 기본적으로 blocking API이므로 reactive 타입만 받는다고 non-blocking이
되지는 않는다.

### 완료 기준

- Excel과 CSV Writer가 `Iterable<T>`과 `Stream<T>` 입력을 동일하게 제공한다.
- 전달받은 Stream은 성공, extractor 예외, output 예외 모두에서 Writer가 닫지 않는다.
- 빈 Iterable, 반복 중간 예외, extractor 예외에서도 workbook과 임시 파일이 정리된다.
- Iterable은 전체 복사 없이 한 행씩 소비된다.
- Stream 소유권 변경과 try-with-resources migration을 문서화한다.

## 제안 8: Header normalization

사용자가 만든 파일의 header에는 앞뒤 공백, 대소문자, Unicode 표현 차이가 흔히 발생한다. alias를
모두 등록하지 않고 Excel과 CSV에서 동일한 정규화 함수를 적용할 수 있게 한다.

별도 `HeaderNormalizer` 타입 계층은 만들지 않고 Java 표준 함수형 타입인
`UnaryOperator<String>` 하나를 받는다.

```java
reader.headerNormalizer(header ->
    Normalizer.normalize(header.trim(), Normalizer.Form.NFC)
        .toLowerCase(Locale.ROOT));
```

- 설정하지 않으면 현재와 동일한 identity normalization을 사용한다.
- null 함수와 null normalization 결과는 설정 또는 처리 시점에 명확하게 거부한다.
- 원본 header는 오류 리포트에 보존한다.
- 정규화된 header 충돌은 기존 `DuplicateHeaderPolicy`로 처리한다.
- 대소문자 변환에는 명시적인 locale을 사용한다.
- Excel과 CSV에 동일한 normalization 규칙을 적용한다.
- schema header, aliases, selected map columns에도 같은 함수를 적용해 비교 기준을 통일한다.

정규화는 header matching에만 사용한다. `RowData.headerNames()`, map-mode의 key와 오류 리포트에서
원본과 정규화 값 중 무엇을 노출할지는 구현 전에 contract test로 확정한다. 기본 원칙은 사용자가
업로드한 원본 header를 진단 정보에 보존하는 것이다.

### 완료 기준

- Excel과 CSV Reader에 동일한 `headerNormalizer(UnaryOperator<String>)` API가 있다.
- header 이름, alias, required/strict header, selected map column 비교에 일관되게 적용된다.
- normalization 이후 중복은 `FIRST`, `LAST`, `FAIL` 정책을 그대로 따른다.
- 원본 header와 정규화된 비교 값이 오류 메시지에서 혼동되지 않는다.
- identity, trim, case-folding, Unicode NFC와 충돌 사례를 테스트한다.

## 향후 검토 후보

아래 항목은 현재 구현 순서와 우선순위에서 제외한다. 다른 제안에 흡수된 내용은 별도 기능으로
중복 구현하지 않고, 보류 항목은 반복되는 실제 요구가 확인될 때만 다시 검토한다.

### 다른 제안에 흡수

- `InputStream`, `Path`, stream supplier와 Spring adapter 경계는 제안 2에서 다룬다.
- 최대 읽기 오류 수는 제안 5의 `maxErrors()`에서 다룬다.
- 읽기 lifecycle과 조기 종료는 제안 6의 callback API와 `readWhile()`에서 다룬다.
- Writer 입력과 Stream 소유권은 제안 7에서 다룬다.

### 독립 검토 1: 다중 sheet 일괄 읽기

하나의 workbook에서 서로 다른 DTO/schema를 가진 여러 sheet를 실제로 import해야 하는 요구가 생기면
별도 설계 제안으로 승격한다.

```java
ExcelWorkbookReader.create()
    .sheet("Users", userSchema, userConsumer)
    .sheet("Orders", orderSchema, orderConsumer)
    .read(input);
```

- sheet name/index 선택, alias와 optional sheet
- 서로 다른 DTO generic 타입을 한 builder에서 표현하는 방법
- sheet별 header와 오류 처리 정책
- 한 sheet 실패 시 나머지 sheet 처리 여부
- workbook을 sheet마다 복사하거나 reopen하지 않는 단일 session

### 독립 검토 2: 안전 제한 점검

새로운 거대 options 객체를 바로 만들지 않고 현재 보호 기능의 누락부터 점검한다.

- 기존 `maxRows()`, `maxErrors()`, Spring upload size 제한의 적용 범위
- 최대 열 수와 cell text 길이
- Apache POI zip bomb 방어 설정과 library 전역 설정 충돌 가능성
- shared strings, 압축 해제 비율과 ZIP entry 크기
- 기존 URL image byte 제한 외 timeout, redirect, content type 정책

기본값을 변경할 때는 정상적인 대용량 파일과 서버 자원 보호 사이의 기준을 benchmark와 fixture로
검증한다.

### 보류

- CSV delimiter/charset 자동 감지: 명시적 `CsvDialect`가 있고 오탐 위험이 있으므로 실제 요구 전에는 추가하지 않는다.
- column converter/error-message DSL: 현재 setter lambda로 표현 가능하며 overload 복잡도가 커질 수 있다.
- password API 재설계: 현재 검증과 secret 노출 방지를 우선 점검하고 새 credentials 타입은 만들지 않는다.
- 상세 progress event: 실제 SSE/UI 요구가 확인되기 전에는 기존 callback을 유지한다.
- Micrometer integration: core 또는 신규 module에 추가하지 않고 애플리케이션에서 callback/Timer로 구성한다.

## 권장 구현 순서

```text
1. compileOnly 정책 문서화와 호환성 matrix 구축
        │
        ▼
2. Reader/Writer options snapshot과 credentials/source 타입
        │
        ▼
3. Reader 실행 API 및 resource lifecycle
        │
        ▼
4. maxErrors 제한
        │
        ▼
5. Java Stream 읽기 API 제거와 `readWhile()` 추가
               │
               ▼
6. Writer `Iterable` 입력과 Stream 소유권 변경
        │
        ▼
7. Header normalization
```

의존성 정책과 호환성 matrix는 다른 public API 개편과 독립적이므로 먼저 완료할 수 있다.
options snapshot은 Reader API와 실행 session 전달 구조를 동시에 정리하는 기반이다. `maxErrors`는
새 Reader lifecycle의 공통 행 처리 상태에서 구현한다. Schema 개선안은 재검토 조건이 충족될 때까지 순서에서 제외한다.

## 버전 및 migration 전략

breaking change를 허용하더라도 사용자가 한 번에 이동할 수 있는 경로는 제공해야 한다.

### 권장 전략

1. 현재 artifact 구조와 `compileOnly` 정책은 유지한다.
2. 기존 API와 새 API를 동시에 오래 유지하지 않고 migration adapter만 제한적으로 둔다.
3. 변경되는 API와 resource ownership 동작표를 작성한다.
4. before/after 예제를 Excel read, CSV read, schema, Spring upload별로 제공한다.
5. OpenRewrite recipe까지 만들 필요는 없지만 기계적으로 바꿀 수 있는 변경은 검색 패턴을 제공한다.

### Migration 표 예시

| Before | After |
|--------|-------|
| 단일 설치 예제에 모든 integration 나열 | Excel/CSV/Validation별 설치 예제 |
| `reader.build(input).read(c)` | `reader.read(input, c)` |
| handler one-shot lifecycle | reader reusable, session per execution |
| 무제한 오류 처리 | 필요 시 `maxErrors(n)` 설정 |
| `readAsStream()` pipeline | callback `read()` 또는 `readWhile()` |

## 테스트와 성능 기준

구조 개편은 기능 테스트 통과만으로 완료로 판단하지 않는다.

### 호환성과 contract 테스트

- Excel/CSV 공통 reader option contract suite
- resource ownership별 close 여부
- 임시 파일 cleanup
- one-shot session과 reusable configuration 구분
- `maxErrors(0)`, 경계값, 제한 초과 시 행 처리 수
- 최소/현재 POI 및 OpenCSV 버전 조합

### 성능 기준

현재 benchmark를 기준선으로 저장하고 최소한 다음을 측정한다.

- Excel/CSV 1만, 10만, 100만 행 throughput
- peak heap과 임시 디스크 사용량
- `read()`와 `readWhile()` 차이
- Stream API 제거 전후 throughput과 peak memory
- `countRows()` 사전 scan 비용
- Path 입력과 InputStream 입력의 복사 비용

허용 regression은 변경 전에 수치로 정한다. 예를 들어 일반 `read()` throughput 5% 이내,
peak memory 10% 이내처럼 기준을 명시하고 CI benchmark는 변동성을 고려해 별도로 운영한다.

## 우선순위 요약

| 우선순위 | 제안 | 사용자 가치 | 구현 위험 | Breaking 영향 |
|----------|------|-------------|-----------|---------------|
| 1 | 선택적 의존성 정책/호환성 matrix | 높음 | 낮음~중간 | 없음 |
| 2 | Reader 실행 API | 매우 높음 | 높음 | 매우 큼 |
| 3 | Reader/Writer options snapshot | 높음 | 중간 | 낮음~중간 |
| 보류 | Schema immutable mapping 개선 | 불확실 | 높음 | 큼 |
| 5 | `maxErrors()` | 중간~높음 | 낮음~중간 | 없음 |
| 6 | Stream API 제거와 `readWhile()` | 높음 | 중간 | 큼 |
| 7 | Writer `Iterable` 입력/Stream 소유권 | 중간 | 낮음 | 동작 변경 |
| 8 | Header normalization | 중간~높음 | 낮음~중간 | 없음 |
| 향후 검토 | 다중 sheet, 안전 제한 | 요구 확인 필요 | 항목별 상이 | 미정 |

가장 먼저 할 수 있는 작업은 현재 의존성 정책을 문서와 CI의 명시적인 지원 계약으로 만드는 것이다.
가장 큰 public API 결정은 Reader lifecycle이며, 이후 schema, Spring adapter, 오류 모델의 형태를 좌우한다.
