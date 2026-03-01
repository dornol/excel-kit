# excel-kit 프로젝트 분석 및 개선 계획

## 1. 프로젝트 개요

Apache POI 기반 경량 Excel(.xlsx)/CSV 스트리밍 읽기/쓰기 유틸리티 라이브러리.
Fluent DSL API, Bean Validation 연동, 대용량 파일 스트리밍 처리를 지원한다.

- **버전:** 0.1.4
- **빌드:** Gradle (Kotlin DSL), Java
- **배포:** Maven Central (Vanniktech plugin)
- **라이선스:** MIT

---

## 2. 현재 아키텍처

```
io.github.dornol.excelkit
├── shared/          # 공통 유틸 (CellData, ReadResult, TempResource*)
├── excel/           # Excel 읽기/쓰기 (13개 파일, ~1,400줄)
└── csv/             # CSV 읽기/쓰기 (8개 파일, ~500줄)
```

**핵심 설계 패턴:**
- Fluent Builder (ExcelColumnBuilder, ExcelReadColumnBuilder)
- Streaming Write (SXSSFWorkbook)
- SAX Event Parsing (XSSFSheetXMLHandler)
- AutoCloseable 리소스 관리
- Strategy (ExcelDataType enum)

---

## 3. 강점

| 항목 | 평가 |
|------|------|
| 메모리 효율성 | SXSSFWorkbook(쓰기), SAX(읽기) 사용으로 대용량 처리 우수 |
| API 설계 | Fluent DSL이 직관적이고 사용하기 편리함 |
| 타입 안전성 | ExcelDataType enum + 제네릭으로 컴파일타임 타입 체크 |
| 테스트 커버리지 | 12개 테스트 파일, ~2,000줄으로 핵심 기능 대부분 커버 |
| 문서화 | Javadoc이 대부분의 public API에 존재 |
| 리소스 관리 | TempResourceContainer로 임시 파일 자동 정리 |
| 의존성 관리 | compileOnly로 사용자가 의존성 버전을 제어 가능 |

---

## 4. 개선 항목

### 4.1 [높음] ExcelWriter의 최소 컬럼 제한 완화

**현재 문제:**
```java
// ExcelWriter.java:197
if (this.columns.size() < 2) {
    throw new IllegalStateException("columns setting required");
}
```
컬럼이 1개인 경우에도 유효한 Excel 파일이지만, 현재 2개 미만이면 예외를 발생시킨다.

**개선안:** 최소 1개 이상으로 변경.

---

### 4.2 [높음] CellData의 null 체크 로직 불일치

**현재 문제:**
```java
// CellData.java:48-55 (compact constructor)
public CellData {
    if (formattedValue == null) {
        formattedValue = "";  // null → "" 변환
    }
}

// CellData.java:67
public Number asNumber(Locale locale) {
    if (formattedValue == null || formattedValue.isBlank()) {  // null 체크 불필요
        return null;
    }
```
compact constructor에서 `null`을 `""`로 변환하고 있으므로, 이후 모든 메서드에서 `formattedValue == null` 체크가 불필요하다. `isNull()` 메서드도 항상 `false`를 반환하게 되어 dead code가 된다.

**개선안:**
- 모든 메서드에서 불필요한 null 체크 제거
- `isNull()` 메서드의 Javadoc에 항상 false를 반환한다는 점을 명시하거나, 생성 시 null 허용 여부를 재설계

---

### 4.3 [높음] Excel/CSV ReadHandler 중복 코드 제거

**현재 문제:**
`ExcelReadHandler`와 `CsvReadHandler`에 거의 동일한 코드가 반복된다:
- 생성자의 파라미터 검증 로직
- `validateIfNeeded()` 메서드
- 매핑 실패 시 에러 메시지 생성 패턴
- `read()` 내부의 매핑 + 검증 + 결과 전달 흐름

**개선안:** 공통 추상 클래스 `AbstractReadHandler<T>` 추출:
```java
public abstract class AbstractReadHandler<T> extends TempResourceContainer {
    // 공통 필드: columns, instanceSupplier, validator, headerNames
    // 공통 메서드: validateIfNeeded(), mapValuesToInstance()
    // 추상 메서드: read(Consumer<ReadResult<T>>)
}
```

---

### 4.4 [높음] ExcelCursor / CsvCursor 중복 제거

**현재 문제:**
`ExcelCursor`와 `CsvCursor`는 동일한 필드(rowOfSheet, currentTotal)와 메서드를 가진 거의 복사-붙여넣기 수준의 클래스이다.

**개선안:** `shared` 패키지에 단일 `Cursor` 클래스 통합 또는 공통 인터페이스 추출.

---

### 4.5 [중간] CsvReadHandler 인코딩 누락

**현재 문제:**
```java
// CsvReadHandler.java:87
try (CSVReader reader = new CSVReader(new FileReader(getTempFile().toFile()))) {
```
`FileReader`는 시스템 기본 인코딩을 사용한다. `CsvWriter`에서는 UTF-8로 작성하지만, 읽기 쪽에서는 인코딩이 보장되지 않는다.

**개선안:**
```java
new CSVReader(new InputStreamReader(
    Files.newInputStream(getTempFile()), StandardCharsets.UTF_8))
```

---

### 4.6 [중간] CsvReadHandler 에러 메시지 오류

**현재 문제:**
```java
// CsvReadHandler.java:119
throw new IllegalStateException("Failed to read excel", e);  // "excel" → "csv"
```

**개선안:** 에러 메시지를 `"Failed to read CSV"`로 수정.

---

### 4.7 [중간] CsvWriter의 CSV 헤더 이스케이프 누락

**현재 문제:**
```java
// CsvWriter.java:105-107
writer.println(columns.stream()
    .map(CsvColumn::getName)
    .reduce((a, b) -> a + "," + b).orElse(""));
```
헤더명에 쉼표, 따옴표 등 특수문자가 포함될 경우 CSV 표준을 위반한다. 데이터 행에는 `escapeCsv()`를 적용하면서 헤더에는 적용하지 않고 있다.

**개선안:** 헤더도 `escapeCsv()` 처리 적용.

---

### 4.8 [중간] CsvWriter의 String 연결 성능

**현재 문제:**
```java
// CsvWriter.java:114-118
String line = columns.stream()
    .map(col -> col.applyFunction(row, cursor))
    .map(CsvWriter::escapeCsv)
    .reduce((a, b) -> a + "," + b)  // O(n^2) 문자열 연결
    .orElse("");
```

**개선안:** `Collectors.joining(",")` 사용으로 효율적인 문자열 결합.

---

### 4.9 [중간] ExcelWriter에서 title 메서드 중복

**현재 문제:**
`title(String)`, `title(String, int)`, `title(String, int, IndexedColors)` 세 메서드의 본문이 거의 동일하며, 중복된 null 체크 로직이 반복된다.

**개선안:** 하나의 private 메서드로 위임:
```java
public ExcelWriter<T> title(String title) {
    return title(title, 0, IndexedColors.BLACK);
}
public ExcelWriter<T> title(String title, int fontSize) {
    return title(title, fontSize, IndexedColors.BLACK);
}
public ExcelWriter<T> title(String title, int fontSize, IndexedColors color) {
    if (this.title != null) {
        throw new IllegalStateException("title setting already exists");
    }
    this.title = title;
    this.titleStyle = titleStyle(this.wb, HorizontalAlignment.CENTER, color, fontSize);
    return this;
}
```

---

### 4.10 [중간] ExcelReadHandler에서 시트 롤오버 시 title 처리 버그

**현재 문제:**
```java
// ExcelWriter.java:265-269
if (isOverMaxRows()) {
    turnOverSheet();
    setSheetTitle();    // title이 null이어도 호출됨
    setColumnHeaders();
}
```
`setSheetTitle()`은 `title != null` 체크 없이 호출된다. 다만 `setSheetTitle()` 내부에서 `title` 필드를 직접 사용하므로 title이 null이면 NPE 가능성이 있다.

반면, 최초 시트 생성 시에는:
```java
// ExcelWriter.java:204
if (this.title != null) {
    setSheetTitle();
}
```
처럼 null 체크를 수행한다.

**개선안:** 시트 롤오버 로직에도 동일하게 `title != null` 분기 적용:
```java
if (isOverMaxRows()) {
    turnOverSheet();
    if (this.title != null) {
        setSheetTitle();
    }
    setColumnHeaders();
}
```
또한 `cursor.initRow()` 시 title 존재 여부에 따라 시작 행을 올바르게 설정해야 한다.

---

### 4.11 [중간] ExcelStyleSupporter의 CellStyle 캐싱 미적용

**현재 문제:**
모든 컬럼마다 `cellStyle()`, `headerStyle()` 등을 새로 생성한다. SXSSFWorkbook은 스타일 개수에 제한이 있으며(최대 ~64,000개), 동일한 포맷의 컬럼이 많을 경우 불필요하게 스타일이 누적된다.

**개선안:** 동일한 (alignment, format) 조합에 대해 CellStyle 캐싱:
```java
private final Map<String, CellStyle> styleCache = new HashMap<>();
```

---

### 4.12 [낮음] Locale 하드코딩

**현재 문제:**
```java
// CellData.java:91
public Number asNumber() {
    return asNumber(Locale.KOREA);  // 한국 로케일 하드코딩
}
```

**개선안:** `CellData` 생성 시 또는 `ExcelReader` 수준에서 Locale을 설정할 수 있도록 옵션 제공. 기본값은 `Locale.KOREA`를 유지하되, 국제화 대비.

---

### 4.13 [낮음] ExcelReadHandler의 ZipSecureFile/IOUtils 전역 설정

**현재 문제:**
`ExcelReadHandler` Javadoc에서 언급된 `ZipSecureFile.setMaxFileCount(10_000_000)`와 `IOUtils.setByteArrayMaxOverride(2_000_000_000)` 설정이 전역 상태를 변경한다. 이는 동일 JVM 내 다른 POI 사용 코드에 부작용을 줄 수 있다. 다만 현재 코드에서는 실제로 호출하는 부분이 보이지 않아, 사용자가 수동으로 해야 하는 것인지 불명확하다.

**개선안:**
- 실제로 필요하다면 읽기 시작 시 설정하고, 완료 후 원래 값으로 복원
- 또는 설정 여부를 Builder 옵션으로 제공

---

### 4.14 [낮음] 멀티시트 읽기 미지원

**현재 문제:**
```java
// ExcelReadHandler.java:110
try (InputStream sheet = reader.getSheetsData().next()) {  // 첫 번째 시트만
```

**개선안:** 시트 인덱스 또는 시트명 지정 옵션 추가. `readAll()` 메서드로 모든 시트 순회 지원.

---

### 4.15 [낮음] 예외 계층 구조 설계

**현재 문제:**
대부분의 예외가 `IllegalStateException`, `IllegalArgumentException` 같은 기본 예외로 처리된다. `TempResourceCreateException`만 커스텀 예외이다.

**개선안:** 라이브러리 전용 예외 계층 도입:
```
ExcelKitException (추상 베이스)
├── ExcelWriteException
├── ExcelReadException
├── CsvWriteException
├── CsvReadException
└── TempResourceException
```
사용자가 라이브러리 예외만 선택적으로 catch 할 수 있게 된다.

---

### 4.16 [낮음] ExcelWriter가 AutoCloseable 미구현

**현재 문제:**
`ExcelWriter`는 내부적으로 `SXSSFWorkbook`을 생성하지만 `AutoCloseable`을 구현하지 않는다. `write()` 호출 전에 예외가 발생하면 workbook 리소스가 누수될 수 있다.

**개선안:** `AutoCloseable` 구현 또는 write() 시도 중 예외 발생 시 workbook 정리 로직 추가.

---

### 4.17 [낮음] README 개선

**현재 문제:**
- CSV 읽기 기능이 이미 구현되어 있지만, README에는 "CSV는 쓰기만 지원"으로 표기됨
- 영문 README 부재 (Maven Central 배포 라이브러리로서 국제 사용자 접근성 제한)

**개선안:**
- CSV 읽기 관련 문서 업데이트
- 영문 README 추가 또는 bilingual 구성

---

## 5. 개선 우선순위 로드맵

### Phase 1: 버그 수정 및 안정성 (즉시)
1. 시트 롤오버 시 title null 체크 누락 수정 (4.10)
2. CsvReadHandler 인코딩 지정 (4.5)
3. CsvReadHandler 에러 메시지 오류 수정 (4.6)
4. CsvWriter 헤더 이스케이프 누락 수정 (4.7)
5. CellData null 체크 로직 정리 (4.2)

### Phase 2: 코드 품질 개선 (단기)
6. ExcelCursor/CsvCursor 통합 (4.4)
7. ReadHandler 공통 코드 추출 (4.3)
8. ExcelWriter title 메서드 중복 제거 (4.9)
9. CsvWriter String 연결 최적화 (4.8)
10. ExcelWriter 최소 컬럼 제한 완화 (4.1)

### Phase 3: 설계 개선 (중기)
11. 예외 계층 구조 도입 (4.15)
12. CellStyle 캐싱 (4.11)
13. ExcelWriter AutoCloseable 구현 (4.16)
14. Locale 설정 옵션화 (4.12)

### Phase 4: 기능 확장 (장기)
15. 멀티시트 읽기 지원 (4.14)
16. ZipSecureFile 설정 관리 (4.13)
17. README 갱신 및 영문화 (4.17)

---

## 6. 요약

excel-kit은 잘 설계된 라이브러리로, 대용량 Excel/CSV 처리라는 핵심 목표를 효과적으로 달성하고 있다. Fluent API, 스트리밍 처리, Bean Validation 연동 등 실무에서 필요한 기능을 갖추고 있으며, 테스트 커버리지도 양호하다.

가장 시급한 개선 사항은 **시트 롤오버 시 title null 체크 버그**(4.10)와 **CsvReadHandler 인코딩 문제**(4.5)이며, 중기적으로는 **중복 코드 통합**(4.3, 4.4)과 **예외 계층 설계**(4.15)가 라이브러리의 완성도를 높이는 데 기여할 것이다.
