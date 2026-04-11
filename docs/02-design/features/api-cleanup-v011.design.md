---
template: design
version: 1.2
feature: api-cleanup-v011
project: excel-kit
project-version: 0.10.0 → 0.11.0
author: DongHyeok Kim
date: 2026-04-12
status: Draft
---

# api-cleanup-v011 Design Document

> **Summary**: Java 17 기반 Public API 정리 설계. ExcelWriter 생성자 → Builder, FileHandler sealed interface, Reader column unify (v0.10.0 Writer 정렬과 대칭), Map Writer/Reader 4개 클래스 → 정적 팩토리 흡수.
>
> **Project**: excel-kit
> **Version**: 0.10.0 → 0.11.0
> **Author**: DongHyeok Kim
> **Date**: 2026-04-12
> **Status**: Draft
> **Planning Doc**: [api-cleanup-v011.plan.md](../../01-plan/features/api-cleanup-v011.plan.md)

---

## 1. Overview

### 1.1 Design Goals

1. v0.10.0 "Unify column API across all writers" 작업의 **Reader 쪽 완결**
2. 신규 설정 옵션 추가 시 생성자 오버로드 폭발 방지 (Builder 도입)
3. Excel/CSV 핸들러 다형적 처리 가능 (Spring Controller 등)
4. Map Writer/Reader 의 "별도 클래스로 존재할 이유 없음" 해소
5. v1.0.0 안정화 전 마지막 breaking window — 확실히 끊고 가기

### 1.2 Design Principles

- **기존 의도 존중**: v0.10.0 Writer 가 `Consumer<Builder> configurer` 패턴으로 "체인 중단 없는 Builder 접근" 을 확립함. Reader 도 동일 패턴으로 정렬.
- **Java 17 네이티브 기능 적극 활용**: `sealed interface`.
- **즉시 삭제**: 저자 외 사용자 없음 → deprecation 경유 없이 한 번에 삭제. CHANGELOG 에 Before/After 표만 기록.
- **최소 surface**: 새 public API 는 "꼭 필요한 것만" — 새 기능 추가 금지.

---

## 2. Architecture

### 2.1 Type Hierarchy (After)

```
shared/
├── FileHandler                      <<sealed interface>>  [NEW]
│   └── write(OutputStream) throws IOException
│
excel/
├── ExcelHandler final implements FileHandler
│   ├── write(OutputStream) throws IOException            [renamed from consumeOutputStream]
│   └── consumeOutputStreamWithPassword(...)              [kept — Excel-only]
│
├── ExcelWriter<T>
│   ├── (no public constructors)                          [5개 생성자 삭제]
│   ├── static Builder<T> builder()                       [NEW]
│   ├── static ExcelWriter<Map<String,Object>> forMap(String...) [NEW]
│   └── static ExcelWriter<Map<String,Object>> forMap(String[], Consumer<Builder>...) [NEW]
│
├── ExcelReader<T>
│   ├── column(setter)                                    [NEW: returns Reader]
│   ├── column(setter, Consumer<Builder>)                 [NEW]
│   ├── column(name, setter)                              [replaces Builder-returning version]
│   ├── column(name, setter, Consumer<Builder>)           [NEW]
│   ├── columnAt(idx, setter)                             [unchanged]
│   ├── columnAt(idx, setter, Consumer<Builder>)          [NEW overload]
│   ├── skipColumn() / skipColumns(int)                   [unchanged]
│   ├── addColumn(...) × 2                                [DELETED]
│   └── columnAtBuilder(...)                              [DELETED]
│
├── ExcelMapWriter                                        [FILE DELETED]
└── ExcelMapReader                                        [unchanged — v0.12.0 scope]

csv/
├── CsvHandler final implements FileHandler
│   └── write(OutputStream) throws IOException            [renamed from consumeOutputStream]
│
├── CsvWriter<T>
│   └── static CsvWriter<Map<String,Object>> forMap(String...) [NEW]
│
├── CsvReader<T>     (ExcelReader 와 동일 변경)
├── CsvMapWriter                                          [FILE DELETED]
└── CsvMapReader                                          [unchanged — v0.12.0 scope]
```

### 2.2 Package Placement

`FileHandler` sealed interface 를 어디에 둘지 결정:

| 옵션 | 위치 | 장점 | 단점 |
|------|------|------|------|
| **A** | `io.github.dornol.excelkit.shared` | 이미 `CellData`, `RowData`, `TempResourceContainer` 등 공유 타입이 있는 패키지 | `permits` 가 cross-package 로 가리킴 (Java 17+ 허용) |
| B | `io.github.dornol.excelkit` (루트) | 최상위 노출 | 기존 최상위 패키지에 타입 없음 — 규칙 바꿈 |

**선택: 옵션 A** (`shared` 패키지) — 기존 규칙과 일관성.

### 2.3 sealed interface cross-package 유의점

Java 17 sealed interface 는 `permits` 가 다른 패키지에 있어도 **같은 모듈 안이면 OK**. excel-kit 은 단일 module 이므로 문제 없음. 단, `permits` 절에 전체 FQCN 필요:

```java
package io.github.dornol.excelkit.shared;

public sealed interface FileHandler
    permits io.github.dornol.excelkit.excel.ExcelHandler,
            io.github.dornol.excelkit.csv.CsvHandler {
    void write(OutputStream out) throws IOException;
}
```

`ExcelHandler` / `CsvHandler` 는 `non-sealed` 가 아닌 `final` 로 선언하여 외부 확장 완전 차단 (현재도 생성자가 package-private 이라 사실상 불가능하지만 명시적으로).

---

## 3. Type Model

### 3.1 ExcelWriter Builder (T1)

현재 `ExcelWriter` 생성자 5개가 수용하는 파라미터는 **단 3개**:
- `ExcelColor color` (default: `WHITE`)
- `int maxRows` (default: `1_000_000`)
- `int rowAccessWindowSize` (default: `1000`)

나머지 설정 (`rowHeight`, `autoFilter`, `sheetName`, `password`, `headerFontName`, ...) 은 이미 fluent 메서드로 노출되어 있음.

**설계**: `ExcelWriter` 내부 정적 중첩 클래스 `Builder<T>`. 기존 5개 public 생성자 즉시 삭제.

```java
public class ExcelWriter<T> {

    // 기존 5개 public 생성자 전부 삭제
    // (ExcelColor, int, int) / (ExcelColor, int) / (ExcelColor) / (int) / ()

    // 유일한 생성자 — package-private, Builder 전용
    ExcelWriter(Builder<T> builder) {
        this.wb = new SXSSFWorkbook(builder.rowAccessWindowSize);
        this.maxRows = builder.maxRows;
        this.headerColor = new XSSFColor(new byte[]{
            (byte) builder.color.getR(),
            (byte) builder.color.getG(),
            (byte) builder.color.getB()
        });
        this.headerStyle = ExcelStyleSupporter.headerStyle(wb, headerColor);
    }

    public static <T> Builder<T> builder() {
        return new Builder<>();
    }

    public static final class Builder<T> {
        private ExcelColor color = ExcelColor.WHITE;
        private int maxRows = 1_000_000;
        private int rowAccessWindowSize = 1000;

        private Builder() {}

        public Builder<T> color(ExcelColor color) {
            this.color = color;
            return this;
        }

        public Builder<T> maxRows(int maxRows) {
            if (maxRows <= 0) throw new IllegalArgumentException("maxRows must be positive");
            this.maxRows = maxRows;
            return this;
        }

        public Builder<T> rowAccessWindowSize(int size) {
            if (size <= 0) throw new IllegalArgumentException("rowAccessWindowSize must be positive");
            this.rowAccessWindowSize = size;
            return this;
        }

        public ExcelWriter<T> build() {
            return new ExcelWriter<>(this);
        }
    }
}
```

**사용 예**:
```java
// Before
new ExcelWriter<>(ExcelColor.STEEL_BLUE, 500_000, 500)

// After
ExcelWriter.<User>builder()
    .color(ExcelColor.STEEL_BLUE)
    .maxRows(500_000)
    .rowAccessWindowSize(500)
    .build()
```

**결정**: 별도 `record ExcelWriterConfig` 분리 안 함. 수용 필드 3개뿐이므로 Builder 내부 private 필드로 충분. Plan 의 `ExcelWriterConfig` 언급은 취소됨.

### 3.2 FileHandler sealed interface (T2)

```java
package io.github.dornol.excelkit.shared;

import java.io.IOException;
import java.io.OutputStream;

public sealed interface FileHandler
    permits io.github.dornol.excelkit.excel.ExcelHandler,
            io.github.dornol.excelkit.csv.CsvHandler {

    /**
     * Writes the generated file content to the given output stream.
     * Can only be called once per handler instance.
     */
    void write(OutputStream out) throws IOException;
}
```

**메서드 이름 결정**: `write(OutputStream)` 로 리네이밍. 기존 `consumeOutputStream(OutputStream)` 은 **삭제**.

**이유**:
1. `write` 가 POI/Jackson/기타 자바 생태계 관례와 일치
2. 30자 → 5자로 간결
3. `consumeOutputStream` 이름은 "한 번만 호출 가능" 이라는 의미를 담으려 했으나 javadoc 으로 충분
4. 인터페이스 도입 시 짧은 이름이 더 자연스러움

**구현체 예**:
```java
package io.github.dornol.excelkit.excel;

public final class ExcelHandler implements FileHandler {
    private final SXSSFWorkbook wb;
    private final @Nullable String password;
    private final AtomicBoolean consumed = new AtomicBoolean(false);

    ExcelHandler(SXSSFWorkbook wb) { this(wb, null); }
    ExcelHandler(SXSSFWorkbook wb, @Nullable String password) { ... }

    @Override
    public void write(OutputStream outputStream) throws IOException {
        if (password != null) encryptAndWrite(outputStream, password);
        else writePlain(outputStream);
    }

    // Excel 전용 — FileHandler 인터페이스에 속하지 않음
    public void consumeOutputStreamWithPassword(OutputStream out, String password) throws IOException { ... }
    public void consumeOutputStreamWithPassword(OutputStream out, char[] password) throws IOException { ... }
}
```

### 3.3 IOException throws 정책

새 메서드 `write(OutputStream)` 는 `FileHandler` 인터페이스 계약대로 `throws IOException` 을 선언함.

| Handler | 새 시그니처 |
|---------|-------------|
| `ExcelHandler.write` | `throws IOException` (실제로 던짐) |
| `CsvHandler.write` | `throws IOException` (선언만, 내부는 `CsvWriteException` 으로 래핑되므로 실제로 안 던짐) |

`CsvHandler.write` 는 실제 IOException 을 던지지 않지만 인터페이스 계약상 declared throws 필요. 호출부는 try/catch 또는 throws 선언 필수.

기존 `consumeOutputStream` 은 삭제되었으므로 마이그레이션 = 이름 변경 + (CSV 쪽은) try/catch 추가.

### 3.4 Reader column API (T4)

#### 신규 시그니처 (Excel/Csv 공통 패턴)

```java
// 순차 (다음 컬럼)
public ExcelReader<T> column(BiConsumer<T, CellData> setter);
public ExcelReader<T> column(BiConsumer<T, CellData> setter,
                              Consumer<ExcelReadColumnBuilder<T>> configurer);

// 헤더명 기반
public ExcelReader<T> column(String headerName, BiConsumer<T, CellData> setter);
public ExcelReader<T> column(String headerName, BiConsumer<T, CellData> setter,
                              Consumer<ExcelReadColumnBuilder<T>> configurer);

// 명시 인덱스
public ExcelReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter);
public ExcelReader<T> columnAt(int columnIndex, BiConsumer<T, CellData> setter,
                                Consumer<ExcelReadColumnBuilder<T>> configurer);

// 스킵 (기존 유지)
public ExcelReader<T> skipColumn();
public ExcelReader<T> skipColumns(int count);
```

총 **8개 메서드**. 모두 `ExcelReader<T>` 반환 (체인 유지).

#### Return Type 변경

현재 `column(BiConsumer)` / `column(String, BiConsumer)` 은 **Builder 반환**. 신규 정의는 **Reader 반환**. 같은 이름·파라미터로 반환 타입만 다른 오버로드는 Java erasure 충돌로 **불가능**.

**해결**: 기존 `column(...)` Builder 반환형 2개를 **삭제** 후 같은 이름으로 Reader 반환형 재도입. 함께 삭제되는 것:
- `addColumn(BiConsumer)` — `column(setter)` 로 대체
- `addColumn(String, BiConsumer)` — `column(name, setter)` 로 대체
- `columnAtBuilder(int, BiConsumer)` — `columnAt(idx, setter, cfg)` 로 대체

**Breaking 예시** (CHANGELOG 용):
```java
// Before (v0.10.0)
reader.column(User::setName).required().build(in)       // Builder 체인
reader.addColumn(User::setEmail).build(in)              // 빠른 경로

// After (v0.11.0)
reader.column(User::setName, cfg -> cfg.required()).build(in)
reader.column(User::setEmail).build(in)
```

`example/*.java` 및 `kit/src/test/**/*.java` 에서 기존 호출부 전체 수정 필요 — T4 와 T8/T9 는 **같은 커밋** 에서 진행.

### 3.5 Map Writer 클래스 삭제 + `Writer.forMap()` (T5)

#### 방침

| 대상 | 처리 |
|------|:----:|
| `ExcelMapWriter.java` | **파일 삭제** |
| `CsvMapWriter.java` | **파일 삭제** |
| `ExcelMapReader.java` | **유지** (v0.12.0 범위) |
| `CsvMapReader.java` | **유지** (v0.12.0 범위) |

Reader 쪽은 `XSSFSheetXMLHandler` SAX 콜백 상태 머신 재작성이 필요해 회귀 리스크가 큼. v0.11.0 에서는 건드리지 않음. 결과적으로 v0.11.0 은 "Map Writer 는 `forMap()` 정적 팩토리, Map Reader 는 기존 클래스 생성" 이라는 일시적 비대칭 상태가 되지만, 외부 사용자 없으므로 수용 가능.

#### Writer 구현

```java
// ExcelWriter.java
public static ExcelWriter<Map<String, Object>> forMap(String... columnNames) {
    ExcelWriter<Map<String, Object>> w =
        ExcelWriter.<Map<String, Object>>builder().build();
    for (String name : columnNames) {
        w.column(name, map -> map.get(name));
    }
    return w;
}

@SafeVarargs
public static ExcelWriter<Map<String, Object>> forMap(
        String[] columnNames,
        Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>>... configurers) {
    if (configurers.length > columnNames.length) {
        throw new IllegalArgumentException(
            "configurers length (" + configurers.length + ") exceeds columnNames length (" + columnNames.length + ")");
    }
    ExcelWriter<Map<String, Object>> w =
        ExcelWriter.<Map<String, Object>>builder().build();
    for (int i = 0; i < columnNames.length; i++) {
        String name = columnNames[i];
        Consumer<ExcelColumn.ExcelColumnBuilder<Map<String, Object>>> cfg =
            (i < configurers.length) ? configurers[i] : null;
        w.column(name, map -> map.get(name), cfg);
    }
    return w;
}
```

`CsvWriter.forMap` 은 더 단순 (CsvWriter 컬럼에 별도 Builder 없음):

```java
public static CsvWriter<Map<String, Object>> forMap(String... columnNames) {
    CsvWriter<Map<String, Object>> w = new CsvWriter<>();
    for (String name : columnNames) {
        w.column(name, map -> map.get(name));
    }
    return w;
}
```

#### 이득

- `CsvMapWriter` 의 `dialect`/`delimiter`/`charset`/`bom` shortcut 은 불필요 — `CsvWriter` 가 직접 제공
- 사용자는 `CsvWriter.forMap("a","b").csvInjectionDefense(false).quoting(...).afterData(...)` 등 **CsvWriter 의 모든 메서드** 를 체인 중단 없이 사용 가능 (기존 `CsvMapWriter` 는 4개만 노출했음)

---

## 4. New Public API Specification

### 4.1 변경 요약

| Task | 파일 | 추가 | 삭제 |
|------|------|------|------|
| T1 | `ExcelWriter.java` | `Builder<T>`, `builder()` | 5개 public 생성자 |
| T2 | `shared/FileHandler.java` | sealed interface + `write()` | — |
| T2 | `ExcelHandler.java` | `implements FileHandler`, `write()`, `final` | `consumeOutputStream(OutputStream)` |
| T2 | `CsvHandler.java` | `implements FileHandler`, `write() throws IOException`, `final` | `consumeOutputStream(OutputStream)` |
| T4 | `ExcelReader.java` | 6개 신규 `column*` 메서드 (전부 Reader 반환) | `addColumn(setter)`, `addColumn(name, setter)`, `column(setter)` Builder형, `column(name, setter)` Builder형, `columnAtBuilder(idx, setter)` |
| T4 | `CsvReader.java` | 동일 | 동일 |
| T5 | `ExcelWriter.java` | 2개 `forMap` 정적 팩토리 | — |
| T5 | `CsvWriter.java` | 1개 `forMap` 정적 팩토리 | — |
| T5 | `ExcelMapWriter.java` | — | **파일 전체 삭제** |
| T5 | `CsvMapWriter.java` | — | **파일 전체 삭제** |

### 4.2 범위 외 (v0.12.0)

- `ExcelMapReader`, `CsvMapReader` — 파일·API 변경 없음
- `ExcelReader.forMap()`, `CsvReader.forMap()` — 도입 안 함
- v0.11.0 은 일시적으로 "Map Writer 는 `forMap()` / Map Reader 는 기존 클래스" 비대칭 상태. CHANGELOG 에 "v0.12.0 에서 Map Reader 도 `forMap` 으로 통일 예정" 명시.

---

## 5. Error Handling

### 5.1 IOException 전파

- `FileHandler.write(OutputStream) throws IOException` 이 공식 계약.
- `ExcelHandler` / `CsvHandler` 구현체 모두 동일 시그니처.
- `CsvHandler` 는 실제 IOException 을 던지지 않지만 declared throws 추가 (인터페이스 계약 준수 + 일관성).

### 5.2 Builder Validation

- `Builder.maxRows(int)`: `<= 0` 시 `IllegalArgumentException`
- `Builder.rowAccessWindowSize(int)`: `<= 0` 시 `IllegalArgumentException`
- `Builder.color(ExcelColor)`: null 허용 안 함 → `Objects.requireNonNull`

### 5.3 CHANGELOG Migration Guide

Deprecation 경유 없이 즉시 삭제하므로 IDE 경고로는 안내할 수 없음. 대신 **CHANGELOG 에 Before/After 마이그레이션 표** 를 상세 수록하여 저자·미래 AI 에이전트가 코드 이해·자동 업데이트에 참고할 수 있게 함 (섹션 6 참조).

---

## 6. Migration Matrix

### 6.1 ExcelWriter 생성

| Before (v0.10.0) | After (v0.11.0) |
|-------------------|------------------|
| `new ExcelWriter<User>()` | `ExcelWriter.<User>builder().build()` |
| `new ExcelWriter<>(ExcelColor.STEEL_BLUE)` | `ExcelWriter.builder().color(ExcelColor.STEEL_BLUE).build()` |
| `new ExcelWriter<>(500_000)` | `ExcelWriter.builder().maxRows(500_000).build()` |
| `new ExcelWriter<>(ExcelColor.STEEL_BLUE, 500_000)` | `ExcelWriter.builder().color(...).maxRows(500_000).build()` |
| `new ExcelWriter<>(ExcelColor.STEEL_BLUE, 500_000, 500)` | `ExcelWriter.builder().color(...).maxRows(500_000).rowAccessWindowSize(500).build()` |

### 6.2 Handler write

| Before | After |
|--------|-------|
| `handler.consumeOutputStream(out)` | `handler.write(out)` |
| `excelHandler.consumeOutputStreamWithPassword(out, "pw")` | `excelHandler.consumeOutputStreamWithPassword(out, "pw")` *(유지 — Excel only)* |

### 6.3 Reader column binding

| Before | After |
|--------|-------|
| `reader.addColumn(User::setName)` | `reader.column(User::setName)` |
| `reader.addColumn("Name", User::setName)` | `reader.column("Name", User::setName)` |
| `reader.column(User::setName).required()` | `reader.column(User::setName, cfg -> cfg.required())` |
| `reader.columnAt(2, User::setAge)` | `reader.columnAt(2, User::setAge)` *(동일)* |
| `reader.columnAtBuilder(2, User::setAge).required()` | `reader.columnAt(2, User::setAge, cfg -> cfg.required())` |

### 6.4 Map Writer

| Before | After |
|--------|-------|
| `new ExcelMapWriter("a", "b")` | `ExcelWriter.forMap("a", "b")` |
| `new ExcelMapWriter(writer, new String[]{"a","b"}, c1, c2)` | `ExcelWriter.forMap(new String[]{"a","b"}, c1, c2)` |
| `new CsvMapWriter("a", "b")` | `CsvWriter.forMap("a", "b")` |
| `csvMapWriter.dialect(EXCEL).delimiter(';')` | `CsvWriter.forMap("a","b").dialect(EXCEL).delimiter(';')` |
| **ExcelMapReader 사용 코드** | **변경 없음** *(v0.12.0 예정)* |
| **CsvMapReader 사용 코드** | **변경 없음** *(v0.12.0 예정)* |

---

## 7. Test Plan

### 7.1 신규 테스트 (각 Task 별 `@Nested`)

`kit/src/test/java/io/github/dornol/excelkit/excel/` 및 `csv/` 아래:

| Task | 테스트 클래스 | 케이스 |
|------|---------------|--------|
| T1 | `ExcelWriterBuilderTest` | 기본값, color/maxRows/rowAccessWindowSize 개별/조합, 예외 (maxRows<=0 등) |
| T2 | `FileHandlerSealedTest` | `FileHandler` 다형적 사용, `instanceof pattern matching`, cross-package permits |
| T2 | `CsvHandlerThrowsTest` | `write()` 가 선언된 IOException 을 계약대로 전파 |
| T4 | `ExcelReaderColumnUnifiedTest` | 6가지 오버로드 전부 호출, configurer 동작 확인, Builder 종결 불필요 검증 |
| T4 | `CsvReaderColumnUnifiedTest` | 동일 |
| T5 | `ExcelWriterForMapTest` | `forMap(String...)`, `forMap(String[], Consumer[])`, 기존 `ExcelMapWriter` 와 출력 동등성 확인 |
| T5 | `CsvWriterForMapTest` | 동일 |

### 7.2 기존 테스트 처리

- `ExcelMapWriter`, `CsvMapWriter` 테스트 파일: 삭제 (클래스와 함께)
- `ExcelMapReader`, `CsvMapReader` 테스트: 변경 없이 유지 (클래스 자체 유지)
- `ExcelWriter` 기존 생성자 사용 테스트: 전부 `Builder` 로 재작성
- `ExcelReader`/`CsvReader` 의 `addColumn`/`columnAtBuilder` 사용 테스트: 전부 새 API 로 재작성
- `consumeOutputStream` 사용 테스트: `write` 로 리네이밍

### 7.3 example 앱 수동 검증

| Endpoint | 검증 항목 |
|----------|-----------|
| `/showcase/write/*` 전체 | 새 API 로 마이그레이션 후 다운로드 정상 |
| 복잡한 case (`conditional-formatting`, `chart`, `protection`) | Builder + fluent 체인 정상 |

---

## 8. Implementation Order

의존 관계를 고려한 순서:

각 단계는 완료 시점에 **반드시 `./gradlew test` 통과** 해야 함. 이렇게 해야 회귀 발생 위치를 좁힐 수 있음.

1. **T2 (FileHandler sealed interface)** — 타입 시스템 기반, 다른 Task 와 독립
   - `shared/FileHandler.java` 생성
   - `ExcelHandler` / `CsvHandler` 에 `implements FileHandler`, `final`, `write()` 추가
   - 기존 `consumeOutputStream(OutputStream)` 삭제 (`consumeOutputStreamWithPassword` 는 유지)
   - 호출부 전체 (example + tests) 를 `write()` 로 리네이밍
   - CSV 쪽은 try/catch 또는 throws 추가
2. **T1 (ExcelWriter Builder)** — 독립적
   - `ExcelWriter$Builder<T>` 중첩 클래스 추가 + `builder()` 정적 팩토리
   - 기존 생성자 5개 삭제
   - 신규 package-private 생성자 (Builder 전용)
   - 호출부 전체를 `builder()` 로 이전
3. **T5 (Writer forMap 정적 팩토리 + Map 클래스 삭제)**
   - `ExcelWriter.forMap(...)` 2개 오버로드
   - `CsvWriter.forMap(...)` 1개
   - `ExcelMapWriter.java`, `CsvMapWriter.java` 파일 삭제
   - 해당 테스트 파일 삭제
   - 호출부 전체를 `forMap()` 로 이전
4. **T4 (Reader column unify)** — 가장 큰 변경
   - Excel/Csv Reader 양쪽에서 기존 `addColumn`×2, `column`×2 Builder형, `columnAtBuilder` 삭제
   - 신규 6개 메서드 추가 (전부 Reader 반환)
   - **T8 (example 마이그레이션) 을 같은 커밋에서 수행** — 컴파일 에러 방지
5. **T9 (테스트)** — 각 Task 완료 시점에 동반. 단계별 누락 방지
6. **T7 (문서)** — 구현 완료 후
   - CHANGELOG: Breaking 섹션 상단, Before/After 마이그레이션 표
   - README: Installation 버전, Features, 사용법
   - `META-INF/AI.md`, `docs/llms.txt`

### 8.1 예상 파일 변경 범위

| 파일 | Task | 변경 라인 수 예상 |
|------|------|-------------------|
| `shared/FileHandler.java` | T2 | +15 (신규) |
| `excel/ExcelHandler.java` | T2 | ±20 |
| `csv/CsvHandler.java` | T2 | ±10 |
| `excel/ExcelWriter.java` | T1, T5 | +80 / -70 |
| `csv/CsvWriter.java` | T5 | +20 |
| `excel/ExcelReader.java` | T4 | +80 / -90 |
| `csv/CsvReader.java` | T4 | +70 / -70 |
| `excel/ExcelMapWriter.java` | T5 | **파일 삭제 (-89)** |
| `csv/CsvMapWriter.java` | T5 | **파일 삭제 (-107)** |
| `example/**/*.java` | T8 | ±60 (마이그레이션) |
| `kit/src/test/**/*.java` | T9 | ±400 (재작성 + 신규) |

---

## 9. Breaking Change Summary (CHANGELOG 용)

v0.11.0 은 **전면 breaking** 릴리스. deprecation 경유 없이 즉시 삭제.

### Removed (삭제)

1. `ExcelWriter` public 생성자 5개 → `ExcelWriter.builder()` 사용
2. `ExcelHandler.consumeOutputStream(OutputStream)` → `write(OutputStream)` 사용
3. `CsvHandler.consumeOutputStream(OutputStream)` → `write(OutputStream)` 사용 (+`throws IOException`)
4. `ExcelReader.addColumn(BiConsumer)`, `addColumn(String, BiConsumer)` → `column(setter)`, `column(name, setter)` 사용
5. `ExcelReader.columnAtBuilder(int, BiConsumer)` → `columnAt(idx, setter, cfg)` 사용
6. `ExcelReader.column(BiConsumer)`, `column(String, BiConsumer)` **Builder 반환형** → 같은 이름으로 **Reader 반환형** 재도입 (Builder 접근은 `column(setter, cfg -> ...)` 형태)
7. `CsvReader` 에서 위 4~6 동일하게 적용
8. `ExcelMapWriter` 클래스 파일 → `ExcelWriter.forMap(...)` 사용
9. `CsvMapWriter` 클래스 파일 → `CsvWriter.forMap(...)` 사용

### Added (신규)

1. `shared/FileHandler` sealed interface + `write(OutputStream) throws IOException`
2. `ExcelWriter$Builder<T>` + `ExcelWriter.builder()` 정적 팩토리
3. `ExcelWriter.forMap(String...)` + `ExcelWriter.forMap(String[], Consumer<Builder>...)`
4. `CsvWriter.forMap(String...)`
5. `ExcelReader` / `CsvReader` 신규 6개 column 메서드 (전부 Reader 반환)
6. `ExcelHandler` / `CsvHandler` 에 `final` 수식자

### 변경 없음 (v0.12.0 예정)

- `ExcelMapReader`, `CsvMapReader` — `Reader.forMap()` 으로 흡수 예정 (SAX 콜백 재작성 필요)
- `ExcelTemplateWriter`, `TemplateListWriter`
- `ColumnStyleConfig`, `ExcelColumnBuilder`, `ColumnConfig` 통합

---

## 10. Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 0.1 | 2026-04-12 | Initial draft — Java 17 확정, Builder/FileHandler/Reader unify/forMap 설계. Map Reader 는 v0.11.0 범위에서 제외로 결정. | DongHyeok Kim |
| 0.2 | 2026-04-12 | 외부 사용자 없음 컨텍스트 반영 — deprecation 경유 전면 제거, 즉시 삭제 방침. T3(Handler interface 추출), T6(Deprecation 전략) 범위에서 삭제. Map Writer 파일 2개 즉시 삭제로 변경. | DongHyeok Kim |
