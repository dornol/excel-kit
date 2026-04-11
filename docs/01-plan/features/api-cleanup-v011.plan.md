---
template: plan
version: 1.2
feature: api-cleanup-v011
project: excel-kit
project-version: 0.10.0 → 0.11.0
author: DongHyeok Kim
date: 2026-04-11
status: Draft
---

# api-cleanup-v011 Planning Document

> **Summary**: Public API 정리 — 생성자 오버로드 축소, Handler 공통 타입 도입, Writer/Reader 네이밍 대칭화 (v0.11.0 breaking change)
>
> **Project**: excel-kit
> **Version**: 0.10.0 → 0.11.0
> **Author**: DongHyeok Kim
> **Date**: 2026-04-11
> **Status**: Draft

---

## Executive Summary

| Perspective | Content |
|-------------|---------|
| **Problem** | v0.10.0에서 Writer만 `column()` 으로 통일되고 Reader는 `addColumn`/`column`(Builder 반환) 이 공존 중. `ExcelWriter` 생성자 5개 오버로드, `ExcelHandler`/`CsvHandler` 공통 타입 부재, `ExcelMapWriter`/`CsvMapWriter`/`ExcelMapReader`/`CsvMapReader` 가 별도 클래스로 존재하며 설정 메서드 노출 수준이 제각각. |
| **Solution** | (1) `ExcelWriter.builder()` 정적 팩토리로 생성자 5개 즉시 삭제, (2) `sealed interface FileHandler` + `write(OutputStream)` 신규 이름 (기존 `consumeOutputStream` 삭제), (3) Reader `column()` API 를 v0.10.0 Writer 통합 규칙에 맞춰 즉시 정렬 (`addColumn`/`columnAtBuilder`/기존 Builder 반환형 삭제, `Consumer<Builder> configurer` 오버로드 추가), (4) `ExcelMapWriter`/`CsvMapWriter` 클래스 삭제 → `Writer.forMap()` 정적 팩토리로 대체. Map **Reader** 흡수는 SAX 콜백 재작성 리스크로 v0.12.0 으로 연기. |
| **Function/UX Effect** | IDE 자동완성에서 의도가 명확히 드러나고, 신규 설정 옵션 추가 시 생성자 폭발 없이 확장 가능. 호출부에서 Excel/CSV 다형적 처리 가능. |
| **Core Value** | "Fluent API"라는 라이브러리 정체성에 맞춰 진입점·체이닝·확장성의 일관성을 확보. v1.0.0 안정화 전 마지막 breaking window. |

---

## 1. Overview

### 1.1 Purpose

excel-kit 공개 API의 **일관성·확장성·발견가능성** 개선. 기능 추가가 아니라 **기존 API 정리**가 목적.

### 1.2 Background

- v0.10.0까지 기능이 빠르게 추가되면서 (`column` Unify, AI context docs 등) 진입점과 콜백 타입이 별개로 성장함.
- 특히 `ExcelWriter` 생성자는 현재 5개 오버로드(`ExcelWriter.java:62,75,84,93,100`) — 다음 옵션(예: header 폰트, window size 별도 설정) 추가 시 2의 n제곱으로 폭발.
- `ExcelHandler`(class)와 `CsvHandler`(class, `extends TempResourceContainer`)는 동일한 `consumeOutputStream(OutputStream)` 역할인데 공통 타입이 없어 Spring Controller 등에서 다형적 사용 불가.
- v1.0.0 안정화 전 마지막 breaking change 기회로 v0.11.0 타겟.

### 1.3 Related Documents

- 이전 릴리스 분석: 대화 로그 (excel-kit API 서베이)
- 참조 파일 (주요):
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelWriter.java`
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelHandler.java`
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvHandler.java`
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelReader.java`
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvReader.java`
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelMapWriter.java`
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvMapWriter.java`

---

## 2. Scope

### 2.1 In Scope

- [ ] **[T1] ExcelWriter Builder 도입**: 기존 생성자 5개 **삭제**, `ExcelWriter.builder()` 정적 팩토리 + 내부 `Builder<T>` 클래스 추가. Builder 는 `color`/`maxRows`/`rowAccessWindowSize` 3개 필드 수용. package-private 생성자 1개만 유지 (Builder 전용).
- [ ] **[T2] FileHandler sealed interface + write 리네이밍**:
  - `shared/FileHandler.java` 신규 생성 — `sealed interface FileHandler permits ExcelHandler, CsvHandler`
  - 공통 메서드: `void write(OutputStream out) throws IOException`
  - `ExcelHandler` / `CsvHandler` 에 `implements FileHandler` 추가, `final` 로 선언
  - 기존 `consumeOutputStream(OutputStream)` **삭제** → `write(OutputStream)` 로 리네이밍
  - `ExcelHandler.consumeOutputStreamWithPassword(...)` 은 Excel 전용 메서드로 유지 (FileHandler 인터페이스 밖)
- [ ] **[T4] Reader column API 즉시 정렬**:
  - `addColumn(setter)`, `addColumn(name, setter)` **삭제**
  - `columnAtBuilder(idx, setter)` **삭제**
  - 기존 `column(setter)` / `column(name, setter)` Builder 반환형 **삭제** (erasure 충돌)
  - 신규 8개 메서드 추가 (전부 `Reader<T>` 반환):
    - `column(setter)`, `column(setter, Consumer<Builder>)`
    - `column(name, setter)`, `column(name, setter, Consumer<Builder>)`
    - `columnAt(idx, setter)`, `columnAt(idx, setter, Consumer<Builder>)`
  - `skipColumn()`, `skipColumns(int)` 유지
  - Excel/Csv Reader 양쪽 동일 적용
  - **결과**: `column` / `columnAt` / `skipColumn` / `skipColumns` 4종
- [ ] **[T5] Map Writer 클래스 삭제 + `Writer.forMap()` 정적 팩토리**:
  - `ExcelMapWriter.java` **파일 삭제**
  - `CsvMapWriter.java` **파일 삭제**
  - `ExcelWriter.forMap(String...)` + `ExcelWriter.forMap(String[], Consumer<Builder>...)` 추가
  - `CsvWriter.forMap(String...)` 추가
  - **Map Reader 는 범위 외**: `ExcelMapReader` / `CsvMapReader` 는 v0.11.0 에서 변경 없음 (SAX 콜백 재작성 리스크로 v0.12.0 에서 별도 작업)
- [ ] **[T7] 문서 업데이트**: CHANGELOG (Breaking 섹션 상단 강조) / README Installation 버전, Features, 사용법 / `META-INF/AI.md` / `docs/llms.txt`. Migration 가이드 (Before/After) 는 CHANGELOG 에 전부 기록 (미래 저자·AI 에이전트가 코드 이해할 때 유용).
- [ ] **[T8] example 앱 마이그레이션**: `example/` 내 모든 showcase 를 새 API 로 이전. T4 가 컴파일 에러를 유발하므로 **T4 와 같은 커밋** 에서 동반 수정 필수.
- [ ] **[T9] 테스트**: 기존 테스트 전부 새 API 로 재작성. `Builder`, `FileHandler` sealed, Reader column unify, `Writer.forMap()` 각각 `@Nested` 로 케이스 추가.

### 2.2 Out of Scope

- `ColumnStyleConfig` / `ExcelColumnBuilder` / `ColumnConfig` 통합 — v0.12.0 이후로 미룸.
- **Map Reader 흡수** (`ExcelMapReader` / `CsvMapReader` → `Reader.forMap()`) — SAX 콜백 재작성 리스크로 v0.12.0 에서 별도 작업.
- 새로운 기능 추가 (차트, 피벗 등).
- Apache POI 버전 업그레이드.
- Template Writer (`ExcelTemplateWriter`, `TemplateListWriter`) 개편 — 현재 형태 유지.
- `CsvWriter` 생성자 변경 — 현재 no-arg 만 존재해 손댈 이유 없음 (T1 은 `ExcelWriter` 전용).

---

## 3. Requirements

### 3.1 Functional Requirements

| ID | Requirement | Priority | Status |
|----|-------------|----------|--------|
| FR-01 | `ExcelWriter$Builder<T>` 정적 중첩 클래스 — `color`/`maxRows`/`rowAccessWindowSize` 3개 필드 + 기본값 + 검증 | High | Pending |
| FR-02 | `ExcelWriter.builder()` 정적 팩토리 + package-private 생성자 1개 (Builder 전용). 기존 public 생성자 5개 삭제 | High | Pending |
| FR-03 | `shared/FileHandler` sealed interface — `write(OutputStream) throws IOException` 단일 메서드. `permits ExcelHandler, CsvHandler`, 구현체는 `final` | High | Pending |
| FR-04 | `ExcelHandler.consumeOutputStream(...)` 및 `CsvHandler.consumeOutputStream(...)` 삭제 → `write(OutputStream)` 로 리네이밍. `ExcelHandler.consumeOutputStreamWithPassword(...)` 는 Excel 전용으로 유지 | High | Pending |
| FR-05 | Reader 기존 column 계열 6개 메서드 (`addColumn`×2, `column`×2 Builder형, `columnAt`, `columnAtBuilder`) 삭제. 신규 6개 메서드 추가 (`column` × {pos, pos+cfg, name, name+cfg}, `columnAt` × {idx, idx+cfg}) — 전부 Reader 반환. Excel/Csv 양쪽 | High | Pending |
| FR-06 | `ExcelMapWriter`, `CsvMapWriter` 클래스 파일 삭제. `ExcelWriter.forMap(String...)` + configurer 오버로드, `CsvWriter.forMap(String...)` 정적 팩토리 제공 | High | Pending |
| FR-07 | example 앱은 새 API 로 **동일 커밋** 에서 이전 (컴파일 에러 방지) | High | Pending |
| FR-08 | CHANGELOG 에 Breaking 섹션 + Before/After 마이그레이션 표 기록 (저자·AI 가 미래 참고용) | Medium | Pending |

### 3.2 Non-Functional Requirements

| Category | Criteria | Measurement Method |
|----------|----------|-------------------|
| 성능 | 생성자·Handler 리팩터로 인한 성능 저하 없음 | SXSSF 스트리밍 경로 벤치마크 (선택) |
| Javadoc | 신규 public API 전체에 예시 코드 포함 | `./gradlew javadoc` 경고 0 |
| 빌드 | `./gradlew test` + `./gradlew compileJava` (example 포함) 통과 | Gradle |

---

## 4. Success Criteria

### 4.1 Definition of Done

- [ ] T1, T2, T4, T5, T7, T8, T9 모두 완료 (T3/T6 은 범위 축소로 삭제됨)
- [ ] `./gradlew test` 통과 (전부 새 API 로 재작성)
- [ ] `./gradlew compileJava` (example 포함) 통과
- [ ] `./gradlew javadoc` 경고 0
- [ ] CHANGELOG / README / `META-INF/AI.md` / `docs/llms.txt` 업데이트
- [ ] Gap analysis Match Rate ≥ 90%

### 4.2 Quality Criteria

- [ ] 신규 API 각각 최소 1개 unit test
- [ ] javadoc 예시 코드 포함 (새 API)
- [ ] example 앱 모든 showcase 엔드포인트 수동 동작 확인
- [ ] CHANGELOG 에 Before/After 마이그레이션 표 포함

---

## 5. Risks and Mitigation

| Risk | Impact | Likelihood | Mitigation |
|------|--------|------------|------------|
| T4 Reader `column` 리네이밍 — example 앱 동반 수정 누락 시 컴파일 에러 | Medium | Medium | T4 와 T8 을 **같은 커밋** 에서 수행. Definition of Done 에 example 빌드 포함 |
| Map Reader 를 범위에서 뺐지만 Writer 쪽만 `forMap` 이 있어 일관성 일시 결여 | Low | High | CHANGELOG 에 "v0.12.0 에서 Map Reader 도 forMap 으로 통일 예정" 명시 |
| 대량 변경 중 회귀 발생 | Medium | Medium | Task 순서를 독립성 순으로 배치 (T2 → T1 → T5 → T4), 각 Task 완료 시점 `./gradlew test` 반드시 통과 |

---

## 6. Architecture Considerations

### 6.1 Project Level Selection

excel-kit은 **Java 라이브러리**이므로 bkit의 web-app level 분류(Starter/Dynamic/Enterprise)는 직접 적용되지 않음. 대신 라이브러리 관점의 원칙을 따름:

- **공개 API 표면 최소화**: 필요한 것만 `public`, 나머지는 package-private
- **불변성 우선**: config 객체는 record 또는 builder로 생성 후 불변
- **Fluent API 일관성**: 모든 엔트리 포인트가 동일 패턴 (`builder() → chain → build()`)

### 6.2 Key Architectural Decisions

| Decision | Options | Selected | Rationale |
|----------|---------|----------|-----------|
| Config 객체 형태 | record / 전통 class / builder 패턴 / 없음 (Builder 내부 필드) | **Builder 내부 필드** | 수용 필드 3개뿐 — 별도 타입 분리 불필요 |
| Handler 공통 타입 | abstract class / interface / sealed interface | **sealed interface** | `final` 구현체 + 다형성 확보 |
| 변경 전략 | Deprecation 경유 / 즉시 삭제 | **즉시 삭제** | 외부 사용자 없음. CHANGELOG 에 Before/After 표만 남김 |
| Reader 네이밍 | 2-tier 유지 / column 단일화 | **column / columnAt 단일화** | v0.10.0 Writer 통합과 일치. Builder 접근은 `Consumer<Builder>` 람다로 |
| Map Writer 처리 | 클래스 유지 + shortcut 확장 / 정적 팩토리 흡수 / 클래스 삭제 | **클래스 삭제 + 정적 팩토리** | `Writer.forMap()` 반환 타입이 `Writer<Map<...>>` 이므로 Writer 의 모든 메서드 자동 상속 |
| 빌드 타겟 | Java 17 / 21 | **Java 17** (확정) | `kit/build.gradle.kts:9-11` 확인 — sealed/record 전부 사용 가능 |

### 6.3 Package Layout

변경 없음. 기존 구조 유지:

```
kit/src/main/java/io/github/dornol/excelkit/
├── excel/     (ExcelWriter, ExcelReader, Handler, Config 등)
├── csv/       (CsvWriter, CsvReader, Handler 등)
└── shared/    (FileHandler sealed interface — 공통 타입 위치)
```

`FileHandler`는 `shared/` 패키지에 배치 (Excel·Csv 양쪽에서 참조).

---

## 7. Convention Prerequisites

### 7.1 Existing Project Conventions

- [x] `CLAUDE.md` 릴리스 체크리스트 존재 — 이 플랜은 해당 체크리스트를 준수해야 함
- [x] JUnit 5 + `@Nested` 테스트 구조
- [x] `ColumnStyleConfig` 상속 기반 컬럼 설정 (현 상태 유지)
- [x] `ExcelColumn` 생성자 변경 시 테스트 파일 직접 생성자 호출도 수정 (CLAUDE.md 명시)

### 7.2 Conventions to Define/Verify

| Category | Current State | To Define | Priority |
|----------|---------------|-----------|:--------:|
| Builder 패턴 명명 | 미정 | `builder()` 정적 팩토리 + `build()` 종결. `Builder<T>` 정적 중첩 클래스 | High |
| Sealed interface | 미정 | `permits` 명시 + 구현체 `final` + javadoc 에 "single-user library" 컨텍스트 명시 | Medium |
| Builder 필드 네이밍 | 미정 | Apache POI SXSSF 용어와 일치 (`rowAccessWindowSize` 등) | Medium |

### 7.3 Build/Toolchain Verification (사전 필요)

| 항목 | 확인 방법 | 결과 |
|------|----------|------|
| Java source/target 버전 | `kit/build.gradle.kts` 조회 | **Design 단계에서 확인** |
| sealed interface 사용 가능 여부 | Java 17+ 여부 | Design 단계에서 확정 |

---

## 8. Next Steps

1. [ ] `/pdca design api-cleanup-v011` — Design 문서 작성 (클래스 다이어그램, 마이그레이션 표, Java 버전 확정)
2. [ ] Design 승인 후 `/pdca do api-cleanup-v011` — 구현 시작 (T1 → T2 → T4 → T5 → T3 → T7 → T8 → T9 순)
3. [ ] 구현 후 `/pdca analyze api-cleanup-v011` — Gap analysis
4. [ ] Match Rate ≥ 90% 시 `/pdca report api-cleanup-v011` → v0.11.0 릴리스 체크리스트 수행

---

## Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 0.1 | 2026-04-11 | Initial draft — v0.11.0 breaking change 범위 정의 | DongHyeok Kim |
