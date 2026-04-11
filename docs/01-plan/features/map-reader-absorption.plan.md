---
template: plan
version: 1.2
feature: map-reader-absorption
project: excel-kit
project-version: 0.11.0 → 0.12.0
author: DongHyeok Kim
date: 2026-04-12
status: Draft
---

# map-reader-absorption Planning Document

> **Summary**: v0.11.0 에서 미룬 Map **Reader** 흡수를 완결. `ExcelMapReader` / `CsvMapReader` 를 삭제하고 `ExcelReader.forMap()` / `CsvReader.forMap()` 정적 팩토리로 이관하여 Writer 쪽 (`forMap()`) 과 완전한 대칭 확보.
>
> **Project**: excel-kit
> **Version**: 0.11.0 → 0.12.0
> **Author**: DongHyeok Kim
> **Date**: 2026-04-12
> **Status**: Draft

---

## Executive Summary

| Perspective | Content |
|-------------|---------|
| **Problem** | v0.11.0 에서 Map **Writer** 는 `ExcelWriter.forMap()` / `CsvWriter.forMap()` 정적 팩토리로 흡수했지만, Map **Reader** 는 SAX 콜백 상태 머신 재작성 리스크로 별도 클래스 (`ExcelMapReader`, `CsvMapReader`) 로 남아 있음. Writer 쪽은 `forMap`, Reader 쪽은 여전히 `new XxxMapReader()` — API 가 비대칭. |
| **Solution** | `ExcelReader.forMap()` / `CsvReader.forMap()` 정적 팩토리를 도입하여 `ExcelReader<Map<String, String>>` / `CsvReader<Map<String, String>>` 를 반환. 내부에 "map mode" 플래그를 두고, `build()` 시점에 기존 setter/mapping 모드와 분기하여 `ExcelMapReader.MapSheetHandler` 로직을 이관. `ExcelMapReader.java` / `CsvMapReader.java` 파일 삭제. |
| **Function/UX Effect** | Map I/O 가 Writer/Reader 양쪽 모두 `Writer.forMap()` / `Reader.forMap()` 대칭 패턴으로 통일. 사용자는 "Map 모드는 Reader 클래스의 factory" 라는 단일 멘탈 모델로 기억. `ExcelReader<Map<String, String>>` 타입으로 나오므로 기존 `ExcelReader` fluent API (sheetIndex, headerRowIndex, onProgress 등) 를 그대로 사용 가능. |
| **Core Value** | v0.11.0 의 v0.12.0 연기 약속 이행. excel-kit Map API 의 완전한 대칭 달성. v1.0.0 안정화 전 남은 잔여 정리 작업 중 하나 해소. |

---

## 1. Overview

### 1.1 Purpose

v0.11.0 에서 연기한 Map Reader 흡수를 완결하여 excel-kit Map API 의 Writer/Reader 비대칭 해소.

### 1.2 Background

- v0.11.0 에서 `ExcelMapWriter` / `CsvMapWriter` 클래스 2개는 삭제하고 `Writer.forMap()` 정적 팩토리로 흡수함
- Map Reader 는 회귀 리스크 (SAX `XSSFSheetXMLHandler` 상태 머신 재작성 필요) 로 범위 제외
- v0.11.0 CHANGELOG 및 Design 문서에 "v0.12.0 에서 Map Reader 도 `forMap` 으로 통일 예정" 명시
- 현재 코드 상황:
  - `ExcelMapReader.java` — 265 lines, inner `ExcelMapReadHandler` + `MapSheetHandler` 포함
  - `CsvMapReader.java` — 297 lines, inner `CsvMapReadHandler`
  - `ExcelReader` 는 이미 2-mode 구조 (setter / mapping) — `build()` 에서 `rowMapper != null` 로 분기
  - `CsvReader` 도 동일 구조

### 1.3 Related Documents

- Previous work: `docs/01-plan/features/api-cleanup-v011.plan.md` (v0.11.0 Plan)
- Previous work: `docs/02-design/features/api-cleanup-v011.design.md` §3.5 — Map Reader 범위 제외 결정
- v0.11.0 CHANGELOG `[0.11.0]` Deferred to v0.12.0 섹션
- Key files:
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelReader.java`
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelReadHandler.java`
  - `kit/src/main/java/io/github/dornol/excelkit/excel/ExcelMapReader.java` (삭제 대상)
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvReader.java`
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvReadHandler.java`
  - `kit/src/main/java/io/github/dornol/excelkit/csv/CsvMapReader.java` (삭제 대상)

---

## 2. Scope

### 2.1 In Scope

- [ ] **[M1] ExcelReader.forMap() 정적 팩토리 추가**: `public static ExcelReader<Map<String, String>> forMap()` — 기본 `sheetIndex=0`, `headerRowIndex=0`. 내부 `mapMode=true` 플래그 세팅.
- [ ] **[M2] CsvReader.forMap() 정적 팩토리 추가**: `public static CsvReader<Map<String, String>> forMap()` — 기본 `headerRowIndex=0`, `delimiter=','`, `charset=UTF-8`. 내부 `mapMode=true` 플래그 세팅.
- [ ] **[M3] ExcelReadHandler 에 map mode 경로 추가**: 기존 setter/mapping 2-mode 를 3-mode 로 확장. Map mode 생성자 오버로드 + `MapSheetHandler` 이관. 기존 `ExcelMapReader.MapSheetHandler` 로직을 그대로 이식 (header 자동 감지 + `Map<String, String>` 빌드).
- [ ] **[M4] CsvReadHandler 에 map mode 경로 추가**: 동일 패턴. `CsvMapReadHandler` 의 header 자동 감지 + OpenCSV 기반 `Map<String, String>` 빌드 로직 이관.
- [ ] **[M5] ExcelMapReader / CsvMapReader 파일 삭제**: 2 파일 완전 제거.
- [ ] **[M6] 테스트 마이그레이션**: 기존 `MapReaderStreamTest`, `MapWriterReaderTest` 등에서 `new ExcelMapReader()` / `new CsvMapReader()` 호출부를 `ExcelReader.forMap()` / `CsvReader.forMap()` 로 이전. 기존 테스트 케이스는 기능적 동등성을 검증하는 데 그대로 사용.
- [ ] **[M7] Map mode 충돌 방지**: `mapMode == true` 인 Reader 에서 `.column(...)` / `.columnAt(...)` / `.skipColumn(...)` 호출 시 `IllegalStateException` 발생 ("column registration is not allowed on map-mode reader; use forMap() + build() only"). Setter 모드와 혼동 방지.
- [ ] **[M8] 신규 테스트 강화**: `ExcelReaderMapModeTest` / `CsvReaderMapModeTest` 신설 — map mode + column() 혼용 금지, `forMap()` 반환 타입 검증, 기존 `ExcelMapReader` 시절 동작과의 동등성 (header 자동 감지, `readAsStream()`, `onProgress()`, sheetIndex 등)
- [ ] **[M9] 문서 업데이트**: CHANGELOG 에 `[0.12.0]` 항목 + Migration Guide (`new ExcelMapReader()` → `ExcelReader.forMap()`). README, `META-INF/AI.md`, `docs/llms.txt` 업데이트.
- [ ] **[M10] example 앱 마이그레이션**: `example/` 내 `ReadShowcaseController`, `CsvShowcaseController` 에서 `new ExcelMapReader()` / `new CsvMapReader()` 호출부 수정 (T4 처럼 **같은 커밋** 에서).

### 2.2 Out of Scope

- **ColumnStyleConfig / ExcelColumnBuilder / ColumnConfig 통합** — 별도 feature 로 분리 (v0.12.0 에 같이 넣을지, v0.13.0 으로 미룰지는 이 작업 완료 후 결정).
- 새로운 기능 추가 (새 Reader 기능, 새 cell type 등).
- `readAsStream()` 구현 방식 변경 — 기존 producer-thread + `BlockingQueue` 패턴 그대로 이관, 리팩터링 없음.
- Apache POI 버전 업그레이드.
- `ExcelReader<T>` 의 generic type parameter T 를 `Map<String, String>` 으로 고정하는 타입 시스템 장치 — 단순히 `forMap()` 의 반환 타입으로만 명시 (사용자가 `ExcelReader<T>` 변수에 담아 쓰면 T 가 자유로움. 강제 장치는 과설계).
- CsvWriter/ExcelWriter 쪽 변경 — 이미 v0.11.0 에서 완료.

---

## 3. Requirements

### 3.1 Functional Requirements

| ID | Requirement | Priority | Status |
|----|-------------|----------|--------|
| FR-01 | `ExcelReader.forMap()` 가 `ExcelReader<Map<String, String>>` 를 반환하며 `mapMode=true` 플래그가 설정됨 | High | Pending |
| FR-02 | `CsvReader.forMap()` 가 `CsvReader<Map<String, String>>` 를 반환하며 `mapMode=true` 플래그가 설정됨 | High | Pending |
| FR-03 | `mapMode` 상태에서 `.column(...)` / `.columnAt(...)` / `.skipColumn(...)` / `.skipColumns(...)` 호출 시 `IllegalStateException` | High | Pending |
| FR-04 | `ExcelReader.forMap().build(in).read(...)` 가 기존 `new ExcelMapReader().build(in).read(...)` 과 **동일한 결과** (row 순서, 값, header 자동 감지, null 처리) | High | Pending |
| FR-05 | `CsvReader.forMap().build(in).read(...)` 가 기존 `new CsvMapReader().build(in).read(...)` 과 **동일한 결과** | High | Pending |
| FR-06 | `readAsStream()` 동작 유지 — map mode 에서도 producer-thread 기반 stream 제공, `try-with-resources` 로 정리 | High | Pending |
| FR-07 | `sheetIndex(int)`, `headerRowIndex(int)`, `onProgress(int, callback)`, `dialect(...)` (CSV), `delimiter(char)` (CSV), `charset(Charset)` (CSV) 등 기존 Reader fluent API 가 map mode 에서도 동작 | High | Pending |
| FR-08 | `ExcelMapReader.java` / `CsvMapReader.java` 파일 삭제 (package-private inner classes 포함) | High | Pending |
| FR-09 | 기존 테스트는 새 API 로 재작성 (기능 커버리지 유지) | High | Pending |
| FR-10 | CHANGELOG `[0.12.0]` 섹션 + Migration Guide (Before/After 표) | Medium | Pending |

### 3.2 Non-Functional Requirements

| Category | Criteria | Measurement Method |
|----------|----------|-------------------|
| 성능 | Map mode 읽기 성능이 기존 `ExcelMapReader` 대비 **동등 이상** (±5% 허용) | 선택적 벤치마크 (대용량 테스트 파일 1회 측정) |
| 메모리 | SAX streaming 유지 — 전체 파일을 메모리에 올리지 않음 | 기존 동작 검증, heap dump 없이 review 기반 |
| 빌드 | `./gradlew test` + `./gradlew compileJava` (example 포함) 통과 | Gradle |
| Javadoc | 신규 `forMap()` 메서드에 예시 코드 포함 | `./gradlew javadoc` 경고 0 |

---

## 4. Success Criteria

### 4.1 Definition of Done

- [ ] M1~M10 모두 완료
- [ ] `./gradlew test` 통과 (기존 + 신규 map mode 테스트)
- [ ] `./gradlew compileJava` (example 포함) 통과
- [ ] `./gradlew javadoc` 경고 0
- [ ] CHANGELOG `[0.12.0]` 섹션 추가, README / `META-INF/AI.md` / `docs/llms.txt` 업데이트
- [ ] Gap analysis Match Rate ≥ 90%
- [ ] `ExcelMapReader.java`, `CsvMapReader.java` 2 파일 실제 삭제 확인

### 4.2 Quality Criteria

- [ ] `ExcelReaderMapModeTest` / `CsvReaderMapModeTest` 각각 최소 8개 테스트 (factory 반환 타입, header 자동 감지, `readAsStream`, `onProgress`, mixed mode 금지 × 4)
- [ ] 기존 `MapWriterReaderTest` / `MapReaderStreamTest` 의 map reader 관련 케이스가 새 API 로 migrate 된 상태로 전부 통과
- [ ] example 앱의 Read 관련 엔드포인트 수동 동작 확인 (`/showcase/read/map` 등)

---

## 5. Risks and Mitigation

| Risk | Impact | Likelihood | Mitigation |
|------|--------|------------|------------|
| **SAX 콜백 상태 이관 중 회귀** — `MapSheetHandler` 의 `startRow`/`endRow`/`cell` 상태 머신이 기존 ExcelReadHandler 의 그것과 다름. 병합 시 row lifecycle 미묘한 차이로 회귀 발생 가능 | High | Medium | `MapSheetHandler` 를 **그대로 inner class 로 이식** (로직 rewrite 금지). `ExcelReadHandler` 내부 SAX 등록 시점에 `mapMode` 분기만 추가. 기존 `MapWriterReaderTest` 의 모든 map read 케이스를 migrate 하여 동등성 보증 |
| `ExcelReader<T>` 의 generic 과 Map mode 의 타입 안전성 충돌 — 사용자가 `ExcelReader<Map<String, String>>` 타입으로 받은 후 `.column(setter)` 를 부를 수 있음 | Medium | High | M7 (runtime IllegalStateException) 로 mixed mode 방지. Javadoc 에 "forMap() 이후 column() 호출 금지" 명시 |
| `readAsStream()` producer thread 재이식 시 리소스 누수 | Medium | Low | 기존 코드를 그대로 이식, `onClose(() -> { producer.interrupt(); close(); })` 유지. 기존 test 재사용 |
| example 앱 컴파일 에러 시 빌드 실패 | Medium | High | M10 을 M5 와 **같은 커밋** 에서 수행 (v0.11.0 T4+T8 패턴) |
| `CsvMapReader` 가 가진 `onProgress` 는 `CsvReader` 에도 이미 있음 — 동작 차이 확인 필요 | Low | Medium | Design 단계에서 양쪽 동작 비교, 차이 있으면 CsvReader 쪽 기준으로 통일 |

---

## 6. Architecture Considerations

### 6.1 Project Level Selection

excel-kit 은 Java 라이브러리 — bkit web-app level 분류 적용 안 됨. 라이브러리 관점 원칙:

- **기존 구조 보존**: `ExcelReader` / `CsvReader` / `ExcelReadHandler` / `CsvReadHandler` 의 public API 구조 유지
- **최소 침습**: `ExcelReadHandler` 에 constructor overload + private inner class 추가만 — 기존 setter/mapping 모드 건드리지 않음
- **Single-user breaking**: deprecation 없이 즉시 삭제 (v0.11.0 방침 계승)

### 6.2 Key Architectural Decisions

| Decision | Options | Selected | Rationale |
|----------|---------|----------|-----------|
| Map mode 플래그 위치 | Reader 클래스 / Handler 클래스 | **Reader 에 `boolean mapMode` 필드** | `forMap()` 이 Reader 상태를 만들고, `build()` 시점에 Handler 생성자로 전달. 기존 `rowMapper != null` 분기 패턴과 일관 |
| Handler 의 map mode 생성자 | 기존 setter 생성자 재사용 / 신규 생성자 오버로드 | **신규 생성자 오버로드** | 입력 파라미터가 다름 (columns 없음, supplier 없음, validator 없음). 오버로드로 의도 명확화 |
| `MapSheetHandler` 위치 | `ExcelReadHandler` 의 inner class / 독립 package-private 클래스 | **inner class** | 기존 `ExcelMapReader.MapSheetHandler` 구조 그대로 이관, 외부 노출 없음 |
| Mixed mode 방지 방식 | 타입 시스템 (별도 Reader 타입) / runtime 체크 | **runtime `IllegalStateException`** | 별도 타입 도입은 과설계. 단일 사용자 프로젝트에서 runtime 체크로 충분 |
| `readAsStream()` 구현 | 새로 작성 / 기존 코드 이식 | **기존 코드 이식** | 회귀 리스크 최소화. 검증된 producer-thread 패턴 유지 |
| 변경 전략 | Deprecation 경유 / 즉시 삭제 | **즉시 삭제** | v0.11.0 방침 계승 (외부 사용자 없음) |

### 6.3 Package Layout

변경 없음. 기존 구조 유지:

```
kit/src/main/java/io/github/dornol/excelkit/
├── excel/
│   ├── ExcelReader.java          (forMap() 추가, mapMode 필드 추가)
│   ├── ExcelReadHandler.java     (map mode 생성자 오버로드 + inner MapSheetHandler 이관)
│   └── ExcelMapReader.java       (← 삭제)
├── csv/
│   ├── CsvReader.java            (forMap() 추가, mapMode 필드 추가)
│   ├── CsvReadHandler.java       (map mode 생성자 오버로드 + inner 로직 이관)
│   └── CsvMapReader.java         (← 삭제)
└── shared/                        (변경 없음)
```

---

## 7. Convention Prerequisites

### 7.1 Existing Project Conventions

- [x] v0.11.0 의 "즉시 삭제" 방침 (feedback memory: `~/.claude/projects/.../memory/project_single_user.md`)
- [x] Java 17 타겟 (`kit/build.gradle.kts`)
- [x] CLAUDE.md 릴리스 체크리스트 — CHANGELOG / README / META-INF / example / test / build / tag / push
- [x] JUnit 5 + `@Nested` 테스트 구조
- [x] `ExcelColumn` 관련 테스트는 직접 생성자 호출 수정 (CLAUDE.md 명시)

### 7.2 Conventions to Define/Verify

| Category | Current State | To Define | Priority |
|----------|---------------|-----------|:--------:|
| Map mode runtime check 메시지 | 미정 | 명확한 에러 메시지 ("column registration is not allowed on map-mode reader; use forMap().build().read() only") | High |
| `ExcelReadHandler` 생성자 네이밍 | package-private | map mode 생성자도 package-private 유지 — 외부 호출 금지 | Medium |
| Map mode forMap() javadoc 예시 | 미정 | v0.11.0 Writer.forMap() javadoc 스타일 일관 | Medium |

### 7.3 Build/Toolchain Verification

| 항목 | 확인 | 결과 |
|------|------|------|
| Java 17 타겟 | `kit/build.gradle.kts:9-10` | ✅ 확인 완료 |
| Gradle toolchain | `build.gradle.kts:15-17` | ✅ JDK 21 (기존 유지) |

---

## 8. Next Steps

1. [ ] `/pdca design map-reader-absorption` — Design 문서 작성 (Handler 클래스 구조, MapSheetHandler 이관 diff, mixed mode 예외 메시지, 마이그레이션 1:1 표)
2. [ ] Design 승인 후 `/pdca do map-reader-absorption` — 구현 시작 (M3/M4 Handler 이관 → M1/M2 Reader factory → M7 runtime check → M5 파일 삭제 → M10 example → M6/M8 테스트 → M9 문서)
3. [ ] 구현 후 `/pdca analyze map-reader-absorption` — Gap analysis
4. [ ] Match Rate ≥ 90% 시 `/pdca report map-reader-absorption` → v0.12.0 릴리스 체크리스트 수행

---

## Version History

| Version | Date | Changes | Author |
|---------|------|---------|--------|
| 0.1 | 2026-04-12 | Initial draft — v0.11.0 에서 연기된 Map Reader 흡수 범위 정의. 10개 Task, 단일 사용자 방침 계승, SAX 콜백 이관이 핵심 리스크. | DongHyeok Kim |
