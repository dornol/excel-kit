# excel-kit

Fluent API 기반 Excel/CSV 읽기·쓰기 Java 라이브러리 (Apache POI SXSSF 스트리밍).

## 프로젝트 구조

- `kit/` — 라이브러리 본체 (Maven Central 배포 대상)
- `excel-kit-spring/` — Spring 연동 모듈 (업로드/다운로드 응답 헬퍼, Maven Central 배포 대상)
- `example/` — Spring Boot 예제 앱 (showcase 엔드포인트)
- `docs/guide/` — 문서 (기능별 분리된 가이드, `index.md`가 진입점)

## 릴리스 체크리스트

버전을 올릴 때 아래 항목을 모두 수행할 것:

1. GitHub PR 확인 — 미처리 PR (dependabot 등) 머지 또는 정리
2. `build.gradle.kts` — `version` 변경 (`kit`·`excel-kit-spring` 동일 버전, `verifyReleaseVersion` 대상)
3. API 호환성 기준선 갱신 — `.github/workflows/api-compatibility.yml` 의 baseline 기본값(2곳)과 `.japicmp-allowlist` 주석을 직전 릴리스 버전으로 변경
4. `CHANGELOG.md` — `[x.y.z] - YYYY-MM-DD` 섹션 추가
5. `README.md` 최신화:
   - Installation 섹션의 버전 번호 업데이트
   - Features at a Glance 테이블에 새 기능 반영
   - Quick Start 코드가 최신 API 사용하는지 확인
6. 문서 최신화 (새 기능이 있으면):
   - `docs/guide/*.md` — 해당 카테고리의 가이드 파일에 새 기능 사용법 추가
   - `META-INF/AI.md` — Quick Reference / Key API Notes 에 새 API 반영
   - `META-INF/excel-kit/*.md` — 해당 카테고리 파일 반영
     - 쓰기 스타일/컬럼 → `column-config.md`
     - 읽기 → `reading.md`
     - group/formula/protection/chart/validation → `advanced.md`
     - CSV 전용 → `csv.md`
     - 기본 예제 → `quick-start.md`
   - **검증**: 이번 버전 CHANGELOG `### Added` 항목의 API 이름이 모든 문서군(README, `docs/guide/`,
     `AI.md`, `META-INF/excel-kit/*.md`)에 최소 1회 이상 등장해야 함. 태그 푸시 시 release 워크플로가
     자동 실행하며, 푸시 전에 로컬에서 먼저 확인할 것:
     ```sh
     scripts/verify-release-docs.sh            # build.gradle.kts 의 version 기준
     ```
     API 가 아닌 이름(JDK 타입, enum 상수 등)이 오탐으로 잡히면 `scripts/release-docs-ignore.txt` 에 추가
7. `example/` 최신화 (showcase할 기능이 있으면):
   - `WriteShowcaseController` — 새 기능 엔드포인트 추가
   - `index.html` — 해당 다운로드 버튼 추가
8. 빌드 확인:
   - `./gradlew clean test` — 단위 테스트 전체 통과
   - `./gradlew :kit:javadoc` — javadoc 경고 0 확인
9. example 앱 실행 확인:
   - `cd example && docker compose up -d` → `./gradlew :example:bootRun`
   - `http://localhost:8080` 에서 showcase 페이지 접근
   - JPQL/HQL 쿼리는 컴파일 시 안 잡히므로 **반드시 앱 기동으로 확인**
   - `docker compose down`
10. 커밋 → 태그(`vx.y.z`) → 푸시(`git push origin main --tags`)

## 코드 컨벤션

- 테스트: JUnit 5 + `@Nested` 클래스로 기능별 그룹핑
- 테스트에서 `ExcelHandler`/`CsvHandler` 를 얻으면 반드시 소비(`writeTo(OutputStream.nullOutputStream())`)하고,
  `ExcelWorkbook`/`ExcelTemplateWriter` 는 try-with-resources 로 닫을 것 — 임시 파일이 남으면
  `TempFileLeakGuard`(`test-support/` 에 두고 kit·excel-kit-spring 테스트에 자동 적용)가 해당 테스트를 실패시킨다. POI `SXSSFWorkbook` 을 직접 만드는
  테스트는 `DisposingSxssfWorkbook` 을 사용
- 컬럼 설정: `ColumnStyleConfig`를 상속하는 `ExcelColumnBuilder`(체이닝)와 `ColumnConfig`(람다) 두 API 유지
- `ExcelColumn` 생성자는 `ColumnStyleConfig`를 통째로 받음 — 새 필드 추가 시 `ColumnStyleConfig`에 필드+setter 추가 후 `ExcelColumn` 내부에서 복사
- 패키지 구조: `core/` (공통 타입), `excel/` (Excel), `csv/` (CSV)
- 생성 진입점: `ExcelWriter.create()` / `ExcelWriter.create(opts -> ...)`, `ExcelWorkbook.create()` / `ExcelWorkbook.create(opts -> ...)` — `InitOptions` Consumer 로 초기 설정만 받고, 그 외 설정은 fluent 메서드
- Reader 진입점: `setter()`, `mapping()`, `forMap()` 정적 팩토리
- 외부 사용자 없음: breaking change 시 deprecation 없이 즉시 삭제 허용
