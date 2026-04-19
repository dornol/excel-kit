# excel-kit

Fluent API 기반 Excel/CSV 읽기·쓰기 Java 라이브러리 (Apache POI SXSSF 스트리밍).

## 프로젝트 구조

- `kit/` — 라이브러리 본체 (Maven Central 배포 대상)
- `example/` — Spring Boot 예제 앱 (showcase 엔드포인트)
- `docs/guide/` — 문서 (기능별 분리된 가이드, `index.md`가 진입점)

## 릴리스 체크리스트

버전을 올릴 때 아래 항목을 모두 수행할 것:

1. GitHub PR 확인 — 미처리 PR (dependabot 등) 머지 또는 정리
2. `build.gradle.kts` — `version` 변경
3. `CHANGELOG.md` — `[x.y.z] - YYYY-MM-DD` 섹션 추가
4. `README.md` 최신화:
   - Installation 섹션의 버전 번호 업데이트
   - Features at a Glance 테이블에 새 기능 반영
   - Quick Start 코드가 최신 API 사용하는지 확인
5. 문서 최신화 (새 기능이 있으면):
   - `docs/guide/*.md` — 해당 카테고리의 가이드 파일에 새 기능 사용법 추가
   - `META-INF/AI.md` — Quick Reference / Key API Notes 에 새 API 반영
   - `META-INF/excel-kit/*.md` — 해당 카테고리 파일 반영
     - 쓰기 스타일/컬럼 → `column-config.md`
     - 읽기 → `reading.md`
     - group/formula/protection/chart/validation → `advanced.md`
     - CSV 전용 → `csv.md`
     - 기본 예제 → `quick-start.md`
   - **검증**: 이번 버전 CHANGELOG Added 항목의 API 이름(예: `headerRows`, `rowNumberColumn`)이
     아래 grep 에서 모든 문서군에 최소 1회 이상 hit 해야 함
     ```sh
     for api in <이번-버전-신규-API-이름들>; do
       echo "=== $api ==="
       grep -rl "$api" README.md docs/guide/ \
         kit/src/main/resources/META-INF/AI.md \
         kit/src/main/resources/META-INF/excel-kit/*.md || echo "MISSING: $api"
     done
     ```
6. `example/` 최신화 (showcase할 기능이 있으면):
   - `WriteShowcaseController` — 새 기능 엔드포인트 추가
   - `index.html` — 해당 다운로드 버튼 추가
7. 빌드 확인:
   - `./gradlew clean test` — 단위 테스트 전체 통과
   - `./gradlew :kit:javadoc` — javadoc 경고 0 확인
8. example 앱 실행 확인:
   - `cd example && docker compose up -d` → `./gradlew :example:bootRun`
   - `http://localhost:8080` 에서 showcase 페이지 접근
   - JPQL/HQL 쿼리는 컴파일 시 안 잡히므로 **반드시 앱 기동으로 확인**
   - `docker compose down`
9. 커밋 → 태그(`vx.y.z`) → 푸시(`git push origin main --tags`)

## 코드 컨벤션

- 테스트: JUnit 5 + `@Nested` 클래스로 기능별 그룹핑
- 컬럼 설정: `ColumnStyleConfig`를 상속하는 `ExcelColumnBuilder`(체이닝)와 `ColumnConfig`(람다) 두 API 유지
- `ExcelColumn` 생성자는 `ColumnStyleConfig`를 통째로 받음 — 새 필드 추가 시 `ColumnStyleConfig`에 필드+setter 추가 후 `ExcelColumn` 내부에서 복사
- 패키지 구조: `core/` (공통 타입), `excel/` (Excel), `csv/` (CSV)
- 생성 진입점: `ExcelWriter.create()` / `ExcelWriter.create(opts -> ...)`, `ExcelWorkbook.create()` / `ExcelWorkbook.create(opts -> ...)` — `InitOptions` Consumer 로 초기 설정만 받고, 그 외 설정은 fluent 메서드
- Reader 진입점: `setter()`, `mapping()`, `forMap()` 정적 팩토리
- 외부 사용자 없음: breaking change 시 deprecation 없이 즉시 삭제 허용
