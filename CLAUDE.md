# excel-kit

Fluent API 기반 Excel/CSV 읽기·쓰기 Java 라이브러리 (Apache POI SXSSF 스트리밍).

## 프로젝트 구조

- `kit/` — 라이브러리 본체 (Maven Central 배포 대상)
- `example/` — Spring Boot 예제 앱 (showcase 엔드포인트)
- `docs/` — 웹 배포용 문서 (`llms.txt`)

## 릴리스 체크리스트

버전을 올릴 때 아래 항목을 모두 수행할 것:

1. GitHub PR 확인 — 미처리 PR (dependabot 등) 머지 또는 정리
2. `build.gradle.kts` — `version` 변경
3. `CHANGELOG.md` — `[x.y.z] - YYYY-MM-DD` 섹션 추가
4. `README.md` 최신화:
   - Installation 섹션의 Maven/Gradle 버전 업데이트
   - 상단 Features 목록에 새 기능 추가
   - 새 기능 사용법 섹션 추가
   - 설정 메서드 목록 업데이트 (ExcelColumnBuilder 메서드 나열 부분)
5. `example/` 최신화:
   - `WriteShowcaseController` — 새 기능 showcase 엔드포인트 추가
   - `index.html` — 해당 다운로드 버튼 추가
6. `META-INF/AI.md` 및 `META-INF/excel-kit/*.md` — 새 기능이 있으면 문서 반영
7. `docs/llms.txt` — 새 기능이 있으면 문서 반영
8. 빌드 확인:
   - `./gradlew clean test` — 단위 테스트 전체 통과
   - `./gradlew :kit:javadoc` — javadoc 경고 0 확인
9. example 앱 실행 확인:
   - Docker Compose 기동: `cd example && docker compose up -d` (MariaDB)
   - 앱 실행: `./gradlew :example:bootRun` — Spring 컨텍스트 정상 초기화 확인
   - 주요 엔드포인트 수동 검증: `http://localhost:8080` 에서 showcase 페이지 접근
   - JPQL/HQL 쿼리는 컴파일 시 안 잡히므로 **반드시 앱 기동으로 확인**
   - Docker 정리: `docker compose down`
10. 커밋 → 태그(`vx.y.z`) → 푸시(`git push origin main --tags`)

## 코드 컨벤션

- 테스트: JUnit 5 + `@Nested` 클래스로 기능별 그룹핑
- 컬럼 설정: `ColumnStyleConfig`를 상속하는 `ExcelColumnBuilder`(체이닝)와 `ColumnConfig`(람다) 두 API 유지
- `ExcelColumn` 생성자 변경 시 테스트 파일의 `ExcelColumn.of()` 호출도 함께 확인
- 패키지 구조: `core/` (공통 타입), `excel/` (Excel), `csv/` (CSV)
- Builder 패턴: `ExcelWriter.builder()`, `ExcelWorkbook.builder()` — 생성자 대신 사용
- Reader 진입점: `setter()`, `mapping()`, `forMap()` 정적 팩토리
- 외부 사용자 없음: breaking change 시 deprecation 없이 즉시 삭제 허용
