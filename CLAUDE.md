# excel-kit

Fluent API 기반 Excel/CSV 읽기·쓰기 Java 라이브러리 (Apache POI SXSSF 스트리밍).

## 프로젝트 구조

- `kit/` — 라이브러리 본체 (Maven Central 배포 대상)
- `example/` — Spring Boot 예제 앱 (showcase 엔드포인트)
- `docs/` — 웹 배포용 문서 (`llms.txt`)

## 릴리스 체크리스트

버전을 올릴 때 아래 항목을 모두 수행할 것:

1. `build.gradle.kts` — `version` 변경
2. `CHANGELOG.md` — `[x.y.z] - YYYY-MM-DD` 섹션 추가
3. `README.md` — Installation 섹션의 Maven/Gradle 버전 업데이트
4. `META-INF/AI.md` 및 `META-INF/excel-kit/*.md` — 새 기능이 있으면 문서 반영
5. `docs/llms.txt` — 새 기능이 있으면 문서 반영
6. 테스트 통과 확인: `./gradlew test`
7. 커밋 → 태그(`vx.y.z`) → 푸시(`git push origin main --tags`)

## 코드 컨벤션

- 테스트: JUnit 5 + `@Nested` 클래스로 기능별 그룹핑
- 컬럼 설정: `ColumnStyleConfig`를 상속하는 `ExcelColumnBuilder`(체이닝)와 `ColumnConfig`(람다) 두 API 유지
- `ExcelColumn` 생성자 변경 시 테스트 파일의 직접 생성자 호출도 함께 수정
