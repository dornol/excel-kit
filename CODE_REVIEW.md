# excel-kit 코드 분석 리포트

> 분석 일자: 2026-03-02
> 대상 버전: 0.2.1
> 분석 범위: `kit/src/main/java/` 전체 소스 코드
> 최종 업데이트: 2026-03-02 (모든 항목 수정 완료)

---

## 목차

1. [보안 위험](#1-보안-위험)
2. [잠재적 버그](#2-잠재적-버그)
3. [개선 사항](#3-개선-사항)
4. [요약](#4-요약)

---

## 1. 보안 위험

### 1.1 ~~[HIGH] Zip Bomb 보호 기본값이 지나치게 높음~~ (해결)

**파일:** `ExcelReader.java:29-30`

**수정 내용:** 기본값을 보수적으로 조정 (`1_000_000` 엔트리, `500_000_000` 바이트).

---

### 1.2 ~~[MEDIUM] Windows 환경에서 임시 파일 권한 미설정~~ (해결)

**파일:** `TempResourceCreator.java`

**수정 내용:** Windows에서 `AclFileAttributeView`를 사용하여 현재 사용자만 접근 가능하도록 ACL 제한 추가.

---

### 1.3 ~~[MEDIUM] 비밀번호가 String으로만 처리됨~~ (해결)

**파일:** `ExcelHandler.java`

**수정 내용:** `char[]` 비밀번호를 받는 오버로드 메서드 추가. 사용 후 `Arrays.fill(password, '\0')`으로 제로 클리어.

---

### 1.4 ~~[LOW] CSV Injection 가능성~~ (해결)

**파일:** `CsvWriter.java`

**수정 내용:** `=`, `+`, `-`, `@`, `\t`, `\r`로 시작하는 셀 값에 작은따옴표(`'`) 접두사 삽입.

---

## 2. 잠재적 버그

### 2.1 ~~[HIGH] getColumnIndex()에서 소문자 셀 참조 미처리 및 오버플로 위험~~ (해결)

**파일:** `ExcelReadHandler.java`

**수정 내용:** `Character.toUpperCase(c)` 적용 + Excel 최대 컬럼(16,384) 초과 시 예외 발생으로 오버플로 방지.

---

### 2.2 ~~[HIGH] CellData.resetDateFormats()의 비원자적 연산~~ (해결)

**파일:** `CellData.java`

**수정 내용:** `clear()` + `addAll()` 대신 새 `CopyOnWriteArrayList` 인스턴스를 `volatile` 참조로 원자적 교체.

---

### 2.3 ~~[MEDIUM] ExcelColumn.applyFunction()의 과도한 예외 흡수~~ (해결)

**파일:** `ExcelColumn.java`, `CsvColumn.java`

**수정 내용:** `catch(Exception)` → `catch(RuntimeException)`으로 범위 축소. 로그 메시지에 컬럼명, 행 데이터, 스택 트레이스 포함.

---

### 2.4 ~~[MEDIUM] ExcelColumn.setColumnData()의 과도한 예외 흡수~~ (해결)

**파일:** `ExcelColumn.java`

**수정 내용:** `catch(Exception)` → `catch(RuntimeException)`. 로그에 컬럼명, 값, 전체 스택 트레이스 포함.

---

### 2.5 ~~[MEDIUM] ExcelHandler의 consumed 필드에 대한 스레드 안전성 미보장~~ (해결)

**파일:** `ExcelHandler.java`

**수정 내용:** `boolean consumed` → `AtomicBoolean consumed`로 교체. `compareAndSet(false, true)`로 원자적 상태 전이.

---

### 2.6 ~~[LOW] CsvReadHandler에서 BOM이 단독 문자인 경우~~ (해결)

**파일:** `CsvReadHandler.java`

**수정 내용:** BOM 제거 후 빈 문자열인 경우 `CsvReadException` 발생.

---

### 2.7 ~~[LOW] ExcelReadHandler의 sheetIndex 상한 미검증~~ (해결)

**파일:** `ExcelReadHandler.java`

**수정 내용:** `sheetIndex > 255` 검증 추가.

---

## 3. 개선 사항

### 3.1 ~~SXSSFWorkbook 버퍼 크기를 설정 가능하게 변경~~ (해결)

**파일:** `ExcelWriter.java`

**수정 내용:** `rowAccessWindowSize`를 받는 5-파라미터 생성자 추가. 기존 생성자는 기본값(1000) 유지.

---

### 3.2 ExcelWriter에 AutoCloseable 사용 패턴 문서화

**파일:** `ExcelWriter.java`

**현재 상태:** `close()` 시 이중 close 가능하나, 예외 없이 안전하게 처리됨. 문서화는 추후 README 업데이트 시 반영 예정.

---

### 3.3 ~~CsvWriter에도 빈 columns 검증 추가~~ (해결)

**파일:** `CsvWriter.java`

**수정 내용:** `write()` 호출 시 `columns.isEmpty()`이면 `CsvWriteException` 발생.

---

### 3.4 ~~CellData의 정적 설정에 대한 스레드 안전성 강화~~ (해결)

**파일:** `CellData.java`

**수정 내용:** `resetDateFormats()`/`resetDateTimeFormats()`를 `volatile` 참조 스왑으로 원자적 교체.

---

### 3.5 ~~AbstractReadHandler.mapColumn()에서 예외 정보 보강~~ (해결)

**파일:** `AbstractReadHandler.java`

**수정 내용:** 에러 메시지에 헤더명, 셀 값(`formattedValue()`), 예외 메시지 포함.

---

### 3.6 ~~ExcelWriter.close()에서 예외 정보 로깅~~ (해결)

**파일:** `ExcelWriter.java`

**수정 내용:** `catch` 블록에 `log.debug()` 추가.

---

### 3.7 ~~ExcelDataType에서 타입 호환성 개선~~ (해결)

**파일:** `ExcelDataType.java`

**수정 내용:** `(Long) value` → `((Number) value).longValue()` 등 `Number` 인터페이스 기반 캐스팅으로 변경하여 타입 호환성 향상.

---

## 4. 요약

| 등급 | 구분 | 항목 | 파일 | 상태 |
|------|------|------|------|------|
| HIGH | 보안 | Zip Bomb 기본 제한값 과다 | `ExcelReader.java` | **해결** |
| HIGH | 버그 | 셀 참조 파싱 시 소문자/오버플로 미처리 | `ExcelReadHandler.java` | **해결** |
| HIGH | 버그 | resetDateFormats()의 비원자적 연산 | `CellData.java` | **해결** |
| MEDIUM | 보안 | Windows 임시 파일 권한 미설정 | `TempResourceCreator.java` | **해결** |
| MEDIUM | 보안 | 비밀번호 String 타입 처리 | `ExcelHandler.java` | **해결** |
| MEDIUM | 버그 | applyFunction() 과도한 예외 흡수 | `ExcelColumn.java`, `CsvColumn.java` | **해결** |
| MEDIUM | 버그 | setColumnData() 과도한 예외 흡수 | `ExcelColumn.java` | **해결** |
| MEDIUM | 버그 | consumed 필드 스레드 안전성 미보장 | `ExcelHandler.java` | **해결** |
| LOW | 보안 | CSV Injection 방어 없음 | `CsvWriter.java` | **해결** |
| LOW | 버그 | BOM 단독 문자 시 빈 헤더 가능 | `CsvReadHandler.java` | **해결** |
| LOW | 버그 | sheetIndex 상한 미검증 | `ExcelReadHandler.java` | **해결** |
| - | 개선 | SXSSFWorkbook 버퍼 크기 설정 가능 | `ExcelWriter.java` | **해결** |
| - | 개선 | AutoCloseable 사용 패턴 문서화 | `ExcelWriter.java` | 보류 |
| - | 개선 | CsvWriter 빈 columns 검증 추가 | `CsvWriter.java` | **해결** |
| - | 개선 | CellData 정적 설정 스레드 안전성 강화 | `CellData.java` | **해결** |
| - | 개선 | mapColumn() 에러 메시지 보강 | `AbstractReadHandler.java` | **해결** |
| - | 개선 | ExcelWriter.close() 예외 로깅 | `ExcelWriter.java` | **해결** |
| - | 개선 | ExcelDataType 타입 호환성 개선 | `ExcelDataType.java` | **해결** |
