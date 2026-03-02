# excel-kit 코드 분석 리포트

> 분석 일자: 2026-03-02
> 대상 버전: 0.2.1
> 분석 범위: `kit/src/main/java/` 전체 소스 코드

---

## 목차

1. [보안 위험](#1-보안-위험)
2. [잠재적 버그](#2-잠재적-버그)
3. [개선 사항](#3-개선-사항)
4. [요약](#4-요약)

---

## 1. 보안 위험

### 1.1 [HIGH] Zip Bomb 보호 기본값이 지나치게 높음

**파일:** `ExcelReader.java:29-30`

```java
private static final int DEFAULT_MAX_FILE_COUNT = 10_000_000;
private static final int DEFAULT_MAX_BYTE_ARRAY_SIZE = 2_000_000_000;
```

**문제:**
- 기본 zip 엔트리 수 제한이 1,000만 개, 최대 바이트 배열 크기가 ~2GB로 설정되어 있어 Zip Bomb 공격에 취약할 수 있다.
- 이 설정은 **JVM 전역 설정**이므로, 한 번 변경하면 동일 프로세스 내 모든 POI 작업에 영향을 미친다.
- 사용자가 `configureLargeFileSupport()`를 명시적으로 호출해야만 적용되므로, 호출하지 않으면 POI 기본 제한이 적용된다. 하지만 호출 시 너무 큰 값이 설정된다.

**권장 조치:**
- 기본값을 보수적으로 조정 (예: `maxFileCount = 1_000_000`, `maxByteArraySize = 500_000_000`)
- JavaDoc에 보안 관련 주의사항 명시

---

### 1.2 [MEDIUM] Windows 환경에서 임시 파일 권한 미설정

**파일:** `TempResourceCreator.java:43-45`

```java
} else {
    // Windows
    return Files.createTempDirectory(UUID.randomUUID().toString());
}
```

**문제:**
- POSIX 시스템에서는 `rwx------` 권한을 명시적으로 설정하지만, Windows 환경에서는 별도 권한 설정 없이 기본값으로 생성된다.
- 다중 사용자 Windows 서버 환경에서 다른 사용자가 임시 파일에 접근할 수 있다.

**권장 조치:**
- Windows에서도 ACL을 통해 현재 사용자만 접근 가능하도록 설정하거나, JavaDoc에 Windows 환경에서의 보안 주의사항을 명시

---

### 1.3 [MEDIUM] 비밀번호가 String으로 처리됨

**파일:** `ExcelHandler.java:72`

```java
public void consumeOutputStreamWithPassword(@NonNull OutputStream outputStream, @NonNull String password)
```

**문제:**
- 비밀번호가 `String`으로 전달되어 GC 전까지 JVM 힙 메모리에 남는다. `String`은 불변이므로 명시적으로 지울 수 없다.
- 메모리 덤프나 리플렉션을 통해 비밀번호가 노출될 수 있다.

**권장 조치:**
- `char[]`를 받는 오버로드 메서드 추가를 고려하고, 사용 후 배열을 0으로 채워서 지우기
- 현재 API의 JavaDoc에 비밀번호 메모리 잔류에 대한 주의사항 명시

---

### 1.4 [LOW] CSV Injection 가능성

**파일:** `CsvWriter.java:179-188`

```java
private static String escapeCsv(Object input) {
    if (input == null) return "";
    String value = input.toString();
    if (value.contains(",") || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
        return "\"" + value.replace("\"", "\"\"") + "\"";
    }
    return value;
}
```

**문제:**
- `=`, `+`, `-`, `@`, `\t`, `\r` 로 시작하는 셀 값은 Excel에서 수식으로 해석될 수 있다 (CSV Injection / Formula Injection).
- 사용자가 입력한 데이터를 그대로 CSV로 내보내는 경우, 악의적인 수식이 삽입될 수 있다.

**권장 조치:**
- 위험 문자로 시작하는 값에 대해 앞에 작은따옴표(`'`) 또는 탭 문자를 삽입하는 방어 로직 추가를 고려
- 최소한 JavaDoc에 CSV Injection 위험 고지

---

## 2. 잠재적 버그

### 2.1 [HIGH] getColumnIndex()에서 소문자 셀 참조 미처리 및 오버플로 위험

**파일:** `ExcelReadHandler.java:252-264`

```java
private int getColumnIndex(String cellReference) {
    StringBuilder sb = new StringBuilder();
    for (char c : cellReference.toCharArray()) {
        if (Character.isLetter(c)) sb.append(c);
        else break;
    }
    String col = sb.toString();
    int colIdx = 0;
    for (char c : col.toCharArray()) {
        colIdx = colIdx * 26 + (c - 'A' + 1);  // 소문자 미처리, 오버플로 위험
    }
    return colIdx - 1;
}
```

**문제:**
1. `Character.isLetter(c)`는 소문자도 통과시키지만, `c - 'A' + 1` 연산은 대문자만 올바르게 처리한다. 소문자 `a`는 `'a' - 'A' + 1 = 33`이 되어 잘못된 인덱스를 반환한다.
2. 비정상적으로 긴 셀 참조 문자열이 들어오면 `colIdx * 26`에서 **int 오버플로**가 발생할 수 있다.
3. `cellReference`가 null이면 `NullPointerException` 발생.

**권장 조치:**
```java
private int getColumnIndex(String cellReference) {
    int colIdx = 0;
    for (char c : cellReference.toCharArray()) {
        if (!Character.isLetter(c)) break;
        colIdx = colIdx * 26 + (Character.toUpperCase(c) - 'A' + 1);
    }
    return colIdx - 1;
}
```

---

### 2.2 [HIGH] CellData.resetDateFormats()의 비원자적 연산

**파일:** `CellData.java:122-134`

```java
public static void resetDateFormats() {
    DATE_FORMAT_PATTERNS.clear();        // (1) 비어있는 순간이 존재
    DATE_FORMAT_PATTERNS.addAll(DEFAULT_DATE_FORMATS);  // (2) 이후 추가
}
```

**문제:**
- `CopyOnWriteArrayList`의 `clear()`와 `addAll()`은 각각 원자적이지만, 두 연산 사이에 다른 스레드가 `DATE_FORMAT_PATTERNS`를 읽으면 **빈 리스트**를 보게 된다.
- 같은 문제가 `resetDateTimeFormats()`에도 존재한다.

**권장 조치:**
- `CopyOnWriteArrayList`를 직접 조작하는 대신, 리스트를 새로 생성하여 `AtomicReference`로 교체하거나, `synchronized` 블록으로 보호

---

### 2.3 [MEDIUM] ExcelColumn.applyFunction()의 과도한 예외 흡수

**파일:** `ExcelColumn.java:52-58`

```java
Object applyFunction(T rowData, Cursor cursor) {
    try {
        return function.apply(rowData, cursor);
    } catch (Exception e) {
        log.error("applyFunction exception caught : {}, {} \n", rowData, cursor, e);
        return null;
    }
}
```

**문제:**
- 모든 `Exception`을 잡아서 `null`을 반환하므로, 개발 시 데이터 추출 함수의 타입 불일치나 로직 오류를 발견하기 어렵다.
- `OutOfMemoryError`를 제외한 대부분의 예외가 무시되며, 셀에 빈 값이 채워진다.
- 같은 패턴이 `CsvColumn.applyFunction()` (CsvColumn.java:48-54)에도 존재한다.

**권장 조치:**
- 최소한 `ClassCastException` 등 특정 예외만 잡거나, 디버그 모드에서는 예외를 전파하는 옵션 제공을 고려

---

### 2.4 [MEDIUM] ExcelColumn.setColumnData()의 과도한 예외 흡수

**파일:** `ExcelColumn.java:82-93`

```java
void setColumnData(SXSSFCell cell, Object columnData) {
    if (columnData == null) {
        cell.setCellValue("");
        return;
    }
    try {
        this.columnSetter.set(cell, columnData);
    } catch (Exception e) {
        log.warn("cast error: {}", e.getMessage());
        cell.setCellValue(String.valueOf(columnData));
    }
}
```

**문제:**
- `ExcelDataType` enum의 setter들은 `(Long) value`, `(Integer) value` 등 강제 캐스팅을 수행한다 (ExcelDataType.java:35-40).
- 잘못된 타입이 전달되면 `ClassCastException`이 발생하지만, 여기서 잡혀서 `String.valueOf()`로 대체된다.
- 이로 인해 Excel 셀에 숫자 대신 문자열이 들어가며, 사용자가 알아채기 어렵다.

**권장 조치:**
- 경고 로그의 수준을 높이거나, 최소한 `log.warn`에 스택 트레이스 포함 (`e.getMessage()`만이 아닌 `e` 자체를 전달)

---

### 2.5 [MEDIUM] ExcelHandler의 consumed 필드에 대한 스레드 안전성 미보장

**파일:** `ExcelHandler.java:30`

```java
private boolean consumed = false;
```

**문제:**
- `consumed` 필드가 `volatile`이 아니며, 동기화도 없다.
- 두 스레드가 동시에 `consumeOutputStream()`을 호출하면, 둘 다 `consumed == false`를 읽고 동시에 `wb.write()`를 실행할 수 있다.

**권장 조치:**
- `consumed`를 `volatile`로 선언하거나, `AtomicBoolean`으로 교체
- 또는 JavaDoc에 "이 클래스는 thread-safe하지 않음"을 명시

---

### 2.6 [LOW] CsvReadHandler에서 BOM이 단독 문자인 경우

**파일:** `CsvReadHandler.java:82-84`

```java
if (line.length > 0 && line[0] != null && line[0].startsWith("\uFEFF")) {
    line[0] = line[0].substring(1);
}
```

**문제:**
- `line[0]`이 BOM 문자 하나(`"\uFEFF"`)만으로 구성되어 있으면 `substring(1)`은 빈 문자열 `""`을 반환한다.
- 이 자체는 에러는 아니지만, 헤더 이름이 빈 문자열이 되어 이후 컬럼 매핑에서 혼동을 줄 수 있다.

---

### 2.7 [LOW] ExcelReadHandler의 sheetIndex 상한 미검증

**파일:** `ExcelReadHandler.java:88-90`

```java
if (sheetIndex < 0) {
    throw new IllegalArgumentException("sheetIndex must be non-negative");
}
```

**문제:**
- 하한만 검증하고 상한은 검증하지 않는다. 매우 큰 `sheetIndex`가 전달되면 파일의 모든 시트를 순회한 후에야 예외가 발생한다 (line 130-132에서 처리됨).
- 실질적인 문제는 크지 않으나, 조기 검증이 가능하다면 추가하는 것이 좋다.

---

## 3. 개선 사항

### 3.1 SXSSFWorkbook 버퍼 크기를 설정 가능하게 변경

**파일:** `ExcelWriter.java:57`

```java
this.wb = new SXSSFWorkbook(1000);
```

**현재 상태:**
- SXSSFWorkbook의 메모리 내 행 버퍼가 1,000으로 하드코딩되어 있다.

**제안:**
- 생성자 파라미터로 버퍼 크기를 받을 수 있도록 확장. 메모리가 제한적인 환경에서는 100~500으로 줄이고, 성능이 중요한 환경에서는 유지하거나 늘릴 수 있도록 한다.

---

### 3.2 ExcelWriter에 AutoCloseable 사용 패턴 개선

**파일:** `ExcelWriter.java:197-219`

**현재 상태:**
- `ExcelWriter`가 `AutoCloseable`을 구현하지만, `write()` 메서드에서 `ExcelHandler`를 반환한 후에도 `wb`에 대한 참조를 유지한다.
- `ExcelHandler.consumeOutputStream()`이 호출되면 `wb.close()`가 실행되지만, `write()` 후 `ExcelHandler`를 사용하지 않고 `ExcelWriter`만 close하는 것도 가능하다.

**제안:**
- `ExcelWriter.write()` 이후 `ExcelWriter.close()`가 호출되면 이미 전달된 workbook과 이중 close가 발생할 수 있으므로, 상태 관리를 명확히 하거나 JavaDoc에 사용 패턴 가이드 추가

---

### 3.3 CsvWriter에도 빈 columns 검증 추가

**파일:** `CsvWriter.java:109`

```java
public CsvHandler write(@NonNull Stream<T> stream) {
    // columns가 비어있으면 헤더만 있는 빈 CSV가 생성됨
```

**현재 상태:**
- `ExcelWriter.write()`는 `columns.isEmpty()` 시 예외를 던지지만 (ExcelWriter.java:198-199), `CsvWriter.write()`는 검증 없이 빈 CSV를 생성한다.

**제안:**
```java
if (this.columns.isEmpty()) {
    throw new CsvWriteException("columns setting required");
}
```

---

### 3.4 CellData의 정적 설정에 대한 스레드 안전성 강화

**파일:** `CellData.java:36, 77-78`

```java
private static volatile Locale defaultLocale = Locale.KOREA;
private static final List<DateTimeFormatter> DATE_FORMAT_PATTERNS = new CopyOnWriteArrayList<>(...);
```

**현재 상태:**
- `defaultLocale`은 `volatile`이지만, `setDefaultLocale()`과 `asNumber()` 사이에 복합 연산의 원자성은 보장되지 않는다.
- `resetDateFormats()`의 `clear()` + `addAll()` 사이에 빈 리스트가 노출될 수 있다.

**제안:**
- `defaultLocale`을 `AtomicReference<Locale>`로 변경
- `resetDateFormats()`를 원자적으로 처리:
```java
public static void resetDateFormats() {
    // CopyOnWriteArrayList에서는 replaceAll 대신 새 리스트로 교체
    synchronized (DATE_FORMAT_PATTERNS) {
        DATE_FORMAT_PATTERNS.clear();
        DATE_FORMAT_PATTERNS.addAll(DEFAULT_DATE_FORMATS);
    }
```
또는 `AtomicReference<List<DateTimeFormatter>>`를 사용하여 스냅샷 교체 방식으로 변경

---

### 3.5 AbstractReadHandler.mapColumn()에서 예외 정보 보강

**파일:** `AbstractReadHandler.java:102-112`

```java
protected boolean mapColumn(BiConsumer<T, CellData> setter, T instance, CellData cellData,
                            int columnIndex, List<String> headerNames, List<String> messages) {
    try {
        setter.accept(instance, cellData);
        return true;
    } catch (Exception e) {
        String header = (columnIndex < headerNames.size()) ? headerNames.get(columnIndex) : "column#" + columnIndex;
        messages.add("Failed to set column: " + header);
        log.warn("Column mapping failed", e);
        return false;
    }
}
```

**제안:**
- 에러 메시지에 실패한 셀의 값(`cellData.formattedValue()`)과 예외 메시지를 포함하면, 사용자가 어떤 데이터에서 문제가 발생했는지 파악하기 쉬워진다:
```java
messages.add("Failed to set column '" + header + "': value='" + cellData.formattedValue() + "', reason=" + e.getMessage());
```

---

### 3.6 ExcelWriter.close()에서 예외 정보 로깅

**파일:** `ExcelWriter.java:339-345`

```java
@Override
public void close() {
    try {
        wb.close();
    } catch (Exception e) {
        // already closed or disposed — safe to ignore
    }
}
```

**제안:**
- 예외를 완전히 무시하기보다, `log.debug` 수준으로 기록하면 문제 발생 시 디버깅에 도움이 된다.

---

### 3.7 ExcelDataType에서 null-safe 캐스팅

**파일:** `ExcelDataType.java:35-85`

```java
LONG((cell, value) -> cell.setCellValue((Long) value), ...),
INTEGER((cell, value) -> cell.setCellValue((Integer) value), ...),
```

**현재 상태:**
- `null`이 전달되면 `NullPointerException`이 발생하고, `ExcelColumn.setColumnData()`에서 잡힌다.
- 하지만 `null` 체크는 `setColumnData()`에서 먼저 수행되므로 실제로는 도달하지 않는다.

**제안:**
- 현재 구조에서는 문제가 발생하지 않지만, 방어적으로 각 setter에서도 null 체크를 추가하면 `ExcelDataType`을 단독으로 사용하는 경우에도 안전해진다.

---

## 4. 요약

| 등급 | 구분 | 항목 | 파일 |
|------|------|------|------|
| HIGH | 보안 | Zip Bomb 기본 제한값 과다 | `ExcelReader.java` |
| HIGH | 버그 | 셀 참조 파싱 시 소문자/오버플로 미처리 | `ExcelReadHandler.java` |
| HIGH | 버그 | resetDateFormats()의 비원자적 연산 | `CellData.java` |
| MEDIUM | 보안 | Windows 임시 파일 권한 미설정 | `TempResourceCreator.java` |
| MEDIUM | 보안 | 비밀번호 String 타입 처리 | `ExcelHandler.java` |
| MEDIUM | 버그 | applyFunction() 과도한 예외 흡수 | `ExcelColumn.java`, `CsvColumn.java` |
| MEDIUM | 버그 | setColumnData() 과도한 예외 흡수 | `ExcelColumn.java` |
| MEDIUM | 버그 | consumed 필드 스레드 안전성 미보장 | `ExcelHandler.java` |
| LOW | 보안 | CSV Injection 방어 없음 | `CsvWriter.java` |
| LOW | 버그 | BOM 단독 문자 시 빈 헤더 가능 | `CsvReadHandler.java` |
| LOW | 버그 | sheetIndex 상한 미검증 | `ExcelReadHandler.java` |
| - | 개선 | SXSSFWorkbook 버퍼 크기 설정 가능 | `ExcelWriter.java` |
| - | 개선 | CsvWriter 빈 columns 검증 추가 | `CsvWriter.java` |
| - | 개선 | CellData 정적 설정 스레드 안전성 강화 | `CellData.java` |
| - | 개선 | mapColumn() 에러 메시지 보강 | `AbstractReadHandler.java` |
| - | 개선 | ExcelWriter.close() 예외 로깅 | `ExcelWriter.java` |
| - | 개선 | ExcelDataType null-safe 캐스팅 | `ExcelDataType.java` |

---

### 우선순위 권장

1. **즉시 수정:** `getColumnIndex()` 소문자 처리 (2.1) - 실제 데이터 오류 유발 가능
2. **단기 수정:** CSV Injection 방어 (1.4), `resetDateFormats()` 원자성 (2.2), CsvWriter 빈 columns 검증 (3.3)
3. **중기 개선:** Zip Bomb 기본값 조정 (1.1), consumed 스레드 안전성 (2.5), 예외 처리 전략 재검토 (2.3, 2.4)
4. **장기 개선:** 나머지 개선 사항 및 문서화
