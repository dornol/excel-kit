# Excel 검증 지시서 (Windows + Excel 에이전트용)

이 문서는 **Excel 이 설치된 Windows 에서 AI 에이전트(Claude Code 등)가 따라 하는 작업 지시서**다.
excel-kit 이 만든 `.xlsx` 파일을 실제 Excel 로 열어, CI 와 단위 테스트가 확인할 수 없는 것을 확인하고 보고한다.

## 목적

excel-kit 의 테스트는 파일을 POI 로 읽거나 XML 로 검사한다. 그러나 **Excel 이 파일을 복구 메시지 없이 여는지,
필터가 데이터 전체에 걸리는지, 차트가 어디에 놓이는지**는 Excel 만 판정할 수 있다. 이번 작업이 그 판정이다.

특히 아래 변경이 Excel 에서 문제없는지가 핵심이다(아직 Excel 로 확인한 적이 없다).

- `autoFilter(true)` 가 헤더 행만이 아니라 **헤더 + 데이터 행**(예: `A1:B8`)을 필터링하도록 바뀌었다.
- `table(...)` 과 `autoFilter(true)` 를 함께 쓰면 시트 수준 필터를 쓰지 않고 **표 XML 안에 `<autoFilter>`** 를 넣는다.
- 롤오버된 시트마다 차트를 그리며, 차트는 그 시트의 요약 행·`afterData` 내용 **아래**에 놓인다.
- 스트리밍 워크북의 ZIP 모드가 `AlwaysWithCompatibility` 로 바뀌었다(모든 파일의 zip 구조에 영향).
- 암호화(AES-256) 두 경로, 템플릿에 값을 채운 결과, 모든 기능을 한꺼번에 켠 파일.
- 업로드 양식 기능(18~26번): 드롭다운 **화살표**(`showDropDown` 을 더 이상 쓰지 않는다 — OOXML 에서 `true` 는 화살표를
  숨긴다), 부모 값에 따라 바뀌는 목록(숫자 부모 포함), 검증 안내 메시지·빈 값, 숨김·완전 숨김 시트와 열리는 시트, 확대 비율,
  페이지 나누기 위치, 같은 값 병합, 줄무늬와 헤더 필수 표시(`@" *"` 서식), 셀 주석이 달린 오류 리포트, 그리고 표 헤더의
  필수 표시, 시트가 롤오버될 때 병합·페이지 나누기·의존 드롭다운(숨김 헬퍼 시트 공유)·필터·확대의 조합(25, 26번), 그룹 소계(`groupSubtotal`)가 `SUBTOTAL` 수식으로 계산되고 총계가 소계를 다시 더하지 않는지(27번), 롤오버·병합·줄무늬·페이지 나누기·필터와의 조합(28번), 보호된 시트에서 소계 셀이 잠겨 있는지(29번), 필터 아래의 총계·조건부 서식·유효성 검사(30번), 막대 위에 선을 겹치고 보조 축을 쓰는 콤보 차트(31번).

## 하지 말 것

- **저장소의 코드·문서를 수정하지 않는다.** 라이브러리 결함처럼 보여도 고치지 말고 증거와 함께 보고만 한다.
- 샘플 파일이나 보고서를 외부 서비스에 올리지 않는다.
- 작업이 끝나면 Excel 프로세스(`EXCEL.EXE`)를 반드시 종료한다. 멈춘 Excel 이 남아 있으면 다음 열기가 막힌다.

## 1. 준비

샘플 파일(`01-…xlsx` ~ `31-…xlsx`, `manifest.json`, `CHECKLIST.md`)을 구한다. 둘 중 하나:

**A. CI 아티팩트 (JDK 불필요, 권장)**

```powershell
gh run list --repo dornol/excel-kit --workflow CI --branch main --limit 3
gh run download <run-id> --repo dornol/excel-kit -n excel-samples -D $env:TEMP\excel-samples
```

**B. 직접 생성 (JDK 21 이상 필요)**

```powershell
git pull
./gradlew :kit:excelSamples        # kit/build/excel-samples 에 만들어진다
```

인터넷에서 받은 파일은 Mark-of-the-Web 때문에 보호된 보기로 열릴 수 있다. 먼저 푼다.

```powershell
Get-ChildItem $env:TEMP\excel-samples -Recurse | Unblock-File
```

`manifest.json` 의 각 항목은 `file`, `title`, `look`(눈으로 볼 것), `checks`(프로그램으로 확인할 것)로 이루어진다.
**각 파일에 대해 `look` 과 `checks` 를 모두 수행한다.**

## 2. 방법

Excel 을 COM 으로 구동한다(PowerShell `New-Object -ComObject Excel.Application` 또는 Python `pywin32`).
Excel 이 보는 것을 직접 읽어 판정한다. 화면은 필요한 경우에만(아래 "시각 확인") 캡처한다.

### 2-1. 복구 메시지 없이 열리는가 (모든 파일, 가장 중요)

"We found a problem with some content … Do you want us to try to recover as much as we can?" 같은 대화상자가
뜨면 **FAIL** 이다. COM 으로 열면 이 대화상자가 호출을 막거나(대기), 예외가 나거나, `DisplayAlerts=$false` 면
조용히 복구하고 열 수도 있다. 그래서 다음 순서로 확인한다.

1. 파일마다 **별도 Excel 인스턴스**를 `Visible=$true`, `DisplayAlerts=$true` 로 띄우고, `Workbooks.Open` 을
   **백그라운드 작업(예: `Start-Job`/스레드)에서 20초 제한**으로 호출한다.
2. 제한 시간 안에 반환하면 열린 것이다. **반환하지 않으면** Excel 창들을 조사한다(`Get-Process EXCEL`,
   UI Automation 또는 `pywinauto` 로 최상위 대화상자 제목·본문 텍스트 수집). 대화상자가 있으면 **본문 전체와
   스크린샷을 증거로 저장**하고 FAIL 로 기록한 뒤 해당 Excel 을 강제 종료한다.
3. 예외가 났으면 메시지를 그대로 기록하고 FAIL.
4. 열렸더라도 `DisplayAlerts=$false` 로 자동 복구되었을 가능성이 있으므로, 열린 뒤 **창 제목에 `[Repaired]`·
   `복구됨` 같은 표시**가 있는지, 그리고 아래 속성 확인이 manifest 와 일치하는지도 함께 본다.

암호화 파일(07, 08)은 비밀번호를 `Workbooks.Open` 의 `Password` 인자로 준다. **틀린 비밀번호로 열면 반드시 실패**해야
하며(예외), 맞는 비밀번호로는 열려야 한다. 비밀번호 입력 대화상자가 뜨는 것은 정상이다(비밀번호 없이 열었을 때).

### 2-2. checks 항목을 Excel 객체 모델에 대응시키기

`checks` 의 각 문자열은 `종류 key=value …` 형태다. 아래 표대로 확인한다(셀 주소는 A1 형식, 시트 이름에 공백이 있으면
따옴표로 감싸져 있다).

| 종류 | Excel 에서 확인하는 방법 |
|---|---|
| `autoFilter sheet range` | 시트의 `AutoFilterMode` 가 True 이고 `AutoFilter.Range.Address($false,$false)` 가 `range` 와 같다. `range` 가 없으면(`none`) `AutoFilterMode` 가 False. |
| `filterRows … criteria expectVisibleDataRows` | 필터 범위의 `column` 번째 열에 `Range.AutoFilter(column, criteria)` 를 적용한 뒤, 데이터 행(헤더 제외, 요약 행 제외) 중 `SpecialCells(xlCellTypeVisible=12)` 로 보이는 행 수를 센다. 요약 행이 보여야 하는지 숨겨져야 하는지는 **manifest 항목의 문구를 따른다**(Excel 이 필터를 바로 아래 행까지 넓히므로 01 번은 숨겨지는 것이 기대값이고, 17 번은 보이는 것이 기대값이다). 확인 후 `ShowAllData` 로 되돌린다. |
| `table sheet name range showAutoFilter [column1=이름]` | `ListObjects(name)` 의 `Range.Address($false,$false)` 와 `ShowAutoFilter`(True/False). `column1` 이 있으면 `ListColumns(1).Name` 도 그 이름과 같아야 한다(헤더 셀에 서식이 있어도 표의 열 이름은 순수한 이름). |
| `cell … value=` / `cell … text=` | 먼저 `Application.Calculate()`. `Value2` 가 값과 같은가. `text=` 는 `.Text`(화면에 보이는 문자열). |
| `cellText … text=` | `.Text` 가 정확히 같은가(표시 형식 확인). |
| `chartSeries sheet index chartType axisGroup` | `ChartObjects(1).Chart.SeriesCollection(index)` 의 `ChartType`(xlColumnClustered=51, xlLine=4, xlLineMarkers=65)과 `AxisGroup`(주축=1, 보조 축=2). |
| `chartAxis …` / `chartAxisTitle …` | `Chart.HasAxis(xlValue=2, xlSecondary=2)` 등과 `Chart.Axes(2, 1).AxisTitle.Text`(주), `Chart.Axes(2, 2).AxisTitle.Text`(보조). |
| `chart sheet count topLeftRowGreaterThan` | `ChartObjects().Count` 와 각 차트의 `TopLeftCell.Row` 가 값보다 큰가(요약 행 아래에 놓였는가). |
| `shapes sheet pictures` | `Shapes` 중 `Type = 13`(msoPicture) 개수. |
| `hyperlinks sheet count` | `Hyperlinks.Count`. |
| `protection …` | `ProtectContents`, 지정한 셀의 `Locked`, `$wb.ProtectStructure`. |
| `validation sheet cell type …` | `Range(cell).Validation.Type` (whole=1, list=3). 뒤에 붙은 속성도 확인한다: `inCellDropdown`(`Validation.InCellDropdown`; **True 여야 화살표가 보인다**), `ignoreBlank`(`IgnoreBlank`), `showInput`(`ShowInput`), `inputTitle`(`InputTitle`), `inputMessage`(`InputMessage`), `formula1`(`Formula1`; 범위 참조는 `=` 로 시작한다). |
| `dependentOptions sheet cell parentCell parent count` | 의존 드롭다운의 선택지 개수. 부모 셀(`parentCell`)에 `parent` 값을 넣고(숫자처럼 보이는 `1` 은 **숫자**로, `(a NUMBER)` 표시가 있으면 반드시 숫자), **`cell` 과 같은 행**의 빈 열(예: Z열)에 `=COUNTA(<cell 의 Validation.Formula1 에서 맨 앞 =를 뺀 것>)` 을 넣고 `Calculate` 한 값이 `count` 인지 본다. 수식이 `#N/A` 등 오류면 빈 목록(0)으로 본다. 끝나면 두 셀을 원래대로 되돌린다. |
| `sheetVisible sheet visible` | `Worksheets(name).Visible` (visible=-1, hidden=0, veryHidden=2). |
| `activeSheet name` | 파일을 연 직후 `ActiveSheet.Name`. |
| `zoom sheet percent` | 시트를 `Activate` 한 뒤 `ActiveWindow.Zoom`. |
| `pageBreaks sheet manualRows` | `HPageBreaks` 중 `Type = -4135`(xlPageBreakManual) 인 것의 `Location.Row`(새 페이지의 첫 행). **프린터 드라이버가 없으면 `HPageBreaks` 를 읽지 못한다** — "확인 불가"로 적고 페이지 나누기 미리보기를 눈으로 본다. |
| `fill sheet cells rgb` | 각 셀의 `DisplayFormat.Interior.Color` 를 RGB 로 바꿔 `rgb`(예: `F2F2F2`)와 같은지 본다(Excel 의 `Color` 는 `R + G*256 + B*65536`). `rgb=none` 은 `DisplayFormat.Interior.ColorIndex = -4142`(xlColorIndexNone). |
| `comment sheet cell text` | `Range(cell).Comment.Text()` 가 `text` 와 같은가. `none` 이면 `Comment` 가 없다. |
| `info …` | 합격·불합격 없이 **관찰값만 보고서에 적는다**(예: 병합으로 가려진 셀의 `Value2`). |
| `conditionalFormats sheet range countAtLeast` | `Range(range).FormatConditions.Count` 가 이상인가. |
| `name name refersTo` | `Names.Item(name).RefersTo` (`=Report!$A$2:$A$6` 형태). |
| `freezePanes sheet rows/columns` | 시트를 `Activate` 한 뒤 `ActiveWindow.FreezePanes`, `SplitRow`, `SplitColumn`. |
| `pageSetup sheet orientation` | `PageSetup.Orientation` (landscape = 2). |
| `headerVisible sheet range` | 범위의 각 셀에서 `DisplayFormat.Font.Color` 와 `DisplayFormat.Interior.Color`(화면에 실제로 적용되는 값)가 **서로 달라야** 한다. 같으면 글자가 배경에 묻혀 안 보이는 것이다(표 헤더가 흰 글자·흰 배경이었던 적이 있다). 값이 같지 않아도 확대 스크린샷으로 글자가 읽히는지 본다. 시트가 숫자면 N 번째 시트. |
| `merged sheet cell mergeArea` | `Range(cell).MergeArea.Address($false,$false)`. 병합되지 않은 셀은 자기 자신이 나온다(`mergeArea=B4`). |
| `sheets names=A|B|C` / `sheets count=N` | `Worksheets` 의 이름 순서 / 개수(숨김 시트도 포함된다). |

manifest 의 표현이 위 표에 없으면 의미에 맞게 해석하되, **해석을 보고서에 적는다**.

### 2-3. 시각 확인 (`look` 및 `visual:` 항목)

다음은 객체 모델로 증명하기 어려우므로 이미지를 만들어 직접 본다.

- 차트: `ChartObjects(i).Chart.Export("…\chart-<file>-<sheet>.png")`. 차트가 그 시트 데이터(첫 시트는 4행, 둘째 4행,
  셋째 2행)를 그리는지, 데이터 셀을 가리지 않는지 확인한다.
- 필터 단추·표 줄무늬·조건부 서식(빨간 배경, 데이터 막대)·서식 있는 텍스트(굵게/파랑 기울임)·사진: 해당 범위를
  `Range.CopyPicture` 로 클립보드에 복사해 이미지로 저장하거나 창 스크린샷을 찍는다.
- 저장한 PNG 는 직접 열어서 눈으로 판단하고, 근거 한 줄을 보고서에 적는다.

### 2-4. 이전 실행에서 알게 된 것

- 복구가 필요한 파일을 `Workbooks.Open` 으로 열면 COM 예외 `0x800A03EC` 가 나고, 셸로 열면 "…의 내용에 문제가 있습니다.
  이 통합 문서의 내용을 최대한 복구하시겠습니까?" 가 뜬다. 둘 다 FAIL 이다. Excel 이 직접 저장한 xlsx 를 같은 경로로 열어
  환경 문제가 아님을 먼저 확인하는 것이 좋다.
- 원인을 좁힐 때는 (a) 내용을 바꾸지 않고 표준 ZIP 으로 다시 묶어 열리는지(ZIP 구조 문제인지), (b) 부분(차트·표·병합)을
  빼면 열리는지 본다. 이 방법으로 ZIP 모드, 템플릿 차트의 빈 `<c:tx/>`, 표 헤더와 병합의 겹침을 찾아냈다.
  변형본은 작업 폴더에만 두고 저장소에는 두지 않는다.
- Excel 은 필터 범위 바로 아래에 이어진 데이터가 있으면 필터를 그 아래까지 넓혀 인식한다(`AutoFilter.Range` 가 파일의
  `ref` 보다 클 수 있다). 필터 범위는 **파일을 연 직후, 어떤 필터 호출도 하기 전에** 읽는다.

## 3. 결과 보고

작업 폴더(저장소 밖, 예: `$env:TEMP\excel-verification`)에 `excel-report.md` 를 만들고, **마지막 응답에도 같은 요약**을 쓴다.

```
# Excel 검증 결과
- Excel 버전: (Application.Version, Application.Build)
- 파일 수: N, PASS: n, FAIL: n, 확인 불가: n

| 파일 | 확인 항목 | 결과 | 근거 |
|---|---|---|---|
| 01-autofilter-plain.xlsx | opens clean | PASS | 20 초 안에 열림, 대화상자 없음 |
| 01-autofilter-plain.xlsx | autoFilter Report A1:B8 | PASS | AutoFilter.Range = A1:B8 |
| 04-table-with-autofilter.xlsx | opens clean | FAIL | 대화상자: "…removed feature: AutoFilter…" (스크린샷: …) |
```

- **FAIL 은 항목마다** 기대값, 실제값, 증거(대화상자 본문·스크린샷 경로·COM 오류 메시지)를 적는다.
- 확인하지 못한 것(Excel 이 해당 속성을 노출하지 않음 등)은 "확인 불가"로 구분하고 이유를 적는다.
- 마지막에 **한 줄 결론**을 쓴다: 지금 상태로 릴리스해도 되는가, 되지 않는다면 어떤 파일/기능 때문인가.
- 결과가 의심스러우면(예: 같은 구조의 파일 중 일부만 FAIL) 원인 후보를 좁히는 관찰을 덧붙인다
  (예: 표만 없는 변형은 열리는가). 단, 저장소는 수정하지 않는다.

## 4. 시간·실패 처리

- 파일 하나에 막혀 전체를 멈추지 말고, 열리지 않으면 FAIL 로 기록하고 다음 파일로 간다.
- Excel 이 응답하지 않으면 `EXCEL.EXE` 를 종료하고 새 인스턴스로 이어간다.
- 모든 파일을 마치면 열린 Excel 과 임시 인스턴스를 정리한다.
