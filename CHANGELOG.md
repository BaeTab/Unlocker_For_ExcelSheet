# 변경 이력 (Changelog)

이 프로젝트의 주요 변경 사항을 기록합니다.
형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/)를 따르며,
버전은 [유의적 버전](https://semver.org/lang/ko/)을 따릅니다.

## [2.0.1] - 2026-07-03

코드 리뷰(UI) 반영 — 버튼 클릭으로 앱이 죽을 수 있던 크래시 경로 제거.

### 수정 (Fixed)
- **`Process.Start` 예외로 인한 앱 크래시 방지**: "업데이트 확인"에서 다운로드 페이지를
  열 때, "결과 폴더 열기"에서 폴더를 열 때 브라우저/핸들러 부재·경로 소멸 등으로
  `Win32Exception` 이 나면 앱이 죽던 문제를 `try/catch` 로 흡수하고 안내 메시지로 대체.
- **업데이트 확인 핸들러**(`async void`)에 예외 처리 및 결과 null 방어 추가.
- **전역 예외 안전망 추가**(`Application.ThreadException` / `AppDomain.UnhandledException`).
- 새 배치 시작 시 이전 배치의 "결과 폴더 열기" 상태를 리셋.

## [2.0.0] - 2026-07-03

문서(`docs/고도화_아이디어.md`) 기반 전면 고도화. **.NET Framework 4.8 → .NET 9** 마이그레이션.

### 추가 (Added)
- **열기 암호(파일 전체 암호화) 해제**: CFB/OLE 컨테이너(ECMA-376 Agile/Standard
  암호화)로 암호화된 `.xlsx`를 암호 입력 후 복호화. (NPOI 기반)
- **통합문서 보호(`workbookProtection`) 제거** — 시트 추가/삭제/이동 잠금 해제.
- **쓰기 예약 암호(`fileSharing`) 제거** — 읽기 전용 권장/수정 암호 해제.
- **일괄 처리(batch)**: 여러 파일·폴더를 한 번에 큐에 넣어 순차 해제.
- **드래그 앤 드롭**: 파일/폴더를 창에 끌어다 놓아 추가.
- **비동기 처리 + 진행률 표시**: 큰 파일에서도 UI 멈춤 없음.
- **결과 로그 패널 + 파일별 상태(완료/보호없음/실패)**.
- **원본 보존**: 기본적으로 `원본명_unlocked.확장자`로 저장(원본 미변경).
  옵션으로 원본 덮어쓰기 선택 가능.
- **지원 확장자 확대**: `.xlsx` 외 `.xlsm`, `.xltx`, `.xltm` 추가. `.xls`는 감지 후 안내.
- **결과 폴더 열기** 버튼, Per-Monitor V2 고해상도(DPI) 대응.

### 변경 (Changed)
- 해제 엔진을 `Services/ExcelUnlocker` 클래스로 **UI와 분리**(테스트/재사용 가능).
- XML 편집 방식을 **정규식 → XML 파서(`XDocument`)**로 교체(정확·안전).
- 압축 처리를 **`Ionic.Zip`(DotNetZip) → .NET 내장 `System.IO.Compression`**으로 교체.
- 임시 파일을 원본 폴더가 아닌 **`%TEMP%`의 GUID 경로**에서 처리 후 정리.

### 수정 (Fixed)
- **보호 여부 판정 순서 버그**: 이전 버전은 보호 태그를 먼저 제거한 뒤 "남았는지"를
  검사해 **항상 "보호 없음"**으로 표시하고 "해제 완료" 안내가 뜨지 않던 문제를 수정.
- **원본 파괴 위험**: 원본을 그 자리에서 `.zip`으로 rename 하던 방식을 제거.
  이제 임시 복사본에서 작업하고 성공 시에만 결과를 저장하므로 실패해도 원본이 안전.
- 임시 폴더 삭제 시 예외가 다시 던져지던 문제(가드 추가).

### 제거 (Removed)
- 미사용 의존성: `EPPlus 4.1`(구버전 취약점), `Microsoft.Office.Interop.Excel`,
  `OfficeOpenXml.Core`, `OfficeOpenXml.Extends`, `Ionic.Zip`.

### 알려진 제한 (Known limitations / 향후)
- 레거시 `.xls`(BIFF8) 보호 해제, VBA 프로젝트 암호 해제는 아직 미지원(감지·안내만).
- 열기 암호를 **모르는** 경우의 복구는 지원하지 않음(정당한 용도 — 본인 소유/권한
  있는 파일에 한함).

## [1.0.0] - 2024

- 최초 버전: `.xlsx` 시트 보호(`sheetProtection`)만 정규식으로 제거.
