# p-net Order Reply Tool

p-net 시스템에서 다운로드한 엑셀 파일과 공장납기회신 엑셀 파일을 비교하여 p-net 업로드용 수동 업로드 파일을 생성하는 Python GUI 애플리케이션입니다.

## 기능

- **파일 비교**: p-net 다운로드 파일과 공장납기회신 파일을 PO# + PO-LINE# 기준으로 비교
- **자동 매칭**: 두 파일의 데이터를 조합하여 통합 정보 생성
- **결과 생성**: p-net 업로드 가능한 형식의 엑셀 파일 자동 생성
- **GUI 인터페이스**: 사용자 친화적인 Tkinter 기반 GUI

## 설치

### 필수 요구사항
- Python 3.8 이상
- pip (Python 패키지 관리자)

### 의존성 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

1. 애플리케이션 실행:
```bash
python main.py
```

2. 다음 순서로 파일을 선택합니다:
   - **p-net 다운로드 파일**: p-net 시스템에서 다운로드한 엑셀 파일
   - **공장납기회신 파일**: 공장에서 제공한 납기회신 엑셀 파일
   - **저장 경로**: 결과 파일을 저장할 위치

3. "파일 처리 및 생성" 버튼을 클릭하여 처리 시작

4. 처리 완료 후 지정한 위치에 결과 파일이 생성됩니다.

## 파일 형식

### 입력 파일 1: p-net 다운로드 파일
| 컬럼 | 설명 |
|------|------|
| E | CPO# |
| F | CPO-LINE# |
| G | LINE SEQ |
| H | CPO QTY |
| I | Material |
| O | EX-F |
| P | ETD |
| R | 납품확정여부 |
| X | PO# |
| Y | PO-LINE# |

### 입력 파일 2: 공장납기회신 파일
| 컬럼 | 설명 |
|------|------|
| A | PO# |
| B | LINE# |
| C | Material |
| D | CPO QTY |
| E | ETD |
| F | EX-F |
| G | 내부노트 |

### 출력 파일: p-net 업로드용 수동 업로드 파일
| 순서 | 컬럼 | 설명 |
|------|------|------|
| 1 | PO# | 주문번호 |
| 2 | PO-LINE# | 주문 라인 번호 |
| 3 | Material | 자재번호 |
| 4 | CPO QTY | 수량 |
| 5 | ETD | 예상 납기일 |
| 6 | EX-F | EX-F 코드 |
| 7 | 내부노트 | 내부 노트 |
| 8 | CPO# | CPO 번호 |
| 9 | CPO-LINE# | CPO 라인 번호 |
| 10 | LINE SEQ | 라인 시퀀스 |
| 11 | 납품확정여부 | p-net 원본 R열 |

## 코드 서명 안내

- 실행파일 사용자 PC에서는 인증서 파일(`.pfx`)이 필요하지 않습니다.
- 인증서는 개발/배포 시점에 실행파일(`.exe`)에 디지털 서명을 붙일 때만 필요합니다.
- 사용자는 서명된 `exe`만 받아서 실행하면 됩니다.

## 서명 릴리스 방법

다음 스크립트가 준비되어 있습니다.

- `sign_release.ps1`: build/dist exe 서명
- `verify_signature.ps1`: 서명 상태 검증
- `build_release_signed.ps1`: 빌드 + 서명 + 검증 + zip 생성

저장소 인증서를 이용한 예시:

```powershell
.\build_release_signed.ps1 -Version 1.0.12 -StoreThumbprint 3F87D815085767D06BCD496D0BFB7D34605AEB73 -StoreScope CurrentUser -EnsureLocalTrust
```

간편 실행용 배치 파일도 사용할 수 있습니다.

```bat
build_release_signed.bat 1.0.12
```

## 프로젝트 구조

```
OrderReply/
├── main.py                  # GUI 애플리케이션 메인
├── excel_processor.py       # 엑셀 파일 처리 로직
├── requirements.txt         # Python 의존성
└── README.md               # 프로젝트 설명서 (현재 파일)
```

## 주요 모듈

### excel_processor.py
엑셀 파일 읽기, 비교, 결과 생성을 담당하는 `ExcelProcessor` 클래스:
- `read_pnet_download()`: p-net 다운로드 파일 읽기
- `read_factory_reply()`: 공장납기회신 파일 읽기
- `compare_and_generate()`: 파일 비교 및 결과 생성
- `save_result()`: 결과 파일 저장

### main.py
GUI 애플리케이션 `OrderReplyApp` 클래스:
- 파일 선택 UI
- 처리 진행상황 표시
- 백그라운드 스레드에서 파일 처리 실행

## 사용 예시

1. p-net에서 주문 정보 엑셀 파일 다운로드
2. 공장에서 받은 납기회신 파일 준비
3. 애플리케이션 실행 후 두 파일 선택
4. 저장할 파일명과 위치 지정
5. "파일 처리 및 생성" 버튼 클릭
6. 생성된 파일을 p-net에 수동 업로드

## 에러 처리

- 파일을 찾을 수 없거나 읽을 수 없는 경우 오류 메시지 표시
- 필수 컬럼이 없는 경우 오류 메시지 표시
- 매칭되지 않는 항목에 대해서도 결과 파일에 포함 (관련 CPO 정보는 빈 칸)

## 향후 개선 사항

- [ ] 로깅 기능 추가
- [ ] 배치 처리 기능
- [ ] 데이터 유효성 검사 강화
- [ ] 매칭되지 않은 항목 분석 리포트
- [ ] 엑셀 형식 및 스타일 커스터마이징

## 라이선스

개인 프로젝트

## 문의

문제가 발생하거나 개선 사항이 있으면 로그 정보와 함께 연락주세요.
