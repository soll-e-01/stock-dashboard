# KRX 일일 엑셀 자동생성

## 목적
아래 3개 파일을 매일 같은 형식으로 생성합니다.
- `YYMMDD_수급_DATA.xlsx`
- `YYMMDD_시세_DATA.xlsx`
- `YYMMDD_신고가.xlsx`

템플릿은 현재 폴더의 기존 파일(`260213_*.xlsx`)을 사용합니다.

## 파일
- `daily_krx_automation.py`: 메인 실행 스크립트
- `config.krx.example.json`: 설정 예시

## 준비
1. `config.krx.example.json`을 복사해 `config.krx.json` 생성
2. `datasets`의 `request_params`(권장) 또는 `otp_params`를 실제 KRX 값으로 채우기
3. 필요 시 `field_map` 컬럼명 보정

## 실행
```bat
.\.venv\Scripts\python.exe daily_krx_automation.py --config config.krx.json --date 20260213
```

`--date`를 생략하면 오늘 날짜(`YYYYMMDD`)를 사용합니다.

## 필수 입력 정보
아래 값이 있어야 완전 자동화가 됩니다.
- `datasets.supply.*.request_params` (권장): 각 투자주체 조회용 `bld` 및 파라미터
- `datasets.highs.*.request_params` (권장): 역사적/52주/60일 신고가 조회용 `bld` 및 파라미터

`price_all`은 `bld=MDCSTAT01501` 예시가 반영되어 있습니다.

## 참고
- 현재 코드의 KRX 수집은 `OTP 발급 -> CSV 다운로드` 방식입니다.
- `getJsonData.cmd` 직접 조회(`request_params`)도 지원합니다.
- 실행 환경에서 외부 네트워크 접근이 막혀 있으면 수집 단계에서 실패합니다.
- 엑셀 서식은 템플릿 기반으로 유지되며, 데이터 값만 교체합니다.
