# QR 차량 2부제 점검 시스템

휴대폰으로 QR 접속 후 번호판 사진을 찍으면,
사전 등록된 차량 리스트와 대조해서 2부제 위반 대상 여부를 판단하고,
최종 위반 등록 후 엑셀 보고서를 내보내는 웹앱입니다.

## 포함 기능

- 지사별 모바일 접속 페이지 `/branch/{지사코드}`
- 차량 기준정보 CSV/XLSX 업로드
- 번호판 사진 업로드
- OCR 시도 후 수동보정 입력 지원
- 등록 차량 대조
- 2부제 위반 판정
- 위반차량 등록
- 점검이력 조회
- 엑셀 보고서 생성

## 기본 가정

- 2부제는 **날짜의 홀짝과 차량번호 끝자리 홀짝이 같으면 운행 제한**이라고 가정했습니다.
- 예외 차량은 위반에서 제외합니다.
- 실제 운영 전에는 기관 규정에 맞게 판정 로직을 바꾸면 됩니다.

## 차량 기준정보 파일 형식

다음 헤더 중 하나를 사용하면 됩니다.

- `branch_code, plate_no, owner_name, department, is_target, exempt, note`
- 또는 한글 헤더: `지사코드, 차량번호, 성명, 부서, 2부제대상, 예외, 비고`

예시:

```csv
branch_code,plate_no,owner_name,department,is_target,exempt,note
ICN,12가3456,홍길동,총무팀,Y,N,상시점검
SWN,123다4567,박민수,행정지원부,Y,Y,공무수행 예외
```

## 실행 방법

```bash
cd qr_plate_app
python -m venv .venv
source .venv/bin/activate   # Windows는 .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --reload
```

브라우저에서 `http://127.0.0.1:8000` 접속.

## OCR 관련

- 코드에는 `pytesseract` 기반 OCR 연결부가 들어 있습니다.
- 서버에 Tesseract 엔진이 설치되어 있어야 실제 OCR이 동작합니다.
- OCR이 불안정할 수 있으므로 **수동보정 입력칸**을 함께 두었습니다.

## 실제 운영 전에 보완할 것

- 실제 기관 엑셀 서식 매핑
- 사용자 로그인
- 지사별 권한 분리
- HTTPS 배포
- 사진 원본 보관 정책
- 번호판 OCR 정확도 개선
- 중복 위반 등록 방지 로직
- QR 코드 PNG 자동 생성


## 배포용 실행 파일 추가
이 폴더에는 서버 배포용 `Dockerfile` 과 `render.yaml` 이 포함되어 있습니다.

### Render 배포
1. 이 폴더를 GitHub 저장소로 올립니다.
2. Render에서 **Blueprint** 또는 **New Web Service** 로 저장소를 연결합니다.
3. `render.yaml` 을 사용하면 `/app/data` 경로에 1GB 디스크가 마운트되어 SQLite 파일이 유지됩니다.
4. 배포 후 `/admin/vehicles` 에서 차량 기준정보를 업로드하고, 지사별 URL(`/branch/ICN` 등)을 QR로 배포하면 됩니다.

### Docker 직접 실행
```bash
docker build -t qr-plate-app .
docker run -p 8000:8000 -v ./data:/app/data qr-plate-app
```
