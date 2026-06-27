# popup-inventory

팝업 스토어 재고 로케이션 매칭 웹사이트 (Express + SQLite + Vanilla JS)

## 주요 기능
- 📷 카메라 바코드 스캔 (듀얼 엔진): **ZXing(@zxing/browser) 우선 + html5-qrcode 폴백**. ZXing의 강력한 CODE128 1D 디코더 + TRY_HARDER 힌트로 작은 의류 라벨 인식률 대폭 향상. 고해상도(1080p→720p fallback) + 연속 자동초점 + 줌/플래시(토치) 자동 노출, EAN/UPC 무시 옵션, 2회 확인 모드, 후보 수동 확정 버튼. 다중 CDN 폴백 / HTTPS 환경 필수
  - SKU(CODE128/CODE39 등 영문+숫자) 우선 인식
  - EAN-13 / UPC 등 숫자 전용 소매 바코드는 기본 무시 (모달 토글로 해제 가능)
  - html5-qrcode 다중 CDN 폴백(unpkg → jsdelivr → cdnjs), HTTPS 환경 필수
  - html5-qrcode 라이브러리는 클릭 시점에 다중 CDN(jsdelivr → unpkg → cdnjs)으로 동적 로드 + 폴백
  - HTTPS 필수 (Render 도메인은 자동 HTTPS)
- 📦 상품등록: 바코드 + 로케이션 + 수량
- 🔍 재고조회: 상품명/바코드/로케이션 검색, 카드에 위치별 수량 + 인라인 매장이동
- 🔄 상품이동(창고-매장): 위치 간 이동
- 📊 현황: 위치별 합계, 전체 통계
- CSV 다운로드, QR/바코드 스캔 지원

## 실행 (로컬)

```bash
npm install
npm start
# http://localhost:3000
```

- `popup.db`가 같이 동봉되어 있어 기존 데이터(587 SKU / 1,543개 / 39 위치)로 바로 동작합니다.
- `popup.db`가 없으면 `popup_location.xlsx`로부터 자동 적재합니다.

## 환경변수

| 변수 | 기본값 | 설명 |
|---|---|---|
| PORT | 3000 | 서버 포트 |

## 프론트엔드 도메인 자동 감지

`public/index.html`의 `BASE_URL`은 `window.location.origin`을 사용합니다.
따라서 Render(`https://popup-ckgy.onrender.com`), 로컬(`http://localhost:3000`),
다른 호스팅 어디로 옮겨도 **별도 수정 없이 같은 오리진의 API를 자동 호출**합니다.
헤더의 "API:" 표시와 QR 코드도 현재 접속 URL을 자동으로 보여줍니다.

## Render Free 배포 가이드 (중요)

### 1. Render Dashboard → New → Web Service
- Repository: `https://github.com/tnals23000-ops/popup-inventory`
- Branch: `main`
- **Runtime: Node** (Docker 아님)
- **Build Command**: `npm install`
- **Start Command**: `npm start`
- **Instance Type**: Free

### 2. Node 버전 고정 (필수)
저장소 루트에 다음 파일이 모두 포함되어 있습니다.
- `package.json` → `"engines": { "node": "20.x" }`
- `.nvmrc` → `20.18.0`
- `.node-version` → `20.18.0`

Render는 위 파일을 자동 인식해 Node 20.18.0을 사용합니다.

### 3. better-sqlite3 빌드 이슈 해결
- **이전 실패 원인**: better-sqlite3 11.3.0 + Node 24.x 조합에서 V8 ABI 비호환(`v8-external.h:55:8`).
- **이번 패키지 해결**:
  - Node 버전을 **20.18.0 LTS**로 고정 (Render Free에서 가장 안정).
  - better-sqlite3를 **11.5.0**으로 고정 — Node 20용 prebuilt 바이너리(`@mapbox/node-pre-gyp`)가 정상 제공됨.
  - prebuilt가 받아지면 native compile 단계 자체를 건너뛰므로 빌드 실패 없음.

### 4. 빌드 시간 단축 팁
Render Free는 빌드 7분 제한. `package-lock.json`을 같이 커밋하면 더 빠르게 설치됩니다 (선택사항).

### 5. 영속성 주의
- Render Free는 디스크가 ephemeral(재시작 시 초기화).
- popup.db가 컨테이너 안에서 변경되어도 재배포 시 git 안의 popup.db로 리셋됩니다.
- 영구 저장이 필요하면 Render Disk(유료) 또는 Postgres로 마이그레이션 권장.

## API 요약

| Method | Path | 설명 |
|---|---|---|
| GET | /health | 헬스체크 |
| GET | /api/stats | 전체 통계 |
| GET | /api/product/:barcode | 바코드별 상세(위치별 수량 포함) |
| GET | /api/search/product?q=&limit= | 상품 검색 (stockByLocation 포함) |
| GET | /api/search/location/:loc | 위치별 검색 |
| GET | /api/locations | 위치 목록(합계/SKU수) |
| POST | /api/assign | 등록: {barcode, location, qty, ...} |
| POST | /api/move | 이동: {barcode, qty, fromLocation, toLocation, memo} |
| POST | /api/restock | 재고추가 |
| POST | /api/sale | 판매 |
| GET | /api/transactions | 거래 이력 |
| GET | /api/export | CSV 내보내기 |

## 트러블슈팅

### Q. Render 빌드가 또 실패한다
- Build Log에서 첫 번째 에러 라인 확인:
  - `node: command not found` → Build Command를 `npm install`만 두고 추가 명령 제거
  - `node-gyp ERR! ... v8-external.h` → Node 버전이 20이 아닐 가능성. Settings → Environment에서 `NODE_VERSION=20.18.0` 추가
  - `EACCES` → Build Command 앞에 `npm config set unsafe-perm true && ` 추가

### Q. 배포 후 첫 호출이 느림
Render Free는 15분 무활동 시 슬립. 첫 호출 시 30~60초 콜드 스타트가 정상.

## 데이터 구조

```
products              마스터 (바코드, 상품명, 색상, 사이즈, 초도재고, 현재재고)
stock_by_location     (바코드, 위치) → 수량
transactions          이력 (type, barcode, qty, location, before_qty, after_qty, memo, is_current)
```

- `type`: ASSIGN / MOVE / SALE / RESTOCK / CANCEL
- 위치 prefix 규칙: `창고-*` (창고), `매장-*` (매장), `미지정` (배치 전)
