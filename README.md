# Smart Inventory Backend

온드라이브 엑셀 데이터를 기반으로 스페어파츠 재고를 관리하는 Express 백엔드 API 서버입니다.

## 기능 

- 📊 OneDrive 엑셀 파일 연동 및 자동 캐싱
- 🔍 부품 검색 및 필터링
- 🤖 Gemini AI 기반 재고 상담
- ⚠️ 재고 부족 알림
- 📈 재고 통계 및 요약

## API 엔드포인트

```
GET  /api/inventory              - 전체 재고 데이터
GET  /api/inventory/categories   - 부품종류별 그룹화
GET  /api/inventory/category/:name - 특정 종류 상세 정보
GET  /api/inventory/summary      - 전체 통계
POST /api/ai/chat                - AI 채팅
GET  /health                     - 서버 상태 확인
```

## 설치 및 실행

### 로컬 개발

```bash
npm install
npm start
```

서버는 `http://localhost:3001` 에서 실행됩니다.

### 환경 설정

`.env` 파일 생성 (필요시):
```
PORT=3001
NODE_ENV=development
```

## 배포

### Render에 배포

1. GitHub 리포지토리 연결
2. Render 대시보드에서 새 Web Service 생성
3. 배포 설정:
   - **Build Command**: `npm install`
   - **Start Command**: `npm start`
   - **Environment**: Node.js

## 기술 스택

- **Runtime**: Node.js 20.x
- **Framework**: Express 4.18.2
- **AI**: Google Generative AI (Gemini)
- **데이터**: Excel via OneDrive
- **문서 처리**: XLSX

## 라이센스

ISC
