# 플로엠 발주 정렬기 (Order-Sorter)

쿠팡 발주 엑셀을 업로드하면 자동으로 F→L→V→FL 순서로 정렬하고 개입수·합계를 계산해주는 웹앱.

## 배포 방법

### 1. GitHub 업로드
- 이 폴더 전체를 GitHub `Order-sorter` 레포에 업로드
- `public/master.xlsx` 에 **플로엠리스.xlsx** 파일을 넣기 (이름을 `master.xlsx`로 변경)

### 2. Vercel 연결
- vercel.com → New Project → GitHub `Order-sorter` 연결
- 자동 빌드 & 배포 완료

### 3. 마스터 품목 업데이트
- `public/master.xlsx` 파일만 교체해서 GitHub에 push
- Vercel 자동 재배포

## 사용 방법
1. URL 접속
2. 마스터 자동 로드 클릭 (또는 파일 직접 선택)
3. 발주 엑셀 업로드
4. 정렬 결과 확인 후 엑셀 다운로드

## 정렬 순서
1. F시리즈 — 숫자 오름차순
2. L시리즈 — 숫자 오름차순
3. V시리즈 — 숫자 오름차순
4. FL시리즈 — 숫자 오름차순
5. 기타
