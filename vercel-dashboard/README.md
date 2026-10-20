# Vercel deployment package

이 폴더는 기존 대시보드의 배포용 복사본입니다. 새 앱은 만들지 않았습니다. 포함 자료는 교육용 가상 날씨 예제뿐이며, 매출 원본, 보고 양식, 행사 준비 원본은 포함하지 않습니다.

## 현재 폴더와 빌드 설정

Vercel에서 저장소를 Import할 때 Root Directory를 `vercel-dashboard`로 지정하세요. `index.html`, `app.js`, `styles.css`, `vendor/xlsx.full.min.js`가 기존 앱 화면이고 `api/weather.js`가 Node 서버 함수입니다. `data/11_날씨응답_예제.json`은 공개 가능한 가상 날씨 자료입니다. `vercel.json`이 정적 파일과 서버 함수 구성을 관리하므로 별도 빌드 명령, Output Directory, 의존성 설치는 필요하지 않습니다.

## 직접 배포 순서

1. **Upload**: 이 폴더의 파일만 기존 저장소에 올립니다. 상위 `.env.local`, `운영대시보드/data/`, `A_매출비교/`~`E_짧은연습/`, `results/`는 올리지 않습니다. `.env.example`은 빈 템플릿입니다.
2. **Import**: Vercel에서 저장소를 Import하고 Root Directory를 `vercel-dashboard`로 선택합니다. Framework Preset은 Other로 둡니다. Build Command와 Output Directory는 비워 두고 `vercel.json`을 사용합니다.
3. **Secret**: Project Settings → Environment Variables에 `OPENWEATHER_API_KEY`를 직접 추가하고 적용할 환경을 선택합니다. 키는 Vercel 서버 환경 변수에만 입력하세요. 키가 없으면 서버 함수가 교육용 예제 날씨로 응답합니다.
4. **Deploy**: Deploy를 실행하기 전에 계정·요금제, 앱을 볼 수 있는 범위, URL 공개 여부를 직접 확인하세요. 공개 URL 방문자도 날씨 조회를 요청할 수 있고 OpenWeather 사용량을 소비할 수 있습니다.

## 배포 내용과 데이터 상태

업무 원본을 제외했으므로 매출·행사준비 화면은 실제 자료가 연결되지 않은 상태를 보여줍니다. 해당 기능을 쓰려면 공개 가능한 가상 데이터만 별도 검토 후 추가하세요. 날씨 API 키는 브라우저 코드·브라우저 번들·응답·Git에 넣지 않습니다. 서버 함수가 `process.env.OPENWEATHER_API_KEY`를 사용해 OpenWeather를 호출합니다.
