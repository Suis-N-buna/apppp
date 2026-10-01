# SSU DORM — 숭실대학교 기숙사 생활 도우미

HTML, CSS, JavaScript로 만든 기숙사 생활 도우미입니다. Vite로 빌드하고 Vercel에서 정적 사이트로 배포합니다.

## GitHub에 올리는 파일

이 README가 있는 폴더의 **내용물 전체를 GitHub 저장소 최상위에** 올리세요. `src` 폴더 구조는 유지해야 합니다.

```text
.gitignore
README.md
index.html
package.json
package-lock.json
vercel.json
src/
  app.js
  style.css
```

`배포파일.zip`을 사용한다면 먼저 압축을 풀고 위 파일들을 올리세요. ZIP 자체를 GitHub에 올리면 앱이 배포되지 않습니다. 숨김 파일인 `.gitignore`도 포함하세요.

`node_modules`, `dist`, `.npm-cache`, `.vercel`, 실제 `.env` 파일은 업로드하지 않습니다. AI Studio 전용 `metadata.json`, 기존 React·Tailwind 설정, `tsconfig.json`, `bun.lock`도 이 배포에는 필요하지 않습니다.

## Vercel 배포

1. Vercel에서 새 프로젝트를 만들고 업로드한 GitHub 저장소를 선택합니다.
2. 위 파일들을 저장소 최상위에 올렸다면 Root Directory를 기본값으로 둡니다. `vercel-upload` 폴더째 올렸다면 Root Directory를 `vercel-upload`로 지정합니다.
3. 아래 설정을 확인하고 Deploy를 누릅니다. 빌드 설정은 `vercel.json`에 포함되어 있습니다.

| 항목 | 값 |
| --- | --- |
| Framework Preset | Vite |
| Install Command | `npm ci` |
| Build Command | `npm run build` |
| Output Directory | `dist` |
| Node.js | `24.x` (`package.json`에 지정) |
| Environment Variables | 필요 없음 |

현재 코드는 Gemini나 외부 API를 호출하지 않으므로 `GEMINI_API_KEY`, `APP_URL`을 입력하지 않아도 됩니다.

공식 안내: [Vite 배포](https://vercel.com/docs/frameworks/frontend/vite), [빌드 설정](https://vercel.com/docs/project-configuration/vercel-json), [Node.js 버전](https://vercel.com/docs/functions/runtimes/node-js/node-js-versions).

## 로컬 실행

Node.js 24를 설치한 뒤 이 폴더에서 실행하세요.

```sh
npm ci
npm run dev
```

개발 주소는 `http://localhost:3000`입니다. Windows PowerShell에서 스크립트 실행 정책 오류가 나면 `npm` 대신 `npm.cmd`를 사용하세요.

배포용 빌드와 미리보기:

```sh
npm run build
npm run preview
```

## 현재 데이터 동작

세탁실·택배·대여·공지 데이터는 예시 데이터입니다. 시설 고장 신고는 사용 중인 브라우저의 `localStorage`에 저장되며 Excel/CSV로 내려받을 수 있습니다. 배포만으로 학교 시스템 연동, 관리자 접수, 사용자 간 실시간 데이터 공유가 이루어지지는 않습니다.
