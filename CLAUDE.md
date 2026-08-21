# KICPA 콘텐츠 정산 관리 Tool — CLAUDE.md

## 프로젝트 개요

한국공인회계사회(KICPA) 온라인 강의 콘텐츠 개발·정산 업무를 통합 관리하는 웹 기반 시스템.
매달 반복되던 Excel 수작업(손익분석서·개발요청서 등)을 자동화하고, 구글 시트 연동으로 실시간 데이터 관리를 지원한다.

- **사이트명**: KICPA 콘텐츠 정산 관리
- **배포 환경**: Render.com (Web Service + PostgreSQL)
- **저장소**: GitHub Private (main 브랜치 push → 자동 배포)
- **로컬 개발**: SQLite (`kicpa.db`), 배포 환경은 PostgreSQL

---

## 기술 스택

| 영역 | 기술 |
|------|------|
| Backend | Python 3.12, FastAPI 0.111, SQLAlchemy 2.0 |
| Template | Jinja2 3.1, Bootstrap 5.3 |
| DB | PostgreSQL (Render) / SQLite (로컬) |
| 문서 생성 | openpyxl 3.1, python-docx 1.1 |
| 구글 시트 | gspread 6.1, google-auth 2.29 (Service Account) |
| 인증 | passlib[bcrypt], itsdangerous (세션 쿠키) |
| 프론트 | Chart.js 4.4.0, Bootstrap Icons, Iconify |
| 배포 | Render.com, Procfile, runtime.txt |

---

## 디렉터리 구조

```
kicpa-tool/
├── main.py              # 모든 라우트 + 비즈니스 로직
├── database.py          # SQLAlchemy 모델 정의 + init_db()
├── docgen.py            # 문서 자동생성 모듈 (4종)
├── requirements.txt
├── Procfile             # web: uvicorn main:app --host 0.0.0.0 --port $PORT
├── runtime.txt          # python-3.12.x
├── .env.example         # 환경변수 샘플
├── doc_templates/       # 문서 생성용 원본 템플릿
│   ├── tpl_pnl.xlsx         # 손익분석서 템플릿
│   ├── tpl_devreq.xlsx      # 개발요청서 첨부자료 템플릿 (콘텐츠개발내역 시트)
│   ├── tpl_devreq.docx      # 개발요청서 본문 템플릿
│   └── tpl_profile.docx     # 프로젝트 프로파일 템플릿
├── templates/           # Jinja2 HTML 템플릿
│   ├── base.html
│   ├── login.html
│   ├── error.html
│   ├── dashboard.html
│   ├── contents.html
│   ├── content_edit.html
│   ├── schedule.html
│   ├── billing.html
│   ├── documents.html
│   ├── import.html
│   ├── studio.html
│   ├── customers.html
│   ├── price_table.html
│   ├── calc_settings.html
│   ├── users.html
│   └── data.html
└── static/css/style.css
```

---

## DB 모델 (database.py)

### Content (kicpa_contents) — 핵심 테이블
| 컬럼 | 타입 | 설명 |
|------|------|------|
| year | SmallInteger | 연도 (기본 2026) |
| shooting_month | String(20) | 촬영월 (예: "1월") |
| course_name | Text | 과정명 |
| required_optional | String(50) | 필수/선택 |
| original_code | Text | 원본 과정코드 (Text — 길이 제한 없음) |
| category | Text | 카테고리 |
| course_code | String(200) | 과정코드 |
| session_count | Integer | 차시수 |
| chapter_count | Integer | 챕터수 |
| instructor | String(200) | 강사명 |
| department | String(100) | 담당부서 |
| kicpa_manager | String(100) | 한공회 담당자 |
| filming_consent | String(100) | 촬영동의 |
| shooting_date | Date | 촬영날짜 |
| shooting_time | String(100) | 촬영시간 |
| shooting_format | String(100) | 촬영형식 |
| location | String(200) | 장소 |
| has_quiz | String(50) | 퀴즈유무 |
| quiz_count | Integer | 퀴즈 총문항수 |
| materials_supply | String(100) | 교안수급 |
| video_marking | String(100) | 영상마킹 |
| dev_outsource_date | Date | 개발외주일 |
| inspection_date | Date | 검수일 |
| open_date | Date | 오픈일 |
| billing | String(100) | 청구여부 |
| billing_month | String(20) | 청구월 (예: "3월") |
| custom_price | Integer | 별도 단가 (수동 입력) |
| travel_hours | Integer | 출장 시간 |
| travel_days | Integer | 출장 일수 |
| travel_expense | Integer | 출장비 직접 입력 |
| notes | Text | 비고 |

### 기타 모델
- **User** (kicpa_users): 로그인 계정, role = "admin" | "director"
- **PriceTable** (kicpa_price_table): 촬영형식별 단가 이력
- **CalcSettings** (kicpa_calc_settings): 손익분석서 계산 파라미터
- **Document** (kicpa_documents): 생성된 문서 메타데이터
- **StudioRental** (kicpa_studio_rental): 스튜디오 대관료 내역
- **CustomerContact** (kicpa_customer_contacts): 부서별 고객 담당자 연락처

---

## 주요 라우트 (main.py)

| 경로 | 메서드 | 설명 |
|------|--------|------|
| `/login` | GET/POST | 로그인 |
| `/logout` | GET | 로그아웃 |
| `/` | GET | 대시보드 |
| `/contents` | GET | 콘텐츠 목록 (필터·페이지네이션) |
| `/content/edit` | GET | 콘텐츠 등록(`id` 없음) / 수정(`?id={id}`) 폼 |
| `/content/edit` | POST | 콘텐츠 저장 (등록 또는 수정) |
| `/schedule` | GET | 촬영 일정 (shooting_date 기준) |
| `/export` | GET | Excel 내보내기 |
| `/import` | GET | Excel 가져오기 페이지 |
| `/import` | POST | Excel 파일 업로드 가져오기 |
| `/import/gsheet` | POST | 구글 시트 동기화 |
| `/billing` | GET | 정산 관리 |
| `/documents` | GET | 문서 생성 페이지 |
| `/documents/generate` | POST | 4종 문서 ZIP 생성·다운로드 |
| `/studio` | GET | 스튜디오 대관료 목록 |
| `/studio/add` | POST | 스튜디오 대관료 추가 |
| `/studio/delete/{rid}` | POST | 스튜디오 대관료 삭제 |
| `/studio/restore-gsheet` | POST | 구글 시트에서 스튜디오 데이터 복원 |
| `/customers` | GET | 고객담당자 목록 |
| `/customers/add` | POST | 고객담당자 추가 |
| `/customers/edit/{cid}` | POST | 고객담당자 수정 |
| `/customers/delete/{cid}` | POST | 고객담당자 삭제 |
| `/price_table` | GET | 단가표 관리 |
| `/price_table/add` | POST | 단가 항목 추가 |
| `/price_table/update` | POST | 단가 항목 수정 |
| `/price_table/delete/{pid}` | POST | 단가 항목 삭제 |
| `/calc_settings` | GET | 손익분석서 설정 |
| `/calc_settings/add` | POST | 설정값 추가 |
| `/calc_settings/delete/{sid}` | POST | 설정값 삭제 |
| `/users` | GET | 사용자 관리 (admin) |
| `/users/add` | POST | 사용자 추가 (admin) |
| `/users/toggle` | POST | 사용자 활성/비활성 토글 (admin) |
| `/users/change_pw` | POST | 비밀번호 변경 (admin) |
| `/data` | GET | 데이터 관리 — DB 백업·복원, 연도별 데이터 삭제 (admin) |
| `/backup/download` | GET | DB 전체 백업 JSON 다운로드 (admin) — 저장 위치: `D:\Work\P08.AI\04.한공회콘텐츠관리Tool\03.Data백업` |
| `/backup/restore` | POST | JSON 파일로 DB 복원 (admin) |
| `/admin/year-stats` | GET | 연도별 콘텐츠 건수 JSON 반환 (admin) |
| `/admin/delete-year` | POST | 특정 연도 콘텐츠 전체 삭제 (admin) |

---

## 구글 시트 연동

### 설정 (Render 환경변수)
```
GOOGLE_CLIENT_EMAIL=서비스계정@project.iam.gserviceaccount.com
GOOGLE_PRIVATE_KEY=-----BEGIN RSA PRIVATE KEY-----\n...
GOOGLE_SPREADSHEET_ID=스프레드시트ID
```

### 탭 네이밍 규칙
- `개발관리_2026`, `개발관리_2025`, `개발관리_2024` ...
- 연도별로 별도 탭, 같은 스프레드시트

### 구글 시트 컬럼 매핑 (0-indexed, 4행부터 데이터)
| 인덱스 | 열 | 필드 |
|--------|-----|------|
| 1 | B | 촬영월 |
| 2 | C | 과정명 |
| 3 | D | 필수/선택 |
| 4 | E | 원본코드 |
| 5 | F | 카테고리 |
| 6 | G | 과정코드 |
| 7 | H | 차시수 |
| 8 | I | 챕터수 |
| 9 | J | 강사 |
| 10 | K | 담당부서 |
| 11 | L | 한공회 담당자 |
| 12 | M | 촬영동의 |
| 13 | N | 촬영날짜 |
| 14 | O | 촬영시간 |
| 15 | P | 촬영형식 |
| 16 | Q | 장소 |
| 17 | R | 퀴즈유무 |
| 18 | S | 퀴즈 총문항수 |
| 19 | T | 교안수급 |
| 20 | U | 영상마킹 |
| 21 | V | 개발외주일 |
| 22 | W | 검수일 |
| 23 | X | 오픈일 |
| 24 | Y | 비용청구 여부 (● 표시 = 청구완료, `billing` 필드에 저장) |
| 25 | Z | 청구월 |
| 26 | AA | 비고 |

### 날짜 파싱 (`to_date_str`)
구글 시트는 날짜를 `"01월 05일"` (연도 없는 한국식) 형식으로 반환한다.
`to_date_str(v, yr=year)` 함수가 다음 형식을 모두 처리:
- ISO: `2026-01-05`
- 슬래시: `2026/01/05`
- 한국식(연도 포함): `2026. 1. 5`
- **한국식(연도 없음)**: `01월 05일` → `yr`년으로 자동 보완

### append 모드 동작
- 기존 레코드가 있으면 업데이트, 없으면 새로 추가
- `shooting_date`, `open_date`: 새 값이 null이면 **기존 값 유지** (덮어쓰기 방지)
- `custom_price`, `travel_hours`, `travel_expense`, `notes`: 수동 입력값은 항상 보존

---

## 문서 자동생성 (docgen.py)

### 4종 문서
| 파일명 | 형식 | 내용 |
|--------|------|------|
| 손익분석서_한국공인회계사회_MM월_V1.0_MMDD_부서.xlsx | XLSX | 매출·인건비·스튜디오·손익률 자동 계산 |
| 프로젝트프로파일_한국공인회계사회_MM월_V0.1_MMDD_부서.docx | DOCX | 계약금액·기간·투입인력 요약 |
| 한공회_YYYY년MM월분_컨텐츠개발요청서_부서.docx | DOCX | 과정 목록·차시·단가·금액 명세 |
| 한공회_YYYY년MM월분_컨텐츠개발요청서_제출시첨부사항_부서.xlsx | XLSX | 상세 내역 (퀴즈문항수 포함) |

### 첨부자료 Excel (tpl_devreq.xlsx) 컬럼 구조
`콘텐츠개발내역` 시트, 행2=헤더, 행3~=데이터 (부서행 삽입 후 행3→헤더, 행4~→데이터)

| col | 헤더 | 데이터 |
|-----|------|--------|
| 1(A) | 연번 | 순서 |
| 2(B) | 구분 | 신규/포팅/출장비 |
| 3(C) | 강좌명 | course_name |
| 4(D) | 강사명 | instructor |
| 5(E) | 강사소속 | (공란) |
| 6(F) | 시간(차시) | session_count |
| 7(G) | 챕터수 | chapter_count |
| 8(H) | 등록일(수정일) | open_date |
| 9(I) | 제작유형 | shooting_format |
| 10(J) | 단가기준 유형 | shooting_format |
| 11(K) | **퀴즈** | **quiz_count** |
| 12(L) | 단가 | unit_price |
| 13(M) | 금액 | =L*F (수식) |
| 14(N) | 총액(VAT 포함) | =M*1.1 (수식) |

금액 관련 열(12, 13, 14): `number_format = '#,##0'`, `alignment = right`

### 촬영형식 분류 (`classify_fmt`)
- 포팅 없음 → 신규 (단가: 차시당 500,000원)
- "포팅" 포함, "편집" 없음 → 포팅 (단가: 챕터당 50,000원)
- "편집포팅" → 편집포팅 (단가: 챕터당 160,000원)
- "출장" 포함 → FullVod 출장 (단가: 차시당 500,000원 + 출장비)

### 출장비 계산
- 1일 4시간 초과 시 4시간으로 캡 적용
- 우선순위: `travel_expense` 직접입력 > 시간 계산 > 기본 1시간

---

## 대시보드 (dashboard.html)

### KPI 카드 (차시수 기준)
- 전체 차시 / 촬영 완료 차시 / 오픈 완료 차시 / 청구 완료 차시

### 현재 연도 레이아웃 (`is_current_year = True`)
- Row 1: KPI 4개
- Row 2: 월별 매출추이 + 미오픈 현황
- Row 3: 부서별 월별 매출비교 + 이번달 촬영예정
- Row 4: 부서별 매출현황 + 촬영형식 분포

### 이전 연도 레이아웃 (`is_current_year = False`)
- Row 1: KPI 4개
- Row 2: 월별 매출추이 + 부서별 매출현황
- Row 3: 부서별 월별 매출비교 + 촬영형식 분포

### Chart.js 차트 (4개 함수)
1. `drawMonthlyChart()` — 월별 매출 막대 차트
2. `drawFormatChart()` — 촬영형식 도넛 차트
3. `drawDeptChart()` — 부서별 매출 가로 막대 차트
4. `drawDeptMonthlyChart()` — 부서별 월별 그룹 막대 차트

---

## 정산 관리 차트 (billing.html)

부서가 2개 이상일 때만 표시:
- **도넛 차트**: 부서별 매출 비중 (%)
- **가로 막대 차트**: 부서별 총액 vs 정산완료 비교 (단위: 만원)

Chart.js는 billing.html `{% block head %}`에서 CDN 로드 (base.html에는 없음).

---

## 단가표 관리

- `PriceTable.effective_from`: 적용 시작일 기준으로 이력 관리
- `get_price_table_for_month(db, year, month)`: 해당 월 기준 유효 단가 조회
- 카테고리: `new_dev` (신규), `porting` (포팅), `travel` (출장)

---

## 계산 설정 (CalcSettings)

| setting_name | 기본값 | 설명 |
|---|---|---|
| work_hours_chromakey | 2.5 | 크로마키·태블릿·전자칠판형 차시당 작업시간 |
| work_hours_porting | 0.5 | 포팅(무편집) 차시당 작업시간 |
| work_hours_edit_porting | 1.0 | 포팅(편집) 차시당 작업시간 |
| work_hours_travel | 3.5 | 출장 차시당 작업시간 (개발2.5+출장1.0) |
| target_profit_pct | 30.0 | 최소 손익률 (%) |
| travel_cap_hours | 4.0 | 출장비 일일 한도 (시간) |
| work_hours_per_day | 8.0 | 1일 근무시간 |

PM 투입비율: 1% 고정. PROD 투입비율: 표준시간 기반 자동 계산, 손익률 30% 미달 시 자동 축소.

---

## 사용자 역할

| role | 접근 권한 |
|------|-----------|
| admin | 전체 메뉴 (Excel 가져오기·내보내기·설정·사용자 관리 포함) |
| director | 대시보드·콘텐츠 목록·촬영일정·정산관리·문서생성·스튜디오 대관 (등록·수정·삭제 가능) |
| viewer | director와 동일 메뉴 조회만 가능 — 등록·수정·삭제·문서생성 불가 (읽기 전용) |

**Admin 전용**: Excel 가져오기/내보내기, 고객담당자, 단가표 관리, 손익분석서 설정, 사용자 관리, 데이터 관리(백업·복원·연도 삭제)

**권한 함수**: `require_admin` (admin only) / `require_editor` (admin+director) / `require_login` (전체)

세션 쿠키 기반 인증 (`itsdangerous.TimestampSigner`).

---

## DB 백업 및 복원

### 배경
Render.com 무료 플랜은 PostgreSQL DB가 비활성 상태 지속 시 초기화될 수 있다.
초기화 시 시드 데이터(기본 계정·단가표·설정값)는 자동 복구되지만, 아래 데이터는 **영구 소실**된다.

- 콘텐츠 목록 전체 (과정명·촬영일·정산 현황 등)
- 수동 입력한 별도 단가 (`custom_price`), 출장비 (`travel_expense`)
- 스튜디오 대관 내역
- 고객담당자 연락처
- 추가 생성한 사용자 계정

### 백업 방법
1. 사이드바 **데이터 관리 → 백업 · 복원** 메뉴 이동
2. "백업 파일 다운로드" 클릭 → `kicpa_backup_YYYYMMDD_HHMMSS.json` 다운로드
3. 아래 경로에 저장:
   ```
   D:\Work\P08.AI\04.한공회콘텐츠관리Tool\03.Data백업\
   ```

### 자동 백업
- **Windows 작업 스케줄러**로 매주 금요일 오전 10시 자동 실행 (`KICPA_AutoBackup` 작업)
- 스크립트: `D:\Work\P08.AI\04.한공회콘텐츠관리Tool\03.Data백업\kicpa_auto_backup.py`
- `/backup/auto?token=BACKUP_TOKEN` 엔드포인트 호출 → JSON 저장 (토큰은 Render 환경변수 `BACKUP_TOKEN`에 설정)
- PC가 켜져 있으면 Cowork 실행 여부와 무관하게 동작

### 수동 백업 권장 시점
- 대량 데이터 입력 후 즉시 (구글 시트 동기화, Excel 가져오기 직후)

### 복원 방법 (DB 초기화 발생 시)
1. 사이드바 **데이터 관리 → 백업 · 복원** 메뉴 이동
2. "복원" 영역에서 백업 파일(`kicpa_backup_*.json`) 선택
3. "복원" 버튼 클릭 → 확인 후 실행
4. 복원 완료 메시지에서 테이블별 복원 건수 확인

> ⚠️ 복원 시 현재 DB의 모든 데이터가 백업 파일 내용으로 **전체 교체**된다. 복원 후 PK 시퀀스는 자동 리셋되어 신규 등록 시 충돌 없음.

### 백업 파일 형식
```json
{
  "created_at": "2026-08-21T10:00:00",
  "version": "1",
  "tables": {
    "users": [...],
    "contents": [...],
    "price_table": [...],
    "calc_settings": [...],
    "studio_rentals": [...],
    "customer_contacts": [...]
  }
}
```

---

## 환경변수

```
DATABASE_URL=postgresql://...       # Render PostgreSQL URL
SECRET_KEY=임의의_시크릿키
RESET_DB=false                      # true로 설정 시 DB 초기화 (위험)
GOOGLE_CLIENT_EMAIL=...             # 구글 서비스 계정 이메일
GOOGLE_PRIVATE_KEY=...              # 구글 서비스 계정 개인키 (\n 이스케이프)
GOOGLE_SPREADSHEET_ID=...          # 구글 스프레드시트 ID
```

---

## 개발 규칙 및 주의사항

### DB 마이그레이션
- `init_db()` 시작 시 `ALTER TABLE ... ADD COLUMN`으로 신규 컬럼 추가 (이미 있으면 무시)
- Render 배포 후 자동 실행됨
- 컬럼 타입 변경은 별도 마이그레이션 필요 (예: `original_code VARCHAR(100) → TEXT`)

### 데이터 가져오기 규칙
- **구글 시트**: 같은 스프레드시트, 연도별 탭 (`개발관리_2026`)
- **로컬 Excel**: 이전 연도 또는 백업용, `개발관리` 시트, 4행부터 데이터
- append 모드: 과정명 기준으로 기존 레코드 업데이트, 없으면 추가
- replace 모드: 해당 연도 전체 삭제 후 재입력 (수동 입력값은 과정명 기준 복원)

### 부서명
실제 사용 중인 부서명 (임의 추정 금지):
- 감사인증기준본부
- 재무보고본부
- 조세지원본부
- 지속가능성본부
- 회계연수원

### 코드 작성 원칙
- 모든 라우트는 `main.py` 단일 파일에 작성
- 문서 생성 로직은 `docgen.py`에 분리
- Jinja2 커스텀 필터: `normalize_month`, `clean_name`, `fmt_date`
- MONTH_ORDER, MONTHS 상수로 월 정렬 통일
- `func.coalesce(func.sum(Content.session_count), 0)` 패턴으로 차시수 합산

### 로컬 개발
```bash
cd kicpa-tool
pip install -r requirements.txt
uvicorn main:app --reload
# http://localhost:8000
# 초기 계정: admin / kicpa1234!
```

---

## 관련 파일 (프로젝트 루트)

| 파일 | 설명 |
|------|------|
| `kicpa_showcase.html` | 12슬라이드 웹 프레젠테이션 (상하 이동, Chart.js 포함) |
