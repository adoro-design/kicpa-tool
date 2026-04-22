from sqlalchemy import create_engine, Column, Integer, String, Text, Date, SmallInteger, DateTime, Boolean, text
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.sql import func
import os
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./kicpa.db")
# Railway PostgreSQL URL이 postgres:// 로 시작할 경우 수정
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ── 모델 ───────────────────────────────────────────
class User(Base):
    __tablename__ = "kicpa_users"
    id         = Column(Integer, primary_key=True, index=True)
    username   = Column(String(50), unique=True, nullable=False)
    password   = Column(String(255), nullable=False)
    name       = Column(String(100), nullable=False)
    role       = Column(String(20), default="director")
    is_active  = Column(Boolean, default=True)
    created_at = Column(DateTime, server_default=func.now())

class Content(Base):
    __tablename__ = "kicpa_contents"
    id                  = Column(Integer, primary_key=True, index=True)
    year                = Column(SmallInteger, default=2026, index=True)
    shooting_month      = Column(String(20))
    course_name         = Column(Text)
    required_optional   = Column(String(50))
    original_code       = Column(String(100))
    category            = Column(Text)
    course_code         = Column(String(200))
    session_count       = Column(Integer)
    chapter_count       = Column(Integer)
    instructor          = Column(String(200))
    department          = Column(String(100), index=True)
    kicpa_manager       = Column(String(100))
    filming_consent     = Column(String(100))
    shooting_date       = Column(Date)
    shooting_time       = Column(String(100))
    shooting_format     = Column(String(100))
    location            = Column(String(200))
    has_quiz            = Column(String(50))
    quiz_count          = Column(Integer)
    materials_supply    = Column(String(100))
    video_marking       = Column(String(100))
    dev_outsource_date  = Column(Date)
    inspection_date     = Column(Date)
    open_date           = Column(Date)
    billing             = Column(String(100))
    billing_month       = Column(String(20))       # 비용청구 월 (엑셀 기준)
    custom_price        = Column(Integer)          # 별도 단가 (단가표 우선 적용)
    travel_expense      = Column(Integer)          # 출장비 (출장 과정에 추가)
    notes               = Column(Text)
    created_at          = Column(DateTime, server_default=func.now())
    updated_at          = Column(DateTime, server_default=func.now(), onupdate=func.now())

class PriceTable(Base):
    __tablename__ = "kicpa_price_table"
    id         = Column(Integer, primary_key=True, index=True)
    category   = Column(String(20), nullable=False)
    type_name  = Column(String(100), nullable=False)
    unit_price = Column(Integer)
    unit       = Column(String(20))
    note       = Column(String(200))
    is_active  = Column(Boolean, default=True)

class Document(Base):
    __tablename__ = "kicpa_documents"
    id         = Column(Integer, primary_key=True, index=True)
    doc_type   = Column(String(20), nullable=False)
    department = Column(String(100))
    period     = Column(String(50))
    file_path  = Column(String(500))
    created_by = Column(Integer)
    created_at = Column(DateTime, server_default=func.now())

def init_db():
    """테이블 생성 + 기본 데이터"""
    if os.getenv("RESET_DB") == "true":
        Base.metadata.drop_all(bind=engine)
    Base.metadata.create_all(bind=engine)

    # 신규 컬럼 자동 추가 마이그레이션 (이미 있으면 무시)
    # 각 ALTER TABLE을 별도 커넥션으로 실행해야 PostgreSQL 트랜잭션 오염 방지
    for col_def in [
        "ALTER TABLE kicpa_contents ADD COLUMN custom_price INTEGER",
        "ALTER TABLE kicpa_contents ADD COLUMN billing_month VARCHAR(20)",
        "ALTER TABLE kicpa_contents ADD COLUMN travel_expense INTEGER",
    ]:
        try:
            with engine.connect() as conn:
                conn.execute(text(col_def))
                conn.commit()
        except Exception:
            pass  # 이미 컬럼 존재

    db = SessionLocal()
    try:
        # 관리자 계정
        if not db.query(User).first():
            from passlib.context import CryptContext
            pwd = CryptContext(schemes=["bcrypt"])
            db.add(User(username="admin", password=pwd.hash("kicpa1234!"), name="관리자", role="admin"))
            db.commit()
        # 단가 데이터
        if not db.query(PriceTable).first():
            prices = [
                PriceTable(category="new_dev", type_name="크로마키",        unit_price=500000, unit="차시"),
                PriceTable(category="new_dev", type_name="FullVod (출장)",   unit_price=500000, unit="차시"),
                PriceTable(category="new_dev", type_name="태블릿형",         unit_price=500000, unit="차시"),
                PriceTable(category="new_dev", type_name="전자칠판형",       unit_price=500000, unit="차시"),
                PriceTable(category="porting", type_name="포팅",            unit_price=50000,  unit="챕터"),
                PriceTable(category="porting", type_name="편집포팅",         unit_price=160000, unit="챕터"),
                PriceTable(category="travel",  type_name="1 ~ 4시간",        unit_price=100000, unit="시간"),
                PriceTable(category="travel",  type_name="4시간 초과",        unit_price=None,   unit="시간", note="별도 협의"),
            ]
            db.add_all(prices)
            db.commit()
        # 단가 항목명 업데이트 (기존 데이터 문구 수정)
        renames = {
            "포팅 (동영상 무편집)": "포팅",
            "포팅 (동영상 편집)":  "편집포팅",
        }
        for old_name, new_name in renames.items():
            p = db.query(PriceTable).filter_by(type_name=old_name).first()
            if p:
                p.type_name = new_name
        db.commit()
    finally:
        db.close()
