# testing_app/testing_database.py

from sqlalchemy import create_engine, Column, Integer, String, DateTime
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from datetime import datetime

# Separate SQLite DB only for testing
TESTING_DATABASE_URL = "sqlite:///./testing_notices.db"

testing_engine = create_engine(
    TESTING_DATABASE_URL,
    connect_args={"check_same_thread": False},
)

TestingBase = declarative_base()

TestingSessionLocal = sessionmaker(
    autocommit=False,
    autoflush=False,
    bind=testing_engine,
)


class TestingNotice(TestingBase):
    """
    Minimal model for testing pipeline.
    One row per uploaded testing PDF.
    """
    __tablename__ = "testing_notices"

    id = Column(Integer, primary_key=True, index=True)
    pdf_filename = Column(String(255), nullable=False)
    pdf_path = Column(String(500), nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)


def get_testing_db():
    """
    Dependency for FastAPI routes in testing_app.main.
    """
    db = TestingSessionLocal()
    try:
        yield db
    finally:
        db.close()


# Call this once at startup (from testing_app/main.py) to create tables
def init_testing_db():
    TestingBase.metadata.create_all(bind=testing_engine)
