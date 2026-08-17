import os
import shutil
import sys
import tempfile
from pathlib import Path

import pytest


BACKEND_ROOT = Path(__file__).resolve().parents[1]
if str(BACKEND_ROOT) not in sys.path:
    sys.path.insert(0, str(BACKEND_ROOT))

# Configure an isolated application database before any test module imports
# app.config/app.database. This is especially important for concurrency tests
# that intentionally exercise the real global engine from worker threads.
TEST_ROOT = Path(tempfile.mkdtemp(prefix="email-bulk-tests-"))
os.environ["APP_ENV"] = "testing"
os.environ["DEBUG"] = "false"
os.environ["DATABASE_URL"] = f"sqlite:///{(TEST_ROOT / 'app.db').as_posix()}"
os.environ["UPLOAD_DIR"] = str(TEST_ROOT / "uploads")
os.environ["RATE_LIMIT_BACKEND"] = "memory"


@pytest.fixture(autouse=True)
def dispose_test_engines(request, monkeypatch):
    """Dispose engines created directly by a test module after every test.

    SQLite connections otherwise remain parked in SQLAlchemy pools until
    garbage collection, which is both noisy on Python 3.13 and can retain file
    handles long enough to make tests order-dependent on Windows.
    """
    module = request.module
    create_engine = getattr(module, "create_engine", None)
    engines = []

    if create_engine is not None:
        def tracked_create_engine(*args, **kwargs):
            engine = create_engine(*args, **kwargs)
            engines.append(engine)
            return engine

        monkeypatch.setattr(module, "create_engine", tracked_create_engine)

    try:
        yield
    finally:
        for engine in reversed(engines):
            engine.dispose()


@pytest.fixture(scope="session", autouse=True)
def shutdown_application_resources():
    """Close singleton workers and the application pool at test-session end."""
    try:
        yield
    finally:
        from app.database import engine
        from app.services import send_engine

        instance = getattr(send_engine, "_engine_instance", None)
        if instance is not None:
            instance.shutdown(timeout=10)
        engine.dispose()
        shutil.rmtree(TEST_ROOT, ignore_errors=True)
