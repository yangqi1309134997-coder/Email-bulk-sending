from fastapi.testclient import TestClient

from app.api.deps import get_current_user
from app.main import app


def test_static_preset_route_is_not_shadowed_by_sender_id_route():
    app.dependency_overrides[get_current_user] = lambda: object()
    try:
        response = TestClient(app).get("/api/senders/presets")
    finally:
        app.dependency_overrides.clear()

    assert response.status_code == 200
    assert any(preset["key"] == "gmail" for preset in response.json())

