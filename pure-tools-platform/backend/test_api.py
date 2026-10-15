"""Test script for the Pure Tools Platform API"""
import httpx
import sys

BASE = "http://127.0.0.1:8000"

def test_api():
    with httpx.Client(base_url=BASE) as client:
        # 1. Health check
        print("=== 1. Health Check ===")
        r = client.get("/api/health")
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code == 200, "Health check failed"

        # 2. Register
        print("\n=== 2. Register ===")
        r = client.post("/api/auth/register", json={
            "username": "testuser",
            "email": "test@gmail.com",
            "password": "Test1234!",
            "full_name": "Test User"
        })
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code in (201, 400), f"Register unexpected: {r.status_code}"

        # 3. Login
        print("\n=== 3. Login ===")
        r = client.post("/api/auth/login", json={
            "username": "testuser",
            "password": "Test1234!"
        })
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code == 200, "Login failed"
        token = r.json()["access_token"]
        headers = {"Authorization": f"Bearer {token}"}

        # 4. Get current user
        print("\n=== 4. Get /api/auth/me ===")
        r = client.get("/api/auth/me", headers=headers)
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code == 200, "Get me failed"

        # 5. Create a tool
        print("\n=== 5. Create Tool ===")
        r = client.post("/api/tools/", json={
            "name": "Convert ESPF to RDS",
            "description": "Converts ESPF files to RDS format",
            "icon": "transform",
            "category": "Conversion",
            "script_path": "pure-convertespftords/main.py"
        }, headers=headers)
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code in (201, 400), f"Create tool unexpected: {r.status_code}"
        tool_id = r.json().get("id")

        # 6. List tools
        print("\n=== 6. List Tools ===")
        r = client.get("/api/tools/", headers=headers)
        print(f"  Status: {r.status_code} | Tools count: {len(r.json())}")
        for t in r.json():
            print(f"    - {t['name']} ({t['category']})")
        assert r.status_code == 200, "List tools failed"

        # 7. Get single tool
        if tool_id:
            print(f"\n=== 7. Get Tool {tool_id} ===")
            r = client.get(f"/api/tools/{tool_id}", headers=headers)
            print(f"  Status: {r.status_code} | Body: {r.json()}")
            assert r.status_code == 200, "Get tool failed"

        # 8. Unauthorized access (no token)
        print("\n=== 8. Unauthorized Access (no token) ===")
        r = client.get("/api/tools/")
        print(f"  Status: {r.status_code} | Body: {r.json()}")
        assert r.status_code == 403, "Should be forbidden without token"

        print("\n" + "=" * 50)
        print("ALL TESTS PASSED!")
        print("=" * 50)

if __name__ == "__main__":
    try:
        test_api()
    except Exception as e:
        print(f"\nTEST FAILED: {e}", file=sys.stderr)
        sys.exit(1)
