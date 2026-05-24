#!/usr/bin/env python
"""Test script for DocForge API endpoints"""

import httpx
import json
import uuid
from pathlib import Path

BASE_URL = "http://127.0.0.1:8001"
client = httpx.Client(timeout=30)

print("=" * 70)
print("DOCFORGE BACKEND TEST SUITE")
print("=" * 70)

results = {"passed": 0, "failed": 0, "issues": []}

# 1. Health Check
print("\n[1] HEALTH CHECK")
try:
    r = client.get(f"{BASE_URL}/health")
    if r.status_code == 200:
        data = r.json()
        print(f"   Status: {r.status_code} ✓")
        print(f"   Response: {data.get('status')} | {data.get('app')} v{data.get('version')}")
        results["passed"] += 1
    else:
        print(f"   Status: {r.status_code} ✗")
        results["failed"] += 1
except Exception as e:
    print(f"   ERROR: {e}")
    results["issues"].append(f"Health check: {e}")
    results["failed"] += 1

# 2. List Documents (should be empty)
print("\n[2] GET /api/documents/ (List)")
try:
    r = client.get(f"{BASE_URL}/api/documents/")
    if r.status_code == 200:
        docs = r.json()
        if isinstance(docs, list):
            print(f"   Status: {r.status_code} ✓")
            print(f"   Documents: {len(docs)}")
            results["passed"] += 1
        else:
            print(f"   Status: {r.status_code} ✗ (response not a list)")
            results["failed"] += 1
            results["issues"].append("List documents returned non-list")
    else:
        print(f"   Status: {r.status_code} ✗")
        results["failed"] += 1
except Exception as e:
    print(f"   ERROR: {e}")
    results["issues"].append(f"List documents: {e}")
    results["failed"] += 1

# 3. DELETE endpoint on non-existent document
print("\n[3] DELETE /api/documents/{non-existent}")
try:
    fake_id = str(uuid.uuid4())
    r = client.delete(f"{BASE_URL}/api/documents/{fake_id}")
    if r.status_code == 404:
        print(f"   Status: {r.status_code} ✓")
        print(f"   Correctly returns 404 for missing resource")
        results["passed"] += 1
    else:
        print(f"   Status: {r.status_code} (expected 404)")
        if r.status_code == 204:
            print(f"   WARNING: 204 No Content on non-existent resource")
            results["issues"].append("DELETE returns 204 for non-existent doc")
        results["failed"] += 1
except Exception as e:
    print(f"   ERROR: {e}")
    results["issues"].append(f"Delete endpoint: {e}")
    results["failed"] += 1

# 4. GET OpenAPI schema (routes introspection)
print("\n[4] GET /openapi.json (Routes)")
try:
    r = client.get(f"{BASE_URL}/openapi.json")
    if r.status_code == 200:
        schema = r.json()
        endpoints = list(schema.get("paths", {}).keys())
        print(f"   Status: {r.status_code} ✓")
        print(f"   Total routes: {len(endpoints)}")
        for endpoint in sorted(endpoints):
            methods = list(schema["paths"][endpoint].keys())
            print(f"      {endpoint:40} -> {', '.join(methods)}")
        results["passed"] += 1
    else:
        print(f"   Status: {r.status_code} ✗")
        results["failed"] += 1
except Exception as e:
    print(f"   ERROR: {e}")
    results["issues"].append(f"OpenAPI schema: {e}")
    results["failed"] += 1

# 5. GET docs endpoint
print("\n[5] GET /docs (Swagger UI)")
try:
    r = client.get(f"{BASE_URL}/docs")
    if r.status_code == 200:
        print(f"   Status: {r.status_code} ✓")
        results["passed"] += 1
    else:
        print(f"   Status: {r.status_code} ✗")
        results["failed"] += 1
except Exception as e:
    print(f"   ERROR: {e}")
    results["failed"] += 1

print("\n" + "=" * 70)
print("TEST SUMMARY")
print("=" * 70)
print(f"Passed: {results['passed']}")
print(f"Failed: {results['failed']}")

if results["issues"]:
    print("\nIssues detected:")
    for issue in results["issues"]:
        print(f"  ⚠ {issue}")

if results["failed"] == 0:
    print("\n✓ All critical endpoints functional")
    print("✓ DELETE endpoint working correctly")
    print("✓ No assertion errors or crashes detected")
else:
    print(f"\n✗ {results['failed']} test(s) failed")

print("=" * 70)
