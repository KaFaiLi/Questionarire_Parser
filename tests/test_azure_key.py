# tests/test_azure_key.py
import os, sys; from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import analyze_kyd_json as akj

def test_env_wins_over_config():
    os.environ["AZURE_OPENAI_API_KEY"] = "env-key"
    try:
        assert akj.azure_api_key({"api_key": "cfg-key"}) == "env-key"
    finally:
        del os.environ["AZURE_OPENAI_API_KEY"]

def test_falls_back_to_config():
    os.environ.pop("AZURE_OPENAI_API_KEY", None)
    assert akj.azure_api_key({"api_key": "cfg-key"}) == "cfg-key"

def test_raises_when_missing():
    os.environ.pop("AZURE_OPENAI_API_KEY", None)
    try:
        akj.azure_api_key({})
        assert False, "expected SystemExit"
    except SystemExit:
        pass

if __name__ == "__main__":
    test_env_wins_over_config()
    test_falls_back_to_config()
    test_raises_when_missing()
    print("OK")
