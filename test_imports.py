import sys
print("Python:", sys.version)

try:
    import fastmcp
    print("fastmcp:", fastmcp.__version__)
except Exception as e:
    print("fastmcp ERROR:", e)

try:
    from playwright.sync_api import sync_playwright
    print("playwright: ok")
except Exception as e:
    print("playwright ERROR:", e)

try:
    import dotenv
    print("dotenv: ok")
except Exception as e:
    print("dotenv ERROR:", e)

# MCP 서버 import 테스트
try:
    import importlib.util
    spec = importlib.util.spec_from_file_location("erp", r"C:\mcp\erp_groupware\erp_groupware_mcp.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    print("MCP server module: ok")
except Exception as e:
    print("MCP server module ERROR:", e)
