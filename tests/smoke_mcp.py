"""Smoke test: start server.py over stdio, initialize, list tools, call list_open_decks + lint_alt_text."""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
py = ROOT / ".venv" / "Scripts" / "python.exe"
p = subprocess.Popen([str(py), str(ROOT / "server.py")], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                     stderr=subprocess.PIPE, text=True, encoding="utf8")


def send(msg):
    p.stdin.write(json.dumps(msg) + "\n")
    p.stdin.flush()


def recv():
    while True:
        line = p.stdout.readline()
        if not line:
            raise SystemExit("server closed stdout: " + p.stderr.read()[-2000:])
        d = json.loads(line)
        if "id" in d:
            return d


send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
      "params": {"protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "smoke", "version": "0"}}})
init = recv()
print("server:", init["result"]["serverInfo"], "protocol", init["result"]["protocolVersion"])
send({"jsonrpc": "2.0", "method": "notifications/initialized"})
send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
tools = recv()["result"]["tools"]
print(len(tools), "tools:", ", ".join(t["name"] for t in tools))
send({"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "lint_alt_text",
                                                                     "arguments": {"text": "image of chart.png"}}})
print("lint:", recv()["result"]["content"][0]["text"])
send({"jsonrpc": "2.0", "id": 4, "method": "tools/call", "params": {"name": "list_open_decks", "arguments": {}}})
print("open decks:", recv()["result"]["content"][0]["text"])
p.stdin.close()
p.terminate()
sys.exit(0 if len(tools) >= 46 else 1)
