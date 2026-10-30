@echo off
REM ============================================================================
REM Exchange & Intune MCP Server Windows Launcher
REM Usage:
REM   run.bat          (Starts the MCP server)
REM   run.bat --login  (Prompts one-time Microsoft browser login)
REM ============================================================================
uv run exchange-device-mcp-demo.py %*
