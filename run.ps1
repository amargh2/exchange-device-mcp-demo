<#
.SYNOPSIS
    Exchange & Intune MCP Server - PowerShell Launcher
.DESCRIPTION
    Runs the MCP server using uv directly in Windows PowerShell.
.EXAMPLE
    .\run.ps1 --login
    Runs interactive browser sign-in.
.EXAMPLE
    .\run.ps1
    Starts the MCP server listener.
#>
uv run exchange-device-mcp-demo.py $args
