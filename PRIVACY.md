# Privacy Policy

**Last Updated:** April 6, 2026

## Overview

MCP Server for Excel ("ExcelMcp") is an open-source tool that enables AI assistants to interact with Microsoft Excel. This privacy policy explains how the software handles your data.

## Data Collection Summary

**ExcelMcp does NOT collect any data.** All operations are performed entirely on your local machine with no external network communication.

### What We DO NOT Collect

- ❌ **Telemetry** - No usage statistics, performance metrics, or error reports are collected
- ❌ **File contents** - We never access or transmit data from your Excel files
- ❌ **File names or paths** - File paths are never transmitted anywhere
- ❌ **Personal information** - No names, emails, or account information
- ❌ **Spreadsheet data** - Cell values, formulas, and data remain completely private
- ❌ **User accounts** - No registration or sign-in required
- ❌ **Analytics** - No tracking, no cookies, no user identification

### Zero Network Communication

ExcelMcp operates with **zero outbound network traffic**:
- No HTTP/HTTPS requests are made
- No data is sent to any external servers
- No telemetry or analytics services are integrated
- DAX and M code formatting is performed locally (offline)

## How It Works

ExcelMcp operates entirely on your local machine:

1. **Local Processing** - All Excel operations are performed locally via Microsoft's COM API
2. **Your Files Stay Local** - Excel files are read from and written to your local filesystem only
3. **Zero Network Usage** - No network communication of any kind

## Data Flow

When you use ExcelMcp with an AI assistant (like Claude):

1. You send a request to the AI assistant
2. The AI assistant calls ExcelMcp tools on your local machine
3. ExcelMcp performs the requested Excel operations locally
4. Results are returned to the AI assistant

**Note:** The AI assistant you use (e.g., Claude) has its own privacy policy governing how it handles your conversations and data. ExcelMcp only handles the local Excel operations and does not send any data externally.

## Third-Party Services

- **Microsoft Excel** - ExcelMcp requires Microsoft Excel installed on your machine. Excel is subject to Microsoft's privacy policy.
- **AI Assistants** - When used with AI assistants like Claude, those services have their own privacy policies.

## Open Source

ExcelMcp is open source software. You can review the complete source code at:
https://github.com/sbroenne/mcp-server-excel

## Security

- ExcelMcp runs with the same permissions as your user account
- It can only access files and Excel instances that your user account can access
- No elevated privileges are required or requested

## Children's Privacy

ExcelMcp does not knowingly collect any information from anyone, including children under 13 years of age.

## Changes to This Policy

If we make changes to this privacy policy, we will update the "Last Updated" date above and publish the updated policy in our GitHub repository.

## Contact

For questions about this privacy policy or the ExcelMcp project:

- **GitHub Issues:** https://github.com/sbroenne/mcp-server-excel/issues
- **Repository:** https://github.com/sbroenne/mcp-server-excel

---

**Summary:** ExcelMcp processes your Excel files entirely locally on your machine with zero network communication. We do not collect any telemetry, analytics, or personal data. Your file contents, file names, and all operations remain completely private.
