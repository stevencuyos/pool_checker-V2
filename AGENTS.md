# Chat Pool ID Checker
Google Apps Script web app. Server: Code.gs. Client: Dashboard.html (served via doGet).
Cannot run locally; no access to the real spreadsheet.

## Sheets
- PHteams: A=agent LDAP, B=team, D=team list, F=agent list
- "Primary enabled per LDAP" / "Secondary enabled per LDAP": row 1 = agent LDAP headers, pool IDs below each
- "Pool ID Names": A=name, B=ID
- Counts: A=agent, B=team, C=primary, D=secondary, E=total (from row 3)
- [ADD: designated pools sheet + layout]
- [ADD: logged-in pools sheet + layout]

## Conventions
- Read sheets once with getDisplayValues(); no getRange calls inside loops
- Normalize IDs with String(x).trim()
- Client uses google.script.run, escapeHtml(), toast(), showTab()
- Keep the Google-style CSS (#1a73e8, #dadce0, pill badges)
- Put calculations in pure functions (no SpreadsheetApp) so they can be tested in Node with sample data
