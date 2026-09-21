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
## Designated pools (source of truth for "proper pools")
Tabs: "Pool Sets" (Set | Pool ID | Pool Name), "Profiles" (Profile | Type | Set | Channels | Pools in set | Note),
"Agent Profiles" (Agent | Profile | Suggested | Match %).
Type = Primary | Secondary | Optional. Channels = "Chat+Email" or "Email only".
designated(agent, channel, type) = union of Pool Sets referenced by the agent's Profile rows of that type, where Channels allows the channel.
Compliance covers Chat and Email only; ignore Phone columns.
Optional pools are never counted as missing or extra.
Profiles: RTT HVU, EN HVU/DVIP, EN NA+GTV, RTT NA+GTV.
Fixtures: fixtures/designated_pools_seed.xlsx (no agent data).
