# fantastic-journey

## Get-GcAgentBehavior.ps1 — Genesys Cloud agent behavior audit (CLM-safe)

A single-file PowerShell 5.1 tool for Genesys Cloud that builds a per-agent
behavior picture and exports it to CSV. Written specifically for
**Constrained Language Mode** endpoints (AppLocker/WDAC): no .NET static
calls, no `::new()`, no `[pscustomobject]` casts — Base64 for the OAuth
Basic header is implemented in pure PowerShell bit math.

### What it reports

| Output file | Contents | API used |
|---|---|---|
| `Agents.csv` | Directory of matched agents | `GET /api/v2/users` |
| `CallActivity.csv` | Offered/answered, talk/hold/ACW/handle hours, AHT, no-answer count | `POST /api/v2/analytics/conversations/aggregates/query` |
| `StatusTime.csv` | Hours per presence (Available, Away, Break, On Queue…) and routing status (IDLE, INTERACTING…) | `POST /api/v2/analytics/users/aggregates/query` |
| `WfmAdherence.csv` | Real-time WFM adherence state, scheduled vs actual activity, impact | `GET /api/v2/workforcemanagement/adherence` |
| `KbFeedback.csv` | Each KB article feedback submission per agent (rating, reason, comment) | `GET /api/v2/knowledge/knowledgebases/…/documents/…/feedback` |
| `CopilotSummaries.csv` | AI Copilot summaries per conversation: feedback given, whether the agent edited the note | `GET /api/v2/conversations/{id}/summaries` |
| `AgentBehaviorSummary.csv` | One row per agent with everything joined | — |

### Usage

```powershell
# Full report, last 14 days, one team
.\Get-GcAgentBehavior.ps1 -ClientId $id -ClientSecret $secret `
    -Region 'mypurecloud.ie' -DaysBack 14 -AgentEmailFilter '*@contoso.com'

# Skip the heavier sections
.\Get-GcAgentBehavior.ps1 -ClientId $id -ClientSecret $secret -SkipKb -SkipCopilot

# Edit an AI Copilot note (conversation summary)
.\Get-GcAgentBehavior.ps1 -ClientId $id -ClientSecret $secret -Region 'usw2.pure.cloud' `
    -EditCopilotNote -ConversationId 'abc-123' -SummaryId 'def-456' `
    -NewNoteText 'Corrected summary text'
```

### OAuth client requirements

Client Credentials grant with a role granting:
`analytics:conversationAggregate:view`, `analytics:userAggregate:view`,
`analytics:conversationDetail:view`, `directory:user:view`,
`wfm:realtimeAdherence:view`, `knowledge:knowledgebase:view`,
`knowledge:document:view`, plus conversation summary view/edit for the
Copilot sections.

### Notes

- KB feedback and Copilot scanning cost one API call per document /
  conversation, so they are capped by `-MaxKbDocuments` (500) and
  `-MaxConversationsForCopilot` (200). Raise the caps for full coverage.
- The Copilot summary endpoints are the newest API surface and Genesys still
  iterates on them. If they 404 in your org, verify the paths in the
  [API Explorer](https://developer.genesys.cloud/devapps/api-explorer) and
  adjust the two `$script:Copilot*PathTemplate` variables at the top of the
  script — nothing else needs to change.
- The wrapper handles 429 rate limits (backoff + retry), mid-run token
  expiry (auto re-auth on 401), and missing permissions (403 skips the
  section with a warning instead of failing the run).
