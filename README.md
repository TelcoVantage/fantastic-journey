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
| `MediaTypeBreakdown.csv` | Interactions + handle hours per agent split by media type (voice, chat, email, message…) | `POST /api/v2/analytics/conversations/aggregates/query` |
| `WrapUpCodes.csv` | Wrap-up/disposition code distribution per agent (code ids resolved to names) | conversation aggregates + `GET /api/v2/routing/wrapupcodes` |
| `WfmAdherence.csv` | Real-time WFM adherence state, scheduled vs actual activity, impact | `GET /api/v2/workforcemanagement/adherence` |
| `EvaluationScores.csv` | QM evaluation count, avg total score %, avg critical score % per agent | `POST /api/v2/analytics/evaluations/aggregates/query` |
| `SurveyScores.csv` | Survey aggregates per agent: scored count, avg total score %, sent/started/abandoned | `POST /api/v2/analytics/surveys/aggregates/query` |
| `SurveyResponses.csv` | Individual survey responses with NPS score + Promoter/Passive/Detractor band | `GET /api/v2/quality/conversations/{id}/surveys` + survey form lookup |
| `Sentiment.csv` | Speech & Text Analytics sentiment score/trend per conversation, attributed to the agent(s) | `GET /api/v2/speechandtextanalytics/conversations/{id}` |
| `KbFeedback.csv` | Each KB article feedback submission per agent (rating, reason, comment) | `GET /api/v2/knowledge/knowledgebases/…/documents/…/feedback` |
| `CopilotSummaries.csv` | AI Copilot summaries per conversation: feedback given, whether the agent edited the note | `GET /api/v2/conversations/{id}/summaries` |
| `AgentBehaviorSummary.csv` | One row per agent with everything joined, including computed **NPS** = (promoters − detractors) ÷ responses × 100 | — |

### Setup (embedded credentials)

The script defaults to the **Australia region** (`mypurecloud.com.au`) and
reads credentials from the `EMBEDDED CREDENTIALS` block near the top of the
file. On your **local copy only**, replace the two placeholders:

```powershell
$script:EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$script:EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
```

> **Never commit real credentials back to this repo** — git history and pull
> requests are readable by everyone with repo access. If a secret is ever
> committed, delete/rotate that OAuth client in Genesys Cloud Admin
> immediately. Command-line `-ClientId`/`-ClientSecret`/`-Region` always
> override the embedded values, so you can also keep the placeholders and
> pass credentials at runtime.

### Usage

```powershell
# Full report, last 14 days, one team (embedded creds, AU region)
.\Get-GcAgentBehavior.ps1 -DaysBack 14 -AgentEmailFilter '*@contoso.com'

# Skip the heavier sections
.\Get-GcAgentBehavior.ps1 -SkipKb -SkipCopilot

# Edit an AI Copilot note (conversation summary)
.\Get-GcAgentBehavior.ps1 -EditCopilotNote -ConversationId 'abc-123' `
    -SummaryId 'def-456' -NewNoteText 'Corrected summary text'
```

### OAuth client requirements

Client Credentials grant with a role granting:
`analytics:conversationAggregate:view`, `analytics:userAggregate:view`,
`analytics:conversationDetail:view`, `analytics:surveyAggregate:view`,
`analytics:evaluationAggregate:view`, `quality:survey:view`,
`quality:surveyForm:view`, `speechAndTextAnalytics:data:view`,
`routing:wrapupCode:view`, `directory:user:view`,
`wfm:realtimeAdherence:view`, `knowledge:knowledgebase:view`,
`knowledge:document:view`, plus conversation summary view/edit for the
Copilot sections. Missing permissions skip their section with a warning
instead of failing the run.

### Notes

- KB feedback and the per-conversation sections (Copilot summaries,
  sentiment, survey responses) cost 1-3 API calls per document /
  conversation, so they are capped by `-MaxKbDocuments` (500) and
  `-MaxScanConversations` (200, alias `-MaxConversationsForCopilot`).
  Raise the caps for full coverage, or use `-SkipSurveys`,
  `-SkipEvaluations`, `-SkipSentiment`, `-SkipCopilot`, `-SkipKb`,
  `-SkipWfm` to trim the run.
- NPS is extracted by looking up each survey's form and finding questions of
  NPS type, then classifying the 0-10 answer (9-10 promoter, 7-8 passive,
  0-6 detractor). The roll-up computes per-agent NPS from those bands. The
  aggregate `SurveyScores.csv` average is a separate 0-100 total-score
  metric, not NPS.
- Sentiment requires Speech & Text Analytics sentiment analysis to be
  enabled for the relevant queues/flows; conversations without STA data are
  silently skipped.
- The Copilot summary endpoints are the newest API surface and Genesys still
  iterates on them. If they 404 in your org, verify the paths in the
  [API Explorer](https://developer.genesys.cloud/devapps/api-explorer) and
  adjust the two `$script:Copilot*PathTemplate` variables at the top of the
  script — nothing else needs to change.
- The wrapper handles 429 rate limits (backoff + retry), mid-run token
  expiry (auto re-auth on 401), and missing permissions (403 skips the
  section with a warning instead of failing the run).
