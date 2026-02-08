# Relationship Navigator

An open-source Microsoft Excel Add-in for modeling, visualizing, and analyzing professional relationship networks. Map out the political landscape of your organization with an interactive force-directed graph, real-time Excel sync, and AI-powered event analysis — all from within Excel.

## Why This Exists

Navigating complex interpersonal dynamics in large organizations — especially family-run or highly political environments — is a significant challenge. Critical knowledge about who holds real influence, who is aligned with whom, and who poses a risk often exists only as unwritten institutional memory.

Relationship Navigator gives you a private, structured, intelligent system for mapping and analyzing these dynamics. Your data stays in your Excel file. AI analysis happens on-demand using your own API key.

## Features

### Interactive Network Graph
- **Force-directed layout** powered by D3.js with smooth physics simulation
- **Pan, zoom, and drag** — scroll to zoom, drag background to pan, drag nodes to reposition
- **Visual encoding** — node color = department, edge color = sentiment (green/red/gray/purple), edge thickness = strength, arrowheads = direction
- **Risk indicators** — high-risk individuals get a prominent red border
- **Click to inspect** — click any node or edge to open a details panel

### Real-Time Excel Sync
- Edit data in Excel, see the graph update within ~750ms
- `onChanged` event handlers on all three tables with 500ms debounce
- Bi-directional: AI suggestions can write results back into Excel cells

### AI-Powered Event Analysis
- Send event descriptions to OpenAI, Anthropic, or Google AI for analysis
- AI returns structured classification: event type, impact, severity, and a political analysis summary
- One-click "Apply Suggestions" writes AI results back to the Excel table
- Privacy-first: only the event description text is sent, never names or IDs
- Explicit consent dialog shown before first API call

### Risk Score Engine
- Automatically calculates risk levels for each person based on negative-impact event severity
- Severity sum >= 15 = High, >= 8 = Medium, otherwise Low
- Updates `tbl_People.RiskLevel` only when the computed level differs (prevents infinite sync loops)

### Filtering
- Filter graph by **Department**, **Risk Level**, **Relationship Type**, and **Sentiment**
- Multi-select filters with dynamic population from your data
- Clear all filters in one click

## Data Model

The add-in uses three Excel tables as its data store:

### tbl_People (Nodes)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| ID | Text | Yes | Unique identifier (e.g., `P001`) |
| Name | Text | Yes | Full name |
| Title | Text | No | Job title |
| Department | Text | No | Department/team (determines node color) |
| Influence | Number | No | 1-10 scale |
| RiskLevel | Text | No | `Low` / `Medium` / `High` (auto-calculated) |
| Notes | Text | No | Free-text observations |

### tbl_Relationships (Edges)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| ID | Text | Yes | Unique identifier (e.g., `R001`) |
| PersonA_ID | Text | Yes | FK to tbl_People.ID |
| PersonB_ID | Text | Yes | FK to tbl_People.ID |
| Type | Text | Yes | e.g., `Reports To`, `Mentors`, `Competes With` |
| Strength | Number | No | 1-10 (determines edge thickness) |
| Sentiment | Text | No | `Positive` / `Negative` / `Neutral` / `Complex` (determines edge color) |
| Direction | Text | Yes | `A->B` / `B->A` / `Both` |
| Notes | Text | No | Free-text notes |

### tbl_Events (History)

| Column | Type | Required | Description |
|--------|------|----------|-------------|
| ID | Text | Yes | Unique identifier (e.g., `E001`) |
| Date | Date | Yes | When it happened |
| People_IDs | Text | Yes | Comma-separated IDs (e.g., `P001, P003`) |
| Type | Text | No | `Meeting`, `Conflict`, `Betrayal`, `Alliance`, etc. |
| Description | Text | Yes | Natural-language account of what happened |
| Impact | Text | No | `Positive` / `Negative` / `Neutral` / `Mixed` |
| Severity | Number | No | 1-10 scale |

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Add-in Framework | Office.js (Office Add-ins) |
| Frontend | React 18 + TypeScript 5 |
| Graph Visualization | D3.js v7 (d3-force, d3-zoom, d3-drag) |
| UI Components | Fluent UI React v9 (@fluentui/react-components) |
| Build | Webpack 5 + Babel |
| AI Integration | Direct fetch() to OpenAI / Anthropic / Google APIs |

## Project Structure

```
relationship_agent/
├── manifest.xml                          # Office Add-in manifest (ribbon, permissions)
├── package.json
├── tsconfig.json
├── webpack.config.js
├── babel.config.json
├── assets/                               # Add-in icons (16/32/64/80px)
└── src/
    ├── models/
    │   └── types.ts                      # All TypeScript interfaces, union types, constants
    ├── utils/
    │   ├── debounce.ts                   # Generic debounce (500ms default)
    │   └── colors.ts                     # Department/sentiment/risk color mappings
    ├── services/
    │   ├── DataService.ts                # Office.js read/write bridge, change handlers, settings
    │   ├── LLMService.ts                 # Multi-provider AI client (OpenAI/Anthropic/Google)
    │   └── RiskService.ts                # Risk score calculation engine
    ├── commands/
    │   ├── commands.ts                   # Ribbon command handler
    │   └── commands.html                 # Commands HTML shell
    └── taskpane/
        ├── index.tsx                     # React entry point (Office.onReady + FluentProvider)
        ├── taskpane.html                 # HTML container (loads Office.js CDN)
        ├── taskpane.css                  # Global reset styles
        └── components/
            ├── App.tsx                   # Root component: state, tabs, data loading, sync
            ├── GraphView.tsx             # D3 force-directed graph (SVG)
            ├── DetailsPanel.tsx          # Slide-over panel: person/relationship details + AI
            ├── EventsView.tsx            # Chronological event list with filters
            ├── SettingsView.tsx          # LLM provider/model/API key configuration
            └── FilterPanel.tsx           # Graph filter toolbar (department/risk/type/sentiment)
```

## Getting Started

### Prerequisites

- **Node.js** 18+ and npm
- **Microsoft Excel** (Desktop on Windows/macOS, or Excel on the Web)

### Installation

```bash
# Clone the repository
git clone https://github.com/R1cK-ChaN/relationship_agent.git
cd relationship_agent

# Install dependencies
npm install

# Build for production
npm run build

# Or start the dev server (HTTPS on localhost:3000)
npm run dev-server
```

### Sideloading the Add-in

#### Excel on the Web
1. Open Excel at [office.com](https://www.office.com)
2. Open a workbook (or create a new one)
3. Go to **Home** > **Add-ins** > **More Add-ins** > **Upload My Add-in**
4. Upload `dist/manifest.xml` (after running `npm run build`) or point to `https://localhost:3000/manifest.xml` (when using dev server)

#### Excel Desktop (Windows)
1. Run `npm run dev-server`
2. Open Excel
3. Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**
4. Add a network share containing the manifest, or use the [Office Add-in DevTools](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing) for sideloading

#### Excel Desktop (macOS)
1. Run `npm run dev-server`
2. Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
3. Restart Excel

### First Use

1. Click **Relationship Navigator** on the Home ribbon tab
2. The task pane opens. If no data tables exist, click **Insert Template**
3. Three worksheets are created: People, Relationships, Events (each with a named table)
4. Fill in your data in the Excel tables
5. Switch back to the task pane — the graph renders automatically

## Configuration

### AI Provider Setup

1. Go to the **Settings** tab in the task pane
2. Select a provider: **OpenAI**, **Anthropic**, or **Google**
3. Choose a model from the dropdown
4. Enter your API key and click **Save Settings**

Your API key is stored in the workbook's document settings (persists when saved). It is only used when you explicitly click "Analyze with AI" on an event.

### Supported Models

| Provider | Model | Notes |
|----------|-------|-------|
| OpenAI | GPT-4o | Strong structured output |
| OpenAI | GPT-4o Mini | Cost-effective |
| Anthropic | Claude 3.5 Sonnet | Nuanced analysis |
| Anthropic | Claude 3.5 Haiku | Fast and affordable |
| Google | Gemini 1.5 Flash | Budget-friendly |

## How It Works

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                     Microsoft Excel                         │
│                                                             │
│   ┌─────────────────┐          ┌─────────────────────────┐ │
│   │   Excel Tables   │  Office  │     Task Pane (React)    │ │
│   │                  │  .js API │                          │ │
│   │  tbl_People     ├─────────►│  Graph   Events  Settings│ │
│   │  tbl_Relationships│◄─────────┤  (D3.js) (List)  (Form) │ │
│   │  tbl_Events      │  onChanged│                          │ │
│   └─────────────────┘  + write  └──────────┬──────────────┘ │
│                                             │ fetch()        │
└─────────────────────────────────────────────┼───────────────┘
                                              ▼
                                    ┌──────────────────┐
                                    │  Cloud LLM API   │
                                    │  (Your API Key)  │
                                    └──────────────────┘
```

### Data Flow

1. **Startup**: Task pane opens → reads all 3 tables via Office.js → parses into typed objects → renders graph
2. **Real-time sync**: User edits Excel → `table.onChanged` fires → 500ms debounce → full reload → React re-render → D3 updates
3. **AI analysis**: User clicks "Analyze with AI" → event description sent to LLM → structured JSON response → displayed in panel → optional write-back to Excel
4. **Risk calculation**: After each data load, severity of negative-impact events is summed per person → risk levels updated in Excel where changed

### Key Design Decisions

- **No backend server** — the add-in runs entirely in Excel. Direct fetch() to LLM APIs. No data passes through any intermediary server.
- **Excel as the database** — your workbook is the single source of truth. No external database. Data is portable wherever your .xlsx file goes.
- **Privacy by default** — only event description text is sent to AI (no names, no IDs). Explicit consent required before first API call.

## Development

```bash
# Start dev server with hot reload
npm run dev-server

# Production build
npm run build

# Lint
npm run lint
```

The dev server runs on `https://localhost:3000` with HTTPS (required by Office Add-ins).

### Pre-commit Hook

A pre-commit hook runs `npm run build` before each commit to catch compilation errors. If the build fails, the commit is blocked.

## Privacy & Security

- **Data stays local**: All people, relationships, and event data remain in your Excel file. Nothing is uploaded anywhere.
- **AI is opt-in**: The AI feature only activates when you explicitly click "Analyze with AI" on a specific event.
- **Minimal data sent**: Only the event `Description` text is sent to the AI provider. Names, IDs, and relationship data are never transmitted.
- **Your key, your control**: API keys are stored in the workbook's document settings — not in any cloud service or global config.
- **Consent required**: A privacy disclosure dialog must be acknowledged before the first AI call.

## License

See [LICENSE](LICENSE) for details.
