# GitHub Copilot Instructions for boothcsv

## 日本語
- 日本語で回答してください。

## Project Overview
This is a **client-side web application** for processing BOOTH (Japanese marketplace) shipping CSV files and generating printable labels/order details. It's built as a pure JavaScript application without external dependencies, running entirely in the browser with no server required.

## Core Architecture

### File Structure & Responsibilities
- **`boothcsv.html`** - Main HTML entry point with fixed header UI and sidebar
- **`boothcsv.js`** (2700+ lines) - Main application logic, CSV processing, label generation
- **`storage.js`** - IndexedDB wrapper (`UnifiedDatabase` class) for persistent storage (v9 schema)
- **`order-repository.js`** - Order data management with in-memory cache and database sync
- **`custom-labels.js`** - Custom label creation and editing functionality
- **`custom-labels-font.js`** - Font management for custom labels
- **External libs**: `papaparse.min.js` (CSV parsing), `jsQR.js` (QR code reading)

### Data Flow Pattern
1. **CSV Upload** → `autoProcessCSV()` → Parse with PapaParse → Store in IndexedDB via `OrderRepository`
2. **Label Generation** → `generateLabels()` → Create DOM elements with A4 44-label layout
3. **Print Workflow** → Browser print → `updateSkipCount()` → Update cache & database for next session

### Key Architectural Decisions
- **No Server Required**: Everything runs client-side for privacy (sensitive shipping data)
- **IndexedDB Storage**: Persistent storage with versioned schema (currently v9)
- **Global State Management**: Uses `window.settingsCache` and `window.orderRepository` for cross-component state
- **Label Layout**: Hardcoded for A4 44-label sheets (Japanese standard)

## Critical Patterns

### Settings Cache Synchronization
```javascript
// Always update both storage AND cache when changing settings
await StorageManager.set(StorageManager.KEYS.LABEL_SKIP, newValue);
settingsCache.labelskip = newValue; // ⚠️ REQUIRED: Update cache too
```

### Order Repository Pattern
```javascript
// Repository manages both cache and database
const repo = window.orderRepository;
await repo.bulkUpsert(csvRows); // Updates cache + IndexedDB
const order = repo.get(orderNumber); // Cache-first lookup
```

### Print Lifecycle
```javascript
// Print completion updates skip count for next session
window.print() → user confirms → updateSkipCount() → recalculate label positions
```

## Development Workflows

### Local Development
- Open `boothcsv.html` directly in Chrome (file:// protocol)
- Add `?debug=1` to URL to enable debug logging
- Use browser DevTools → Application → IndexedDB to inspect data

### App Mode / Environment Branching
- This project currently has 3 meaningful runtime modes: `extension`, `web`, `file`
- **Recommended rule**: do not scatter direct checks of `window.location.protocol`, `chrome.runtime`, or ad-hoc `hidden` toggles across unrelated features
- Centralize environment resolution in one helper such as `getAppMode()` / `resolveAppMode()` and use that result everywhere
- Prefer applying the resolved mode to the document root as a single source of truth, e.g. `document.body.dataset.appMode = 'extension|web|file'`
- UI-only differences should be driven from that single app mode state, ideally by one updater function plus CSS selectors based on `body[data-app-mode="..."]`
- Feature availability should be separated from UI mode. Example: "show extension-like UI for debugging" is different from "actual extension bridge is callable"

### Forced Mode For Debugging
- Keep a **forced app mode override** for debugging, especially when verifying UI with MCP chrome-devtools on localhost
- Recommended override: query parameter such as `?appMode=extension`, `?appMode=web`, `?appMode=file`
- This override is also important for taking screenshots used in help pages, usage guides, and documentation via MCP chrome-devtools
- Precedence should be: explicit query override → actual runtime detection → fallback default
- Forced `extension` mode is for **visual/UI debugging** and layout verification. It must not assume that Chrome extension APIs are really available
- Real extension-only actions must still check capability separately, e.g. `canUseExtensionBridge()` or `isRealChromeExtensionApp()`
- If a mode override is present, it should be easy to inspect in DevTools from `document.body.dataset.appMode`
- Do not support a separate legacy override such as `?extension=1`; use `appMode=...` only

### Testing
- Manual testing via `tests/phase6-tests.html` and `tests/custom-labels-store-test.html`
- No automated test suite - relies on manual QA with sample CSV files in `sample/`

### Debugging Tips
- Enable debug mode: `?debug=1` in URL
- For UI debugging, allow forced mode URLs such as `?debug=1&appMode=extension`
- Use forced `extension` mode when checking extension-only layout in localhost via MCP chrome-devtools
- Use the same forced mode URLs when taking screenshots for help content so the captured UI is reproducible later
- Check `window.orderRepository.cache` for in-memory state
- Inspect IndexedDB "BoothCSVStorage" v9 for persistent data
- Look for console logs with category prefixes like `[csv]`, `[repo]`, `[image]`

## Project-Specific Conventions

### Environment Branching Conventions
- Use **one** environment/app-mode resolver shared by the app, instead of duplicating `isChromeExtensionApp()` logic in multiple files
- Use **one** UI visibility updater for mode-based show/hide rules, rather than per-feature ad-hoc DOM toggles where possible
- Prefer CSS branching from root mode state for pure presentation differences
- Prefer capability checks for behavior differences. Example:
	- app mode decides whether extension-only controls are visible
	- capability check decides whether clicking those controls is actually allowed
- Avoid using URL parameters that are not consumed by the app. If extension launch passes a mode hint, it should map to the same unified app-mode mechanism
- Prefer `appMode=...` as the only documented and supported debug/screenshot override
- When adding new extension-only features, first decide whether the requirement is:
	- visual difference only
	- behavior difference only
	- both visual and behavior difference
	Then implement through the shared mode/capability system instead of introducing a new standalone branch

### Button Styling Conventions
- Prefer composing new buttons from existing shared classes instead of creating one-off button styles
- Default composition should be `btn-base` + one size class + one color class
- For header/main actions, prefer a reusable helper such as `btn-header` on top of `btn-base` rather than a brand new standalone button class
- Reuse existing color variants first (`btn-primary`, `btn-success`, `btn-danger`, `btn-info`, `btn-warning`, `btn-download`, `btn-print-accent`) before introducing a new color token
- New button-specific CSS should be limited to layout needs that cannot be expressed through the shared button classes
- Avoid gradients, custom shadows, or duplicated hover/disabled rules unless a new visual language is intentionally required for the whole app

### Error Handling
- Non-blocking: Most errors are logged but don't halt execution
- User feedback via `alert()` for critical errors
- Storage operations have fallback mechanisms

### Image Handling
- Global images stored as ArrayBuffer in IndexedDB settings
- Individual order images stored per order in repository
- Uses Blob URLs for display to avoid memory leaks

### Custom Label System
- Independent storage with UNIX timestamp keys
- Collision avoidance with monotonic counter
- Rich text editing with font management

## Integration Points

### Browser APIs
- **IndexedDB**: Primary storage (schema migrations handled automatically)
- **File API**: CSV file reading and image handling
- **Print API**: `window.print()` with CSS `@media print` rules
- **Canvas API**: QR code processing with jsQR library

### External Dependencies (CDN)
- Paper.css for A4 print layouts
- No build process - dependencies loaded via CDN or vendored

## Key Gotchas
- **Cache Consistency**: Always update both `settingsCache` and database when modifying settings
- **Label Skip Logic**: Complex calculation involving 44-label sheets and print completion feedback
- **QR Code Processing**: Expects specific format from BOOTH QR codes (3 parts separated by specific delimiters)
- **Print CSS**: Layout relies heavily on CSS Grid and print media queries - test in print preview
- **Mode vs capability**: "extension-looking UI" and "real Chrome extension APIs are callable" are not the same thing; keep them separate when implementing debug overrides

## Common Tasks
- **Add new setting**: Update `StorageManager.KEYS`, `settingsCache`, and UI binding in `registerSettingChangeHandlers()`
- **Modify label layout**: Adjust `CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET` and related CSS
- **Database changes**: Increment `UnifiedDatabase.version` and add migration logic in `createObjectStores()`
