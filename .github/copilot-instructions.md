# GitHub Copilot Instructions for boothcsv

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

### Testing
- Manual testing via `tests/phase6-tests.html` and `tests/custom-labels-store-test.html`
- No automated test suite - relies on manual QA with sample CSV files in `sample/`

### Debugging Tips
- Enable debug mode: `?debug=1` in URL
- Check `window.orderRepository.cache` for in-memory state
- Inspect IndexedDB "BoothCSVStorage" v9 for persistent data
- Look for console logs with category prefixes like `[csv]`, `[repo]`, `[image]`

## Project-Specific Conventions

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

## Common Tasks
- **Add new setting**: Update `StorageManager.KEYS`, `settingsCache`, and UI binding in `registerSettingChangeHandlers()`
- **Modify label layout**: Adjust `CONSTANTS.LABEL.TOTAL_LABELS_PER_SHEET` and related CSS
- **Database changes**: Increment `UnifiedDatabase.version` and add migration logic in `createObjectStores()`
