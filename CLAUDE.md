# Development Context for Claude Code

## Project Structure

```
sheet-snip/
├── src/
│   ├── Code.gs              # Main Apps Script code
│   └── ClipboardDialog.html # UI dialog for clipboard copy
├── .github/
│   └── workflows/
│       └── deploy.yml       # Auto-deploy on push to main
├── .gitignore               # Excludes .clasprc.json (credentials)
├── .claspignore             # Excludes docs from Apps Script push
├── README.md                # Minimal dev setup instructions
├── LICENSE                  # MIT
└── CLAUDE.md                # This file
```

## Apps Script Files

### Code.gs
Main script with:
- `onOpen()` - Creates custom menu in Sheets
- `exportValuesJSON()` - Export cell values as JSON
- `exportFormulasJSON()` - Export formulas as JSON
- `exportBothJSON()` - Export both values and formulas
- Helper functions for building JSON structures

### ClipboardDialog.html
Modal dialog that:
- Displays JSON preview in textarea
- Copies to clipboard (supports modern API + fallback)
- Auto-closes after successful copy

## Deployment Workflow

1. Local changes → `git push`
2. GitHub Actions → `clasp push`
3. Updates live in Apps Script

## Development Commands

```bash
# Push changes manually
clasp push

# Open in Apps Script editor
clasp open

# Check deployment status
gh run list --workflow=deploy.yml
```
