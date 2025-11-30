# Sheets JSON Export Add-on - Claude Code Setup Plan

**Goal**: Create a functional Google Sheets add-on with automated deployment workflow, shareable with at least one other person.

**Key Constraints**:
- Maximum automation via Claude Code
- Manual intervention only for browser-based auth/config steps
- Private GitHub repo initially (public after v1)
- Minimal documentation overhead

---

## Phase 0: Pre-flight Checks

**Automated Tasks**:
```bash
# Check if clasp is already installed
which clasp

# Check dotfiles for npm global install configuration
cat ~/.dotfiles/.npmrc 2>/dev/null || echo "No custom npm config"
ls -la ~/.dotfiles/ | grep npm

# Verify gh CLI is available
gh --version

# Check existing clasp auth status
clasp login --status 2>/dev/null || echo "Not logged in"
```

**Manual Tasks** (if needed):
1. Enable Apps Script API: https://script.google.com/home/usersettings
2. Run `clasp login` if not authenticated

**Validation**:
- [ ] clasp installed and authenticated
- [ ] gh CLI available
- [ ] No conflicting npm global config

---

## Phase 1: Local Project Setup

**Directory**: `~/Developer/sheets-json-export` (or your preferred location)

**Automated Tasks**:

### 1.1 Create Project Structure
```bash
# Create and initialize directory
mkdir -p ~/Developer/sheets-json-export
cd ~/Developer/sheets-json-export
git init

# Create directory structure
mkdir -p src

# Create configuration files
cat > .gitignore << 'EOF'
# clasp credentials (NEVER commit)
.clasprc.json

# Node
node_modules/

# macOS
.DS_Store

# IDEs
.vscode/
.idea/
EOF

cat > .claspignore << 'EOF'
# Version control
.git/**
.gitignore

# Documentation
README.md
LICENSE
PRIVACY.md
TERMS.md

# Config
.claspignore
.github/**
EOF
```

### 1.2 Copy Source Files
```bash
# Copy from project files to src/
cp /mnt/project/script_v0.1 src/Code.gs
cp /mnt/project/html_v0.1 src/ClipboardDialog.html
```

### 1.3 Create Documentation Files

**README.md** (streamlined version):
```markdown
# Sheets JSON Export

Export Google Sheets ranges to LLM-friendly JSON format.

## Features

- Export cell values as JSON
- Export formulas as JSON  
- Export both values and formulas

## Installation

Install from Google Workspace Marketplace: [Link TBD]

## Usage

1. Select range in Google Sheets
2. **Extensions → Export to JSON → [Choose type]**
3. Copy JSON from dialog
4. Use with Claude or other LLMs

## JSON Format

```json
{
  "metadata": {
    "sheetName": "Sheet1",
    "range": "A1:C3",
    "dimensions": {"rows": 3, "columns": 3},
    "exportedAt": "2025-01-15T10:30:00.000Z",
    "exportType": "values"
  },
  "data": [
    ["Header1", "Header2", "Header3"],
    [1, 2, 3],
    [4, 5, 6]
  ]
}
```

## Development

```bash
git clone https://github.com/jshew/sheets-json-export.git
cd sheets-json-export
clasp login
clasp push
clasp open  # Test in browser
```

## License

MIT License - Copyright (c) 2025 Joshua Shew

Built with Claude.
```

**LICENSE** (MIT):
```
MIT License

Copyright (c) 2025 Joshua Shew

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

### 1.4 Create GitHub Repository
```bash
# Create private repo
gh repo create sheets-json-export \
  --private \
  --description "Export Google Sheets selections to LLM-friendly JSON format" \
  --source=. \
  --remote=origin

# Initial commit
git add .
git commit -m "Initial commit: Google Sheets JSON export add-on

- Export values, formulas, or both as JSON
- Clipboard copy dialog
- Metadata includes sheet name, range, dimensions, timestamp"

# Push to GitHub
git branch -M main
git push -u origin main
```

**Validation**:
- [ ] Local git repo initialized
- [ ] Private GitHub repo created
- [ ] Initial code committed and pushed

---

## Phase 2: Apps Script Deployment

**Automated Tasks**:

### 2.1 Create Apps Script Project
```bash
cd ~/Developer/sheets-json-export

# Create standalone Apps Script project
clasp create \
  --type standalone \
  --title "Sheets JSON Export" \
  --rootDir ./src

# This creates src/.clasp.json with script ID
```

### 2.2 Push Code to Apps Script
```bash
clasp push

# Verify deployment
clasp open  # Opens in browser for manual testing
```

**Manual Tasks**:
1. In opened Apps Script editor, create a test spreadsheet
2. Run `onOpen()` function once manually (grants initial permissions)
3. Test all three export functions with sample data
4. Verify JSON output format and clipboard copy

### 2.3 Commit clasp Configuration
```bash
# Add .clasp.json to repo (contains script ID, not credentials)
git add src/.clasp.json
git commit -m "Add Apps Script project configuration"
git push
```

**Validation**:
- [ ] Apps Script project created
- [ ] Code pushed successfully
- [ ] Manual test confirms all features work
- [ ] Script ID committed to repo

---

## Phase 3: Sharing Setup (Private Testing)

**Goal**: Enable sharing with at least one other person for testing.

### 3.1 Create Test Deployment
```bash
clasp open
```

**Manual Tasks** (in browser):
1. Click **Deploy → New deployment**
2. Select type: **Add-on**
3. Description: "v0.1.0-beta - Private testing"
4. Access: Choose **Only people I trust** or **Anyone with link**
5. Click **Deploy**
6. Copy deployment ID and test URL

### 3.2 Document Sharing Process

Create **TESTING.md**:
```markdown
# Testing Instructions

## For Testers

1. Install test version: [Add link from deployment]
2. Open any Google Sheet
3. Grant permissions when prompted
4. Test menu: **Extensions → Export to JSON**

## Test Cases

- [ ] Export values from simple range (A1:C3)
- [ ] Export formulas from range with =SUM(), =AVERAGE()
- [ ] Export both values and formulas
- [ ] Verify clipboard copy works
- [ ] Test with empty cells
- [ ] Test with merged cells
- [ ] Test with different data types (numbers, text, dates)

## Report Issues

Open issue: https://github.com/jshew/sheets-json-export/issues
```

```bash
git add TESTING.md
git commit -m "Add testing instructions"
git push
```

**Validation**:
- [ ] Test deployment created
- [ ] Shared with at least one tester
- [ ] Tester can install and use add-on

---

## Phase 4: GitHub Actions Workflow

**Goal**: Automated deployment on push to main.

### 4.1 Extract clasp Credentials
```bash
# View current clasp auth
cat ~/.clasprc.json

# Copy the entire JSON content for GitHub secret
```

### 4.2 Add GitHub Secret
```bash
# Add secret via gh CLI
gh secret set CLASPRC < ~/.clasprc.json
```

### 4.3 Create Workflow File

```bash
mkdir -p .github/workflows

cat > .github/workflows/deploy.yml << 'EOF'
name: Deploy to Apps Script

on:
  push:
    branches: [main]
  workflow_dispatch:

jobs:
  deploy:
    runs-on: ubuntu-latest
    
    steps:
      - name: Checkout code
        uses: actions/checkout@v4
      
      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'
      
      - name: Install clasp
        run: npm install -g @google/clasp
      
      - name: Write clasp credentials
        run: echo '${{ secrets.CLASPRC }}' > ~/.clasprc.json
      
      - name: Push to Apps Script
        run: clasp push
        working-directory: .
      
      - name: Deployment summary
        run: |
          echo "✅ Successfully deployed to Apps Script" >> $GITHUB_STEP_SUMMARY
          echo "Script ID: $(cat src/.clasp.json | grep scriptId)" >> $GITHUB_STEP_SUMMARY
EOF

git add .github/workflows/deploy.yml
git commit -m "Add automated deployment workflow"
git push
```

### 4.4 Verify Workflow
```bash
# Check workflow status
gh run list --workflow=deploy.yml

# Or visit Actions tab in GitHub
gh browse
```

**Validation**:
- [ ] Workflow file created
- [ ] CLASPRC secret added
- [ ] Push triggers successful deployment
- [ ] Can manually trigger workflow

---

## Phase 5: Documentation for Distribution

### 5.1 Research Privacy Policy & ToS

**Research Tasks**:
```bash
# Search for minimal privacy policy examples for hobby projects
# Look for patterns in similar Google Workspace add-ons
# Determine: In-repo markdown vs hosted page vs GitHub Pages
```

**Questions to Answer**:
- Standard approach for small/hobby add-ons?
- Minimum viable privacy policy content?
- ToS requirements vs just having MIT license?
- GitHub Pages vs simple markdown file?

### 5.2 Create Privacy Policy (Template)

**PRIVACY.md**:
```markdown
# Privacy Policy for Sheets JSON Export

**Last Updated**: [Date]

## Data Collection

This add-on does NOT collect, store, or transmit any user data. All operations occur locally within your Google Sheets session.

## Data Processing

- The add-on only accesses the currently selected range in your active spreadsheet
- JSON export is generated client-side in your browser
- Data is only copied to your clipboard - never sent to external servers
- No analytics, tracking, or telemetry

## Permissions Required

- `spreadsheets.currentonly`: Read selected range data
- `script.container.ui`: Display export dialog

## Third-Party Access

This add-on does not share data with any third parties.

## Contact

Questions: joshua.t.shew@gmail.com
```

### 5.3 Create Terms (Optional - Research Result)

If required, create minimal **TERMS.md** based on research findings.

### 5.4 Determine Visibility Strategy

**Research**: Public vs Private add-on distribution

**Private Benefits**:
- No Google review process
- Share via deployment link
- Faster iteration
- Good for beta testing

**Public Benefits**:
- Discoverable in Marketplace
- Automatic updates for all users
- More credibility
- Professional distribution

**Recommendation**: Start private, go public after testing validates value.

---

## Phase 6: Prepare for Public Release (Future)

**Prerequisites**:
- [ ] Tested by multiple users
- [ ] Privacy policy finalized
- [ ] Icons created (128x128px, 32x32px PNG)
- [ ] Screenshots prepared (1280x800px)
- [ ] Support process defined (GitHub Issues)

**Tasks**:
1. Create Google Cloud Project (separate from default)
2. Configure OAuth consent screen
3. Enable Workspace Marketplace SDK
4. Submit for review
5. Update repo visibility to public
6. Update README with Marketplace link

**Defer to separate phase** - not part of initial sprint.

---

## Success Criteria

Initial sprint is complete when:

✅ **Core Functionality**:
- Add-on exports values, formulas, and combined JSON
- Clipboard copy works reliably
- Proper metadata in output

✅ **Deployment**:
- Private deployment shareable via link
- At least one external tester has successfully used it
- GitHub Actions auto-deploys on push to main

✅ **Documentation**:
- README explains installation and usage
- Testing instructions available
- Privacy policy documented

✅ **Development Workflow**:
- Can make changes locally
- `git push` triggers automatic deployment
- Can test changes in Apps Script editor

---

## Quick Reference Commands

```bash
# Development workflow
git checkout -b feature-name
# Make changes to src/Code.gs or src/ClipboardDialog.html
git add .
git commit -m "Description"
git push origin feature-name
gh pr create
# Merge PR on GitHub
git checkout main && git pull  # Triggers auto-deploy

# Manual deployment
clasp push
clasp open

# Check deployment status
gh run list

# Share test deployment
# Get link from Apps Script editor: Deploy → Manage deployments → Copy link
```

---

## Troubleshooting

### clasp push fails
```bash
# Check auth
clasp login --status

# Re-authenticate
clasp logout
clasp login

# Verify script ID
cat src/.clasp.json
```

### GitHub Actions fails
```bash
# Check logs
gh run view --log

# Verify secret
gh secret list

# Re-add secret if needed
gh secret set CLASPRC < ~/.clasprc.json
```

### Can't share deployment
- Ensure deployment type is "Add-on" not "Web app"
- Check access settings (Anyone with link vs restricted)
- Verify OAuth scopes are configured

---

## Next Steps After Initial Sprint

1. Gather feedback from initial testers
2. Iterate on UX and JSON format
3. Research privacy policy best practices
4. Create professional icons
5. Decide on public vs private distribution
6. Optional: Add changelog automation with release notes
7. Optional: Add basic tests in Apps Script

---

## File Structure Reference

```
sheets-json-export/
├── .git/
├── .github/
│   └── workflows/
│       └── deploy.yml          # Auto-deployment
├── src/
│   ├── Code.gs                 # Main script
│   ├── ClipboardDialog.html    # UI dialog
│   └── .clasp.json             # Script ID (committed)
├── .gitignore                  # Excludes .clasprc.json
├── .claspignore                # Excludes docs from push
├── README.md                   # Main documentation
├── LICENSE                     # MIT License
├── PRIVACY.md                  # Privacy policy
├── TESTING.md                  # Test instructions
└── (future) TERMS.md           # If required
```

**Not in repo**:
- `~/.clasprc.json` - Local auth only, stored in GitHub secret
