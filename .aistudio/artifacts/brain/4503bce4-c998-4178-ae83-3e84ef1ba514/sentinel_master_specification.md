# Sentinel Studio™ — Master Feature Specification & User Guide

**Sentinel Studio™** is an **Enterprise Capability Hub (ERI × EPEx)** — integrating **Enterprise Risk & Intelligence (ERI)** with **Enterprise Performance Excellence (EPEx)**. Architected by **Patrick Jay F. Tanap, CPA, CICA**, Sentinel Studio is an enterprise corporate intelligence workstation, local analytics hub, and document generation suite.

It is engineered for statutory compliance under **RA 10173 (Philippine Data Privacy Act)** with **100% client-side, on-device local WebGPU / WASM execution**, operating with zero paid AI subscriptions, zero external API keys, and zero cloud data egress.

---

## Part 1: Architectural & Security Foundations

### 1. 100% On-Device Local AI Execution (WebGPU / WASM)
* **WebLLM & LiteRT Execution Engines**: AI inference runs entirely inside browser VRAM / RAM using local open-weight models (`Qwen2.5-1.5B-Instruct-q4f16_1-MLC` via WebGPU or `.litertlm` / `.bin` Gemma 4 local binaries via LiteRT-LM WASM).
* **Zero Subscription & Zero API Cost**: Operates completely independently of paid SaaS subscriptions, cloud LLM providers, or third-party API keys.
* **RA 10173 Statutory Security**: Prompts, uploaded spreadsheets/documents, persistent memory rules, custom agents, custom personas, custom capabilities, and generated outputs never leave your physical workstation.

### 2. External File-Backed Storage Workspace & Hot-Directory Auto-Sync
* **Native OS Directory Mapping (File System Access API)**: Maps directly to a local disk folder or workspace file (`sentinel_vault_latest.json`) using browser File System Access API (`showDirectoryPicker`, `showOpenFilePicker`).
* **Real-Time Shadow Mirroring**: Automatically serializes and mirrors chat sessions, persistent memory rules, custom agents, custom personas, and enterprise capabilities in real-time to your mapped local drive in the background.
* **Air-Gapped Local Disk Persistence**: Bypasses browser localStorage limitations and provides true local file-system workspace persistence without cloud synchronization.

### 3. Binary Model Storage & Fast Caching
* **OPFS & IndexedDB Model Vault**: Local model binaries (`.litertlm`, `.bin`, `.task`, `.tflite`) are cached directly inside the browser's **Origin Private File System (OPFS)** and IndexedDB memory stores for instant, offline auto-restoration upon browser refresh.

### 4. Telegram Mobile Activation & Hardware Verification
* **SHA-256 Hardware Fingerprinting**: Computes a unique hardware fingerprint using system parameters, GPU renderer strings, and WebGL canvas signatures.
* **Telegram Approval Gateway**: Encrypts and dispatches activation requests over an encrypted endpoint to trigger a mobile approval prompt on the developer's phone, automatically unlocking the browser workstation upon approval.
* **Offline Key Restoration**: Supports offline `.lic` key file restoration for air-gapped workstations and secure offline environments.

---

## Part 2: Detailed Feature & Function Specification

### 1. Gem Personas Suite: Pre-Built & Custom Persona Creation
* **Built-in Expert Personas**:
  * **💼 General Assistant**: Versatile companion for executive writing, summarization, strategic brainstorming, and document reviews.
  * **🛡️ Intel Audit Lead**: Senior Internal Controls & Enterprise Risk Auditor that structures findings strictly into the **5-part audit observation format**:
    1. *Condition* (What was observed)
    2. *Criteria* (The governing policy, DOA threshold, or statutory requirement)
    3. *Cause* (Root-cause vulnerability)
    4. *Impact* (Financial, operational, or regulatory risk exposure)
    5. *Actionable Recommendation* (Remediation plan)
  * **🧠 Deep Reasoner (R1)**: Step-by-step logic analyzer that utilizes reasoning blocks (`<think>...</think>`) before outputting conclusions. Includes high-contrast light and dark mode styling for complete legibility.
  * **📊 Excel & Data Specialist**: Financial analyst for variance detection, anomaly scanning, DOA threshold split PO detection, margin shift analysis, and markdown table synthesis.
  * **📝 Executive Writer & Scribe**: Advisor for board briefings, executive memos, and committee deliverables.

* **✨ Ability to Create & Edit Custom Personas**:
  * **Custom Persona Builder**: Click **+ Create Persona** or click Edit (✏️) on any persona in the left sidebar or header dropdown.
  * **Custom Attributes**: Persona Name, Custom Emoji Avatar, Role Description, System Prompt Instructions, Category, and **Custom Starter Prompt Pills (`✦`)**.
  * **Storage & Availability**: Custom personas are saved to persistent storage and immediately integrated into the Personas list, header selector, and Agent Studio step binding dropdowns.

* **Mid-Session Persona Switching**: Switch active personas at any turn during a conversation without losing chat history or context.
* **Welcome / Greeting Screen (`#emptyState`)**: On app launch, model reload, new session creation, or switching personas on an empty chat, the welcome screen dynamically displays the active persona's avatar, title, description, and starter prompt pills. System auto-restore notices never hide the welcome screen when the chat has no conversation messages.

---

### 2. Session History & Group Management
* **Group Management ("Create Group")**: Click **Create Group** beside the **SESSION HISTORY** header to organize chat sessions into color-coded, icon-customized history groups.
* **Drag-and-Drop Session Grouping**: Drag any chat session card and drop it directly onto a group card to organize conversations into functional projects.
* **Folder Actions**: Expand/collapse groups, add new chats directly inside a group, edit group icon/color properties, move sessions between groups, or delete groups.
* **Live Session Search**: Instant client-side search filtering across session titles, prompt text, and historical responses.

---

### 3. Enterprise Capabilities Architecture: Custom Capability Creation & Invocation
* **✨ Ability to Create & Manage Custom Capabilities**:
  * **Custom Capability Builder**: Navigate to **Enterprise Settings ⚙️ ➔ Capabilities** tab and click **+ Create New Capability**.
  * **Capability Attributes**: Custom Name, Slash Command (`/#command`), ERI/EPEx Category, Description, Required Parameters, and Execution System Prompt Template.
  * **Integration**: Custom capabilities automatically register for chat slash autocomplete (`/#`), Agent Studio step bindings, and WebGPU memory guard scanning.

* **Chat Slash Invocation**: Type `/#` in the prompt input box to open the autocomplete popup and invoke any custom or built-in capability during chat turns.
* **WebGPU Memory Guard**: Automatically scans prompts for active capabilities and limits simultaneous turn execution to the top 2 capabilities to maintain high WebGPU inference speeds.
* **High-Contrast Theme Dynamics**: All capability badges, command chips (`/#command`), banner elements, and input elements are styled with high-contrast text and background colors in both Light Mode and Dark Mode.

---

### 4. Agent Studio: Pre-Built & Custom Multi-Persona Agent Creation
Automate complex multi-step operational routines with pre-built or custom AI agents:

* **Built-In Service Catalog**:
  * **DOA Split PO & Approval Threshold Auditor**: Scans AP transaction logs, identifies split purchase orders processed within 48 hours totaling below the PHP 500,000 threshold (DOA Section 4.2), and generates a Word audit memo.
  * **Process Defect & SLA Variance Reconciliation Matrix**: Filters processes with cycle time turnaround variance, computes defect rates, and formats findings directly into a 5-column Excel reconciliation matrix.
  * **Management Committee Executive Summary Deck**: Synthesizes cross-functional findings into board-level strategic takeaways and structures a 3-slide executive carousel deck.

* **✨ Ability to Build Custom Multi-Persona Agents**:
  * Click **+ Build New Agent** in Agent Studio to open the 4-Step Guided Agent Builder:
    * **Step 1 (Mandate & Identity)**: Define Agent Name, Icon, ERI/EPEx Category, Primary Persona Anchor, and Mandate.
    * **Step 2 (Runtime Target & Schema)**: Select deliverable target (*Word*, *Excel Grid*, or *Slide Deck*) and set execution temperature (0.0 for deterministic audit precision up to 1.0).
    * **Step 3 (Multi-Persona Routine & Multiselect Capabilities)**:
      * **🎭 Assigned Specialist Persona**: Assign different custom or built-in Gem Personas (e.g. Step 1: *Intel Audit Lead*, Step 2: *Excel Specialist*, Step 3: *Executive Scribe*) to construct multi-specialist pipelines.
      * **⚡ Multiselect Capabilities Dropdown**: Assign multiple custom or built-in Enterprise Capabilities (`/#commands`) to each routine step via an interactive multiselect dropdown with checkbox controls.
    * **Step 4 (Hyperparameters & Safeguards)**: Configure deep reasoning traces (`<think>`), auto-canvas switching, and **☑️ Auto-Trigger File Download (.docx / .xlsx / .pptx)**.

* **Responsive Container Discipline**: The "Run Agent" button on agent cards inside the Work Desk Agent Studio remains neatly anchored within its container during screen/window resizing.

---

### 5. Memory Management & Sentinel Memory Vault
Complete working memory management across sessions and context budget optimization:

* **Sentinel Working Memory Ledger**: Persistent rule vault stored in IndexedDB that retains domain context, corporate governance guidelines, DOA thresholds, corrections, and user preferences across chat sessions.
* **Inline Memory Rule Detection**: Automatically scans user prompts for trigger phrases like *"remember that..."*, *"our SLA limit is..."*, or *"always format as..."*, presenting an instant 1-click **Save Memory Rule** banner to add rules directly to the persistent ledger.
* **Memory Hub Management (Settings ⚙️ ➔ Memory Hub)**: Interactive rule management interface where users can view, filter by category (*Audit Baseline*, *DOA Threshold*, *Reporting Format*, *General Context*), pin/favorite rules, edit rule content inline, or delete obsolete rules.
* **Configurable Memory Profiles**:
  * **Small (Eco)**: 20-turn window, max 400 Excel rows, 35,000 doc chars. (8 GB RAM laptops).
  * **Medium (Balanced - Default)**: 35-turn window, max 1,000 Excel rows, 85,000 doc chars.
  * **Large (Deep)**: 50-turn window, max 2,000 Excel rows, 150,000 doc chars.
  * **No Limit**: Unbounded session memory window up to 1,000,000 chars. (32 GB RAM + Discrete GPU workstations).
* **Context Budgeting**: Injected persistent rules are capped at 2,800 characters (~700 tokens) per prompt to reserve remaining context for local model generation.
* **Self-Healing Token Overflow Recovery**: If deep reasoning traces or file context exceed VRAM allocations, the engine automatically catches the token overflow, resets the KV cache, and runs a self-healing retry prompt to output results safely without crashing.

---

### 6. Interactive Work Desk Suite
A split-screen productivity canvas featuring four primary workspace views plus Agent Studio and Enterprise Settings:

```
+-----------------------------------------------------------------------------------+
|                        SENTINEL STUDIO™ WORK DESK                                 |
+-------------------+-------------------+--------------------+----------------------+
|  📝 Notepad       |  📄 Word Canvas   |  📊 Excel Grid     |  📽️ Slide Carousel  |
|  Raw Markdown &   |  Multi-Page A4    |  Tabular Matrix &  |  16:9 Deck Builder & |
|  Inline AI        |  Formal Reports   |  Sheet Tabs        |  Speaker Notes       |
+-------------------+-------------------+--------------------+----------------------+
```

#### A. 📝 Executive Notepad & Inline AI Rewriter
* **Rich Markdown Editor**: Write, edit, format, and view rendered document drafts.
* **Floating Inline AI Bubble**: Highlight any sentence or paragraph in the Notepad to trigger the inline AI rewrite bubble. Enter directives like *"Convert into formal audit finding"*, *"Summarize as executive bullet points"*, or *"Add root-cause analysis"*.
* **Multiple Revision Options**: Generates AI revision options with a 1-click **✓ Apply to Notepad** button to replace highlighted text in-place.

#### B. 📄 Executive Word Report Canvas
* Renders content on physical letter/A4 styled paper pages with header banners, executive classification markers, auto-numbered page counts, and formatted tables.

#### C. 📊 Excel Grid Reconciliation View
* Parses markdown tables into an interactive spreadsheet grid.
* Supports multiple worksheet tabs (`### Sheet: Name`), lettered column headers (A, B, C...), row numbers, and live row/column counters.

#### D. 📽️ 16:9 Presentation Deck Carousel
* Converts AI outputs or slide markdown into 16:9 presentation slide decks featuring:
  * **Hero Cover Slide**: Title, category badge, presenter name, title, department, and classification.
  * **KPI Stat Tiles**: Renders metric cards with value highlights.
  * **Pillar Cards**: Grouped strategic takeaway cards.
  * **FMEA Risk Grid**: Failure mode, root cause, mitigation, and RPN score tables.
  * **Split Comparison**: Side-by-side As-Is / To-Be analysis.
  * **Roadmap Timeline**: Multi-phase execution stages.
  * **Closing Sign-Off Slide**: Executive endorsement card.
  * **Speaker Notes Panel**: Displays presenter talking points below the slide.

---

### 7. Client-Side Document Extractor Suite
Upload files directly into the prompt box for instant extraction and analysis:
* **Word Documents (`.docx`, `.doc`)**: Extracted locally via Mammoth.js.
* **Excel Spreadsheets (`.xlsx`, `.xls`)**: Parsed via SheetJS (XLSX) into structured markdown tables.
* **PowerPoint Decks (`.pptx`)**: XML DOM parser extracts slide text.
* **PDF Files (`.pdf`)**: Extracted page-by-page via PDF.js.
* **Images (`.png`, `.jpg`)**: Vision / OCR ingestion.

---

### 8. Multi-Format Client Exporters
Export work desk documents into native offline files:
* **Word (`.doc`)**: Formatted Word document with custom typography, headers, and styled tables.
* **Excel (`.xlsx`)**: Standard workbook spreadsheets with auto-fitted column widths and tabbed sheets.
* **PowerPoint (`.pptx`)**: Formatted 16:9 widescreen PowerPoint deck generated via PptxGenJS.
* **Markdown (`.md`)**: Plain-text document export.

---

### 9. 4-Tier Enterprise Disaster Recovery Backup System
* **Tier 1 (OPFS Shadow Ledger)**: Maintains a ring-buffer transaction log of the last 150 state snapshots in `journal.json`. Automatically recovers your work if browser cache is cleared.
* **Tier 2 (Hot-Directory Auto-Sync)**: Maps directly to a local OS directory or file (`sentinel_vault_latest.json`) using the browser File System Access API for background auto-syncing.
* **Tier 3 (Modular Component Exports)**: Download targeted backup files (`.mem`, `.agents`, `.chats`, `.json`).
* **Tier 4 (Universal Restore)**: Drag-and-drop restore engine that identifies backup files and intelligently merges or restores your workspace.

---

### 10. Interface Design & Theme Engine
* **Chrome 3-Way Theme Engine**: Toggle between **Light**, **Dark**, or **System/Chrome Default** appearance modes.
* **Accessibility Guarantee**: Features scoped CSS overrides (`html:not(.dark)`) ensuring high text contrast for badges, banners, slash commands (`/#command`), and model indicator badges in Light Mode, while preserving exact dark theme styling in Dark Mode.

---

## Part 3: Step-by-Step "How to Use" Guide

### Step 1: Device Verification & Telegram Mobile Activation
1. On initial launch, if the Device Verification screen appears, enter your **Full Name**, **Department**, and **Justification Note**.
2. Review your computed **SHA-256 Hardware Fingerprint** and click **Request Mobile Approval (Telegram)**.
3. Once approved on the mobile Telegram gateway, the application automatically unlocks and saves your license. (Or click *Authorize as Patrick Jay F. Tanap* / upload a `.lic` key for offline unlock).

### Step 2: Configure External File-Backed Storage Workspace & Auto-Sync
1. Open **Settings ⚙️ ➔ Backup Suite / Auto-Sync**.
2. Click **Map Directory / Connect Auto-Sync** and select a folder on your computer (or select an existing `sentinel_vault_latest.json` file).
3. Sentinel Studio connects via the browser File System Access API and automatically shadow-mirrors all sessions, memory rules, custom personas, custom agents, and custom capabilities directly to your local drive.

### Step 3: First-Time Local Model Loading (WebGPU / WASM)
1. In the chat box bar beside the memory profile selector, check the **Model Indicator Badge**.
2. If unloaded, click **Load Local Model** or drop a local model file (`.litertlm`, `.bin`, `.task`, or `.tflite`).
3. Wait a few seconds while the browser allocates WebGPU shaders and tensor maps.
4. Once the green **Model Active** badge appears, the AI engine is ready for offline execution.

### Step 4: Create & Switch Gem Personas
1. **Selecting Built-in Personas**: Click the **Active Gem Persona** pill at the top header or in the left sidebar (*Intel Audit Lead*, *Excel Specialist*, *Deep Reasoner*, *Executive Writer*, *General Assistant*).
2. **Creating Custom Personas**:
   * In the left sidebar Personas list, click **+ Create Persona** (or click ✏️ to edit an existing persona).
   * Set Persona Name, Emoji Avatar, Role Instructions/System Prompt, Category, and Custom Starter Prompt Pills (`✦`).
   * Click **Save Persona**. Your custom persona is instantly available for chat turns and Agent Studio step bindings.
3. **Mid-Session Persona Switching**: Switch personas at any turn during a conversation—full chat history remains preserved.

### Step 5: Create & Use Custom Enterprise Capabilities
1. Navigate to **Enterprise Settings ⚙️ ➔ Capabilities**.
2. Click **+ Create New Capability**.
3. Define Name, Slash Command (e.g. `/#audit_check`), Category, Description, Parameters, and Execution Prompt Template.
4. Click **Save Capability**.
5. **Chat Invocation**: Type `/#` in the prompt bar during chat, select your capability from the autocomplete popup, and send your prompt.

### Step 6: Build & Run Custom Multi-Persona Agents
1. Open Work Desk and switch view to **⚡ Agent Studio**.
2. Click **+ Build New Agent**:
   * **Step 1**: Set Agent Name, Icon, Category, Primary Persona, and Mandate.
   * **Step 2**: Select Target Deliverable format (*Word*, *Excel Grid*, or *Slide Deck*).
   * **Step 3 (Multi-Persona Routine)**: Add execution steps.
     * Use **🎭 Assigned Specialist Persona** to assign custom or built-in personas to individual steps.
     * Use **⚡ Assigned Capabilities** multiselect dropdown to attach custom or built-in capabilities to each step.
   * **Step 4**: Check **☑️ 💾 Auto-Trigger File Download (.docx / .xlsx / .pptx)** if desired.
   * Click **💾 Save Agent to Local Storage**.
3. Click **Run Routine** on any agent card to execute.

### Step 7: Manage Memory Rules & Memory Ledger
1. Open **Settings ⚙️ ➔ Memory Hub**.
2. View, category-filter, pin, edit, or delete persistent rules (e.g., policy baselines, DOA thresholds).
3. **Auto-Detection**: When typing prompts containing phrases like *"Remember that..."*, click the **Save Memory Rule** banner to add it directly to your persistent knowledge vault.
4. **Memory Profile Budget**: Adjust the Memory Profile dropdown (*Small*, *Medium*, *Large*, or *No Limit*) based on dataset size and available system RAM/VRAM.

### Step 8: Start Sessions & Manage Session Groups
1. Click **New Intelligence Session (+)** in the left sidebar.
2. **Creating a Group**: Click **Create Group** beside the **SESSION HISTORY** header in the sidebar, pick an icon and color, and enter a group name.
3. **Organizing Sessions**: Drag any session card from the sidebar and drop it directly onto a group card.

### Step 9: Attach & Extract Documents
1. Click the **Attach File** button (📎) next to the prompt input box.
2. Select an Excel spreadsheet, Word document, PDF, PowerPoint deck, or image.
3. Enter your prompt and click **Send** (🚀).

### Step 10: Use Interactive Work Desk & Inline AI Rewriter
1. Click **Work Desk** (🗂️) in the top header.
2. Click **Load into Work Desk** on any chat response.
3. **Inline AI Rewriting**:
   * Highlight any text inside the Work Desk Notepad.
   * Enter a directive in the floating AI bubble (e.g., *"Make this more concise"*).
   * Click **✓ Apply to Notepad**.
4. Switch views between Word Canvas, Excel Grid, Slide Deck Carousel, and Agent Studio.

### Step 11: Export Work Desk Deliverables
In the Work Desk header toolbar, click:
* **📄 Export Word (.doc)**
* **📊 Export Excel (.xlsx)**
* **📽️ Export PowerPoint (.pptx)**
* **📝 Export Markdown (.md)**

### Step 12: Managing Token Limits & Generating Multi-Turn Reports
Because single-turn local model outputs are capped at ~2,048 tokens (~1,500 words) to protect VRAM:
1. **Initial Turn**: Prompt the model for Part 1 & Part 2.
2. **First Continuation Turn**: Type: *"Continue with Part 3 and Part 4: Root Cause Analysis & Financial Exposure"*.
3. **Second Continuation Turn**: Type: *"Continue with Part 5: Actionable Corrective Action Plan & Conclusion"*.
4. Sentinel Studio appends each new section directly into your active Work Desk document seamlessly.

---

*Architected & Engineered by Patrick Jay F. Tanap, CPA, CICA | Sentinel Studio™ Enterprise Corporate Intelligence Workstation*
