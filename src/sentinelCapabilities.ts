/**
 * =====================================================================
 * SENTINEL STUDIO™ — DYNAMIC CAPABILITIES ARCHITECTURE
 * Storage: IndexedDB (`sentinel_capabilities_store`, store: `capabilities`)
 * Lifecycle: Ephemeral Per-Turn Context Injection & WebGPU Memory Safeguards
 * Architect & Developer: Patrick Jay F. Tanap, CPA, CICA
 * Copyright © 2026 Patrick Jay F. Tanap. All Rights Reserved.
 * =====================================================================
 */

export interface SentinelCapability {
  id: string; // e.g. "user:ph-statutory-compliance"
  command: string; // e.g. "tax" (without /# prefix)
  name: string; // e.g. "Philippine Statutory & Tax Compliance"
  description: string; // 1-2 sentence trigger summary used in passive registry and autocomplete
  instructions: string; // Full execution playbook (Markdown)
  enabled: boolean;
  isSystem: boolean; // protected from deletion
  updatedAt: string;
  aliases?: string[]; // secondary aliases (e.g. ["bir"])
}

export interface ParsedCapabilityScanResult {
  matchedCapabilities: SentinelCapability[];
  strippedText: string;
  excessCount: number;
}

export const CAPABILITIES_DB_NAME = 'sentinel_capabilities_store';
export const CAPABILITIES_STORE_NAME = 'capabilities';
export const CAPABILITIES_DB_VERSION = 1;

export const MAX_ACTIVE_CAPABILITIES = 2;

/**
 * Pre-seeded Foundational Enterprise Capabilities
 */
export const FOUNDATIONAL_CAPABILITIES: SentinelCapability[] = [
  {
    id: 'user:ph-statutory-compliance',
    command: 'tax',
    name: 'Philippine Statutory & Tax Compliance',
    description: 'Use when auditing Philippine tax schedules, BIR Form 2307 withholding, SSS/PhilHealth/Pag-IBIG remittances, or RA 10173 DPA compliance.',
    enabled: true,
    isSystem: true,
    updatedAt: new Date().toISOString(),
    aliases: ['bir'],
    instructions: `### PHILIPPINE STATUTORY, TAX & RA 10173 DATA PRIVACY COMPLIANCE PLAYBOOK

#### 1. BIR FORM 2307 & EXPANDED WITHHOLDING TAX (EWT) RULES
- **Governing Mandate**: National Internal Revenue Code (NIRC) as amended by TRAIN (RA 10963) and EOPT (RA 11976).
- **Mandatory Certificate Timing**: Form 2307 must be furnished to payee on or before the 20th day of the month following the close of the taxable quarter, or upon demand simultaneously with payment.
- **Top Withholding Tax Rates & ATCs**:
  * **WI 010 / WC 010**: Professional Fees (Lawyers, CPAs, Engineers, Consultants):
    - Individuals: 5% if gross income ≤ ₱3,000,000 (with sworn BIR declaration); otherwise 10%.
    - Corporations: 10% if gross income ≤ ₱720,000; otherwise 15%.
  * **WI 100 / WC 100**: Rental of Real / Personal Property: 5%.
  * **WI 160 / WC 160**: Contractors & Service Subcontractors: 2%.
  * **WI 158 / WC 158**: Purchase of Goods by Top Withholding Agents (TWA): 1%.
  * **WI 640 / WC 640**: Electronic Marketplace / Digital Platform Merchant Settlements: 1% on 50% gross remittances (under RR 16-2023).

#### 2. STATUTORY REMITTANCES & COMPENSATION THRESHOLDS
- **Social Security System (SSS)**: Total contribution 14% (Employer: 9.5%, Employee: 4.5%), Maximum Monthly Salary Credit (MSC) ₱30,000.
- **PhilHealth**: 5.0% premium contribution rate split 50/50 between employer and employee; ceiling ₱100,000.
- **Pag-IBIG (HDMF)**: Maximum mandatory contribution capped at ₱200/month (₱100 ER / ₱100 EE).
- **De Minimis Benefits**: Rice allowance (₱2,000/mo), Uniform allowance (₱6,000/yr), Medical cash allowance (₱1,500/semester) exempt from withholding tax on compensation.

#### 3. REPUBLIC ACT NO. 10173 (PHILIPPINE DATA PRIVACY ACT OF 2012)
- **Sensitive Personal Information (SPI)**: Tax records, TIN, salary slips, and government IDs require explicit consent and encryption.
- **72-Hour Breach Notification**: Mandatory notification to the National Privacy Commission (NPC) and affected data subjects within 72 hours of discovery if breach involves sensitive data affecting 100+ subjects.
- **Data Minimization & Storage**: PII must not be transmitted in plaintext or exported to untrusted cloud endpoints without verified DPIA.`,
  },
  {
    id: 'user:lss-fmea-risk-calculator',
    command: 'fmea',
    name: 'LSS FMEA & Risk Calculator',
    description: 'Use when conducting Lean Six Sigma risk assessments, AIAG & VDA RPN calculations, or Poka-Yoke mitigation planning.',
    enabled: true,
    isSystem: true,
    updatedAt: new Date().toISOString(),
    instructions: `### LEAN SIX SIGMA AIAG & VDA FMEA RISK ASSESSMENT & POKA-YOKE PLAYBOOK

#### 1. SCORING TAXONOMY (SCALE 1 TO 10)
- **Severity (S)**:
  - **9-10**: Catastrophic; safety, regulatory violation, or complete legal/statutory shutdown without warning.
  - **7-8**: High; primary function disrupted, significant customer dissatisfaction, or major financial audit finding.
  - **4-6**: Moderate; performance degraded, rework required, operational delay within buffer limits.
  - **1-3**: Minor/None; cosmetic defect, noticeable only to discerning operators with minimal impact.
- **Occurrence (O)**:
  - **9-10**: Almost Inevitable; recurring daily or in > 10% of transaction cycles.
  - **7-8**: Frequent; recurring weekly with documented historical incidents (2% to 5%).
  - **4-6**: Occasional; sporadic occurrences (0.1% to 1%).
  - **1-3**: Remote; highly capable process with Cpk >= 1.67, virtually zero historical failures.
- **Detection (D)**:
  - **9-10**: Almost Impossible; defect passes through undetected to external stakeholders/customers.
  - **7-8**: Low; manual visual inspection or sample-based audit with high risk of escape.
  - **4-6**: Moderate; automated batch checks or standard statistical process control sampling.
  - **1-3**: High/Certain; 100% automated Poka-Yoke hardware/software gate blocks defect propagation immediately.

#### 2. RISK PRIORITY NUMBER (RPN) CALCULATION & MITIGATION RULES
- **Formula**: \`RPN = Severity (S) × Occurrence (O) × Detection (D)\` [Range: 1 to 1,000].
- **CRITICAL MITIGATION TRIGGERS**:
  - **RPN >= 100**: Mandatory Poka-Yoke architectural mitigation and engineered control design.
  - **Severity (S) >= 9**: Immediate executive notification and root-cause engineering review, regardless of RPN.
  - **Action Priority (AP)**:
    * **High (H)**: Highest priority for preventive redesign.
    * **Medium (M)**: Secondary priority; enhanced detection or process optimization.
    * **Low (L)**: Acceptable risk under active monitoring.

#### 3. DELIVERABLE MATRIX STRUCTURE
When presenting an FMEA review, formulate structured tables containing:
\`| Process Step | Potential Failure Mode | Potential Effect | S | Potential Cause | O | Current Controls | D | RPN | Poka-Yoke Mitigation Plan |\``,
  },
  {
    id: 'user:anaplan-solution-architect',
    command: 'anaplan',
    name: 'Anaplan Solution Architect',
    description: 'Use when designing Anaplan models, multidimensional modules, or Close tracking architectures.',
    enabled: true,
    isSystem: true,
    updatedAt: new Date().toISOString(),
    instructions: `### ANAPLAN ENTERPRISE ARCHITECTURE & MULTIDIMENSIONAL MODEL DESIGN

#### 1. THE DISCO METHODOLOGY FRAMEWORK
Every Anaplan model must strictly follow DISCO modularization:
- **D — Data Modules**:
  - Purpose: Pure transactional staging, flat system-of-record feeds, zero calculated business logic.
  - Rule: Minimize dimensions (typically List + Line Items only); clear formatting.
- **I — Input Modules**:
  - Purpose: Collect user assumptions, forecast overrides, departmental headcounts, driver entries.
  - Rule: Keep editable grids isolated from heavy dynamic recalculation loops.
- **S — System Modules**:
  - Purpose: Store centralized business attributes, list properties, calendar mappings, versions, security tiers.
  - Rule: Referenced universally via LOOKUP to eliminate duplicate formulas.
- **C — Calculation Modules**:
  - Purpose: Execute heavy multidimensional business logic, allocations, variances, and Boolean filter flags.
  - Rule: Format each Line Item as Boolean whenever possible; optimize dimensional combinations.
- **O — Output Modules**:
  - Purpose: Fast reporting dashboards, Board matrices, executive KPI scorecards, export grids.
  - Rule: Read-only summary grids pulling exclusively from Calculation and System modules.

#### 2. HYPERBLOCK ENGINE PERFORMANCE & MEMORY OPTIMIZATION
- **Avoid Deep Nested IF-THEN-ELSE**:
  - Use Boolean formatted line items as filter masks.
  - Never calculate strings in calculation loops; use Codes and Number formats.
- **Sparsity Management**:
  - Avoid cross-multiplying large sparse lists (e.g., Accounts × Cost Centers × Employees × Products).
  - Break into specialized two-dimension calculation modules.
- **Formula Referencing**:
  - Use \`LOOKUP\` when referencing single source system modules.
  - Restrict \`SUM\` to aggregation mappings.

#### 3. FINANCIAL CLOSE TRACKING SPECIFICATION
- **Core Dimensions**: Time (Month / Day) × Functional Stream (R2R, P2P, O2C, Tax) × Close Step List.
- **Line Items**: Task Status (List), Owner (User List), Target TAT (Hours), Actual Completion (Date/Time), SLA Variance Flag (Boolean).`,
  },
];

/**
 * Opens or initializes the IndexedDB store for capabilities.
 * Also migrates legacy sentinel_skills_store records if found.
 */
export function openCapabilitiesDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(CAPABILITIES_DB_NAME, CAPABILITIES_DB_VERSION);

    request.onupgradeneeded = (event: any) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains(CAPABILITIES_STORE_NAME)) {
        const store = db.createObjectStore(CAPABILITIES_STORE_NAME, { keyPath: 'id' });
        store.createIndex('command', 'command', { unique: false });
        store.createIndex('enabled', 'enabled', { unique: false });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/**
 * Initializes the capabilities store and pre-seeds foundational capabilities if empty.
 */
export async function initCapabilitiesStore(): Promise<SentinelCapability[]> {
  const db = await openCapabilitiesDB();

  // Try migrating any legacy skills from sentinel_skills_store if available
  try {
    const legacyRecords = await checkAndMigrateLegacySkills();
    if (legacyRecords && legacyRecords.length > 0) {
      for (const rec of legacyRecords) {
        await saveCapability(rec);
      }
    }
  } catch (_e) {}

  const current = await getAllCapabilities();
  if (current.length === 0) {
    for (const cap of FOUNDATIONAL_CAPABILITIES) {
      await saveCapability(cap);
    }
    return [...FOUNDATIONAL_CAPABILITIES];
  }

  // Ensure foundational capabilities are always present and up-to-date
  for (const foundCap of FOUNDATIONAL_CAPABILITIES) {
    const exists = current.find((c) => c.id === foundCap.id);
    if (!exists) {
      await saveCapability(foundCap);
      current.push(foundCap);
    }
  }

  return current;
}

/**
 * Checks for any existing records in legacy `sentinel_skills_store` to migrate seamlessly.
 */
async function checkAndMigrateLegacySkills(): Promise<SentinelCapability[] | null> {
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open('sentinel_skills_store');
      req.onsuccess = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('skills')) {
          db.close();
          return resolve(null);
        }
        const tx = db.transaction('skills', 'readonly');
        const store = tx.objectStore('skills');
        const getAllReq = store.getAll();
        getAllReq.onsuccess = () => {
          db.close();
          const items = getAllReq.result || [];
          resolve(items);
        };
        getAllReq.onerror = () => {
          db.close();
          resolve(null);
        };
      };
      req.onerror = () => resolve(null);
    } catch (_e) {
      resolve(null);
    }
  });
}

/**
 * Retrieves all registered capabilities.
 */
export async function getAllCapabilities(): Promise<SentinelCapability[]> {
  const db = await openCapabilitiesDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CAPABILITIES_STORE_NAME, 'readonly');
    const store = tx.objectStore(CAPABILITIES_STORE_NAME);
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

/**
 * Retrieves a capability by ID.
 */
export async function getCapabilityById(id: string): Promise<SentinelCapability | null> {
  const db = await openCapabilitiesDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CAPABILITIES_STORE_NAME, 'readonly');
    const store = tx.objectStore(CAPABILITIES_STORE_NAME);
    const req = store.get(id);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

/**
 * Saves or updates a capability record.
 */
export async function saveCapability(capability: SentinelCapability): Promise<void> {
  const db = await openCapabilitiesDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CAPABILITIES_STORE_NAME, 'readwrite');
    const store = tx.objectStore(CAPABILITIES_STORE_NAME);
    const req = store.put({
      ...capability,
      updatedAt: new Date().toISOString(),
    });
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

/**
 * Deletes a capability record (protects foundational system capabilities).
 */
export async function deleteCapability(id: string): Promise<boolean> {
  const target = await getCapabilityById(id);
  if (target?.isSystem) {
    throw new Error('Foundational system capabilities cannot be deleted.');
  }

  const db = await openCapabilitiesDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(CAPABILITIES_STORE_NAME, 'readwrite');
    const store = tx.objectStore(CAPABILITIES_STORE_NAME);
    const req = store.delete(id);
    req.onsuccess = () => resolve(true);
    req.onerror = () => reject(req.error);
  });
}

/**
 * Toggles a capability's enabled state.
 */
export async function toggleCapabilityEnabled(id: string): Promise<SentinelCapability | null> {
  const target = await getCapabilityById(id);
  if (!target) return null;
  target.enabled = !target.enabled;
  await saveCapability(target);
  return target;
}

/**
 * Scans a user input string globally for all occurrences of /#([a-zA-Z0-9_-]+),
 * matches extracted aliases against enabled capabilities, enforces MAX_ACTIVE_CAPABILITIES = 2,
 * and strips matched tokens from user text sent to the LLM.
 */
export function scanAndMatchCapabilities(
  rawInput: string,
  capabilities: SentinelCapability[]
): ParsedCapabilityScanResult {
  if (!rawInput) {
    return { matchedCapabilities: [], strippedText: '', excessCount: 0 };
  }

  const regex = /\/#([a-zA-Z0-9_-]+)/g;
  const matches = [...rawInput.matchAll(regex)];

  if (matches.length === 0) {
    return { matchedCapabilities: [], strippedText: rawInput, excessCount: 0 };
  }

  const enabled = capabilities.filter((c) => c.enabled);
  const foundCaps: SentinelCapability[] = [];
  const matchedTokens: string[] = [];

  for (const match of matches) {
    const fullToken = match[0];
    const alias = match[1].toLowerCase();

    const matched = enabled.find(
      (c) =>
        c.command.toLowerCase() === alias ||
        c.id.toLowerCase() === `user:${alias}` ||
        c.id.toLowerCase() === alias ||
        (c.aliases && c.aliases.map((a) => a.toLowerCase()).includes(alias))
    );

    if (matched && !foundCaps.some((c) => c.id === matched.id)) {
      foundCaps.push(matched);
      matchedTokens.push(fullToken);
    }
  }

  const excessCount = Math.max(0, foundCaps.length - MAX_ACTIVE_CAPABILITIES);
  const activeCaps = foundCaps.slice(0, MAX_ACTIVE_CAPABILITIES);

  // Strip matched /#alias tokens from prompt text
  let strippedText = rawInput;
  matchedTokens.forEach((token) => {
    const esc = token.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    strippedText = strippedText.replace(new RegExp(`(?:^|\\s)${esc}(?=\\s|$)`, 'g'), ' ');
  });

  strippedText = strippedText.replace(/\s+/g, ' ').trim();

  return {
    matchedCapabilities: activeCaps,
    strippedText,
    excessCount,
  };
}

/**
 * Formulates the ephemeral active capabilities playbook XML block.
 */
export function buildActiveCapabilitiesBlock(capabilities: SentinelCapability[]): string {
  if (!capabilities || capabilities.length === 0) return '';

  const playbooks = capabilities
    .map(
      (c, idx) =>
        `### Capability: ${c.id} (command: /#${c.command})\n${c.instructions.trim()}`
    )
    .join('\n\n---\n\n');

  return `<active_capabilities>
${playbooks}
</active_capabilities>`;
}

/**
 * Formulates the Tier-1 Passive Prompt Index (When no /# tag is explicitly used).
 */
export function buildPassiveCapabilitiesIndex(allCapabilities: SentinelCapability[]): string {
  const enabled = allCapabilities.filter((c) => c.enabled);
  if (enabled.length === 0) return '';

  const listItems = enabled
    .map((c) => `- \`${c.id}\` (command: \`/#${c.command}\`): ${c.description.trim()}`)
    .join('\n');

  return `<available_expandable_capabilities>
${listItems}
</available_expandable_capabilities>

CAPABILITY EXPANSION RULE:
If the user request requires specialized domain procedures found in <available_expandable_capabilities>, emit the tool call:
\`\`\`call:retriever:expand_capabilities{api_names: ["capability_id_1", "capability_id_2"]}\`\`\``;
}

/**
 * Strips heavy <active_capabilities> and <active_skills> playbooks from stored chat history
 * and records only the lightweight metadata tag [CAPABILITY_USED: id1, id2] to protect KV-cache.
 */
export function stripActiveCapabilitiesForStorage(
  composedPrompt: string,
  usedCapabilityIds: string[]
): string {
  let cleaned = composedPrompt
    .replace(/<active_capabilities>[\s\S]*?<\/active_capabilities>/gi, '')
    .replace(/<active_skills>[\s\S]*?<\/active_skills>/gi, '')
    .replace(/<available_expandable_capabilities>[\s\S]*?<\/available_expandable_capabilities>/gi, '')
    .replace(/<available_expandable_skills>[\s\S]*?<\/available_expandable_skills>/gi, '')
    .replace(/CAPABILITY EXPANSION RULE:[\s\S]*?```call:retriever:expand_capabilities[\s\S]*?```/gi, '')
    .replace(/SKILL EXPANSION RULE:[\s\S]*?```call:retriever:expand_skills[\s\S]*?```/gi, '')
    .trim();

  if (usedCapabilityIds.length > 0) {
    cleaned += `\n[CAPABILITY_USED: ${usedCapabilityIds.join(', ')}]`;
  }

  return cleaned.trim();
}

/**
 * Detects autonomous capability expansion tool call emitted by LLM:
 * ```call:retriever:expand_capabilities{api_names: ["capability_id_1", "capability_id_2"]}```
 * Also supports backward compatible `call:retriever:expand_skills`.
 */
export function detectAutonomousCapabilityExpansion(rawOutput: string): string[] | null {
  if (!rawOutput) return null;

  const toolCallRegex = /(?:```)?call:retriever:expand_(?:capabilities|skills)\{api_names:\s*(\[[^\]]*\])\}(?:```)?/i;
  const match = rawOutput.match(toolCallRegex);

  if (match && match[1]) {
    try {
      const rawArrayStr = match[1].replace(/'/g, '"');
      const parsed = JSON.parse(rawArrayStr);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.slice(0, MAX_ACTIVE_CAPABILITIES).map((id) => String(id).trim());
      }
    } catch (_e) {
      const ids = [...match[1].matchAll(/["']([^"']+)["']/g)].map((m) => m[1].trim());
      if (ids.length > 0) {
        return ids.slice(0, MAX_ACTIVE_CAPABILITIES);
      }
    }
  }

  return null;
}

// Aliases for backwards compatibility
export type SentinelSkill = SentinelCapability;
export const MAX_ACTIVE_SKILLS = MAX_ACTIVE_CAPABILITIES;
export const initSkillsStore = initCapabilitiesStore;
export const getAllSkills = getAllCapabilities;
export const getSkillById = getCapabilityById;
export const saveSkill = saveCapability;
export const deleteSkill = deleteCapability;
export const toggleSkillEnabled = toggleCapabilityEnabled;
export const scanAndMatchSkills = scanAndMatchCapabilities;
export const buildActiveSkillsBlock = buildActiveCapabilitiesBlock;
export const buildPassiveSkillsIndex = buildPassiveCapabilitiesIndex;
export const stripActiveSkillsForStorage = stripActiveCapabilitiesForStorage;
export const detectAutonomousSkillExpansion = detectAutonomousCapabilityExpansion;
