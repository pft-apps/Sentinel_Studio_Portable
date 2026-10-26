/**
 * =====================================================================
 * SENTINEL STUDIO™ — PERSISTENT MEMORY ENGINE & MULTI-TIER BACKUP SUITE
 * Enterprise Capability Hub
 * Architect & Developer: Patrick Jay F. Tanap, CPA, CICA
 * Statutory Compliance: RA 10173 (Philippine Data Privacy Act)
 * 100% Client-Side Local Execution • Zero Cloud Egress
 * =====================================================================
 */

export interface MemoryRecord {
  id: string; // mem_<timestamp>_<random>
  category: 'Domain Context' | 'Policy Rule' | 'Threshold' | 'Correction' | 'User Preference';
  content: string;
  confidence: 'Verified' | 'High' | 'User Defined';
  sourceChatId: string;
  createdAt: string; // ISO 8601
  updatedAt: string; // ISO 8601
  isActive: boolean;
  isPinned: boolean;
  usageCount: number;
}

export type MemoryCategory = MemoryRecord['category'];
export type MemoryConfidence = MemoryRecord['confidence'];

export const MEMORY_CATEGORIES: MemoryCategory[] = [
  'Domain Context',
  'Policy Rule',
  'Threshold',
  'Correction',
  'User Preference',
];

export const DEFAULT_SEED_MEMORIES: Omit<MemoryRecord, 'id' | 'createdAt' | 'updatedAt'>[] = [
  {
    category: 'Domain Context',
    content: 'Organization: Filinvest Business Services Corporation – Enterprise Capability Hub (FBSC ECH).',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: true,
    usageCount: 0,
  },
  {
    category: 'Domain Context',
    content: 'Department: Governance & Business Excellence under a cost-neutral operating model.',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: true,
    usageCount: 0,
  },
  {
    category: 'Policy Rule',
    content: 'Audit Observations: Findings must strictly follow the 5-part structure: [Condition, Criteria, Cause, Impact, Recommendation].',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: true,
    usageCount: 0,
  },
  {
    category: 'Policy Rule',
    content: 'Statutory Compliance: Strict zero external cloud data egress under RA 10173; 100% on-device WebGPU execution.',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: true,
    usageCount: 0,
  },
  {
    category: 'Correction',
    content: 'Invoice Processing: Invoice intake metrics represent intake review queues, not completed financial settlement.',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: true,
    usageCount: 0,
  },
  {
    category: 'Threshold',
    content: 'Accounts Payable: Flag split Purchase Orders processed within 48 hours totaling below the PHP 500,000 Delegation of Authority (DOA) threshold.',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: false,
    usageCount: 0,
  },
  {
    category: 'Threshold',
    content: 'Defect Benchmark: Process cycle time variance exceeding >= 45% must be escalated as a high-priority operational defect.',
    confidence: 'Verified',
    sourceChatId: 'system_preseed',
    isActive: true,
    isPinned: false,
    usageCount: 0,
  },
];

const DB_NAME = 'sentinel_memory_vault';
const DB_VERSION = 2;
const STORE_MEMORIES = 'persistent_memory';
const STORE_CONFIG = 'system_config';

let dbInstance: IDBDatabase | null = null;
let hotDirHandleCached: any = null;

// =====================================================================
// SECTION 1: INDEXEDDB MEMORY VAULT INITIALIZATION & CRUD
// =====================================================================

export function openMemoryVaultDB(): Promise<IDBDatabase> {
  if (dbInstance) return Promise.resolve(dbInstance);

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event: IDBVersionChangeEvent) => {
      const db = (event.target as IDBOpenDBRequest).result;
      if (!db.objectStoreNames.contains(STORE_MEMORIES)) {
        db.createObjectStore(STORE_MEMORIES, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(STORE_CONFIG)) {
        db.createObjectStore(STORE_CONFIG, { keyPath: 'key' });
      }
    };

    request.onsuccess = (event: Event) => {
      dbInstance = (event.target as IDBOpenDBRequest).result;
      resolve(dbInstance);
    };

    request.onerror = () => {
      reject(request.error || new Error('Failed to open sentinel_memory_vault IndexedDB'));
    };
  });
}

export const openMemoryDB = openMemoryVaultDB;

/**
 * Initializes the memory vault, verifies once-only seed data, and returns all records.
 * FIX: Uses localStorage guard 'sentinel_memory_preseeded_v1'. Never re-injects deleted rules on reload.
 */
export async function initMemoryVault(): Promise<MemoryRecord[]> {
  const db = await openMemoryVaultDB();
  const existing = await getPersistentMemories();
  const isPreseeded = typeof localStorage !== 'undefined' && localStorage.getItem('sentinel_seed_done') === 'true';

  // Only seed if NOT previously seeded AND current store is empty
  if (!isPreseeded && existing.length === 0) {
    const now = new Date().toISOString();
    const seededRecords: MemoryRecord[] = DEFAULT_SEED_MEMORIES.map((seed, idx) => ({
      ...seed,
      id: `mem_seed_${idx + 1}_${Math.random().toString(36).slice(2, 7)}`,
      createdAt: now,
      updatedAt: now,
    }));

    const tx = db.transaction(STORE_MEMORIES, 'readwrite');
    const store = tx.objectStore(STORE_MEMORIES);
    for (const record of seededRecords) {
      store.put(record);
    }
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });

    try {
      localStorage.setItem('sentinel_seed_done', 'true');
    } catch (_e) {}

    console.info(`[Sentinel Memory] Initialized vault with ${seededRecords.length} default knowledge seeds.`);
    return seededRecords;
  }

  // If already seeded or has existing records, ensure flag is stored so user deletions are permanent
  if (!isPreseeded && (existing.length > 0 || isPreseeded === null)) {
    try {
      localStorage.setItem('sentinel_seed_done', 'true');
    } catch (_e) {}
  }

  return existing;
}

export async function getPersistentMemories(): Promise<MemoryRecord[]> {
  const db = await openMemoryVaultDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_MEMORIES, 'readonly');
    const store = tx.objectStore(STORE_MEMORIES);
    const req = store.getAll();
    req.onsuccess = () => {
      const list = (req.result as MemoryRecord[]) || [];
      resolve(list);
    };
    req.onerror = () => reject(req.error);
  });
}

export const getAllMemories = getPersistentMemories;

export async function addMemoryRecord(
  partial: {
    category: MemoryCategory;
    content: string;
    confidence?: MemoryConfidence;
    sourceChatId?: string;
    isActive?: boolean;
    isPinned?: boolean;
  }
): Promise<MemoryRecord> {
  const db = await openMemoryVaultDB();
  const now = new Date().toISOString();
  const id = `mem_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const record: MemoryRecord = {
    id,
    category: partial.category,
    content: partial.content.trim(),
    confidence: partial.confidence || 'User Defined',
    sourceChatId: partial.sourceChatId || 'manual_entry',
    createdAt: now,
    updatedAt: now,
    isActive: partial.isActive !== false,
    isPinned: partial.isPinned === true,
    usageCount: 0,
  };

  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE_MEMORIES, 'readwrite');
    const store = tx.objectStore(STORE_MEMORIES);
    const req = store.put(record);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });

  return record;
}

export async function saveMemoryRecord(record: MemoryRecord): Promise<void> {
  const db = await openMemoryVaultDB();
  record.updatedAt = new Date().toISOString();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_MEMORIES, 'readwrite');
    const store = tx.objectStore(STORE_MEMORIES);
    const req = store.put(record);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

export const updateMemoryRecord = saveMemoryRecord;

export async function deleteMemoryRecord(id: string): Promise<void> {
  const db = await openMemoryVaultDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_MEMORIES, 'readwrite');
    const store = tx.objectStore(STORE_MEMORIES);
    // Ensure id matches type
    store.delete(String(id));
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function pruneStaleMemories(): Promise<{ deletedCount: number; remainingCount: number }> {
  const memories = await getPersistentMemories();
  const db = await openMemoryVaultDB();
  
  const seenContent = new Set<string>();
  const idsToDelete: string[] = [];

  for (const m of memories) {
    const normalized = m.content.trim().toLowerCase();
    if (seenContent.has(normalized)) {
      // Duplicate rule
      idsToDelete.push(m.id);
    } else {
      seenContent.add(normalized);
      // Unused and unpinned rule
      if (m.usageCount === 0 && !m.isPinned) {
        idsToDelete.push(m.id);
      }
    }
  }

  if (idsToDelete.length > 0) {
    const tx = db.transaction(STORE_MEMORIES, 'readwrite');
    const store = tx.objectStore(STORE_MEMORIES);
    for (const id of idsToDelete) {
      store.delete(id);
    }
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  const remaining = memories.length - idsToDelete.length;
  return { deletedCount: idsToDelete.length, remainingCount: remaining };
}

// =====================================================================
// SECTION 1.3: IN-CONTEXT MEMORY INJECTION PIPELINE
// =====================================================================

export async function compilePromptWithMemory(
  baseSystemPrompt: string = '',
  userPrompt: string = '',
  activeAgent: any = null,
  approvedPastChatExcerpts: string = ''
): Promise<{
  compiledSystemPrompt: string;
  activeCount: number;
  injectedMemories: MemoryRecord[];
  totalChars: number;
}> {
  const allMemories = await getPersistentMemories();
  const activeMemories = allMemories.filter((m) => m.isActive);

  if (activeMemories.length === 0 && !approvedPastChatExcerpts) {
    return {
      compiledSystemPrompt: baseSystemPrompt,
      activeCount: 0,
      injectedMemories: [],
      totalChars: 0,
    };
  }

  // Step A: Core Rules (Pinned, max 7)
  const pinnedMemories = activeMemories
    .filter((m) => m.isPinned)
    .slice(0, 7);

  // Step B: Contextual Scoring for non-pinned active rules (top 10)
  const otherActive = activeMemories.filter((m) => !m.isPinned);
  const scoredMemories = otherActive.map((m) => {
    let score = 0;
    // Agent category/keyword match
    if (activeAgent) {
      if (m.category === activeAgent.category) score += 5;
      if (activeAgent.name.toLowerCase().includes(m.category.toLowerCase())) score += 5;
    }
    // Keyword overlap with user prompt
    if (userPrompt) {
      const stopWords = new Set(['this', 'that', 'with', 'from', 'have', 'your', 'about', 'some', 'what']);
      const queryWords = userPrompt
        .toLowerCase()
        .replace(/[^\w\s]/g, '')
        .split(/\s+/)
        .filter((w) => w.length > 3 && !stopWords.has(w));
      
      const contentLower = m.content.toLowerCase();
      queryWords.forEach((word) => {
        if (contentLower.includes(word)) score += 3;
      });
    }
    // Confidence bonus
    if (m.confidence === 'Verified') score += 1;

    return { record: m, score };
  });

  // Sort and select Top 10
  const topContextual = scoredMemories
    .sort((a, b) => b.score - a.score)
    .slice(0, 10)
    .map((s) => s.record);

  // Combine
  let finalSelection = [...pinnedMemories];
  topContextual.forEach((m) => {
    if (!finalSelection.find((fs) => fs.id === m.id)) {
      finalSelection.push(m);
    }
  });

  // Step C: Enforce Hard Token Budget Cap (Strict combined limit of max 2,800 characters / ~700 tokens)
  const MAX_COMBINED_CHARS = 2800;
  let selectedToInject: MemoryRecord[] = [];
  let currentChars = 0;

  // Reserve space for past chat context if present
  const pastChatReservedChars = approvedPastChatExcerpts ? Math.min(1000, approvedPastChatExcerpts.length + 150) : 0;
  const memoryCharBudget = MAX_COMBINED_CHARS - pastChatReservedChars;

  for (const m of finalSelection) {
    const line = `- [${m.category}] ${m.content}\n`;
    if (currentChars + line.length > memoryCharBudget) break;
    selectedToInject.push(m);
    currentChars += line.length;
  }

  const memoryLines = selectedToInject.length > 0
    ? selectedToInject.map((m) => `- [${m.category}] ${m.content}`).join('\n')
    : '(No persistent memory rules active)';

  let pastContextBlock = '';
  if (approvedPastChatExcerpts) {
    const remainingForPast = MAX_COMBINED_CHARS - currentChars - 250;
    const trimmedExcerpts = approvedPastChatExcerpts.slice(0, Math.max(150, remainingForPast));
    pastContextBlock = `\n\n=== RECALLED HISTORICAL CROSS-CHAT CONTEXT (USER APPROVED) ===\n${trimmedExcerpts}\n============================================================`;
  }

  const executivePreamble = [
    '=== PERSISTENT WORKING KNOWLEDGE & ESTABLISHED CORPORATE RULES ===',
    'You have verified working knowledge of the following established corporate rules, historical corrections, and user preferences:',
    memoryLines,
    'Apply these established rules automatically without asking the user to re-state them.',
    '==================================================================',
    pastContextBlock,
  ].filter(Boolean).join('\n');

  // Increment usageCount and refresh updatedAt in IndexedDB
  const now = new Date().toISOString();
  for (const m of selectedToInject) {
    m.usageCount = (m.usageCount || 0) + 1;
    m.updatedAt = now;
    await saveMemoryRecord(m);
  }

  const compiledSystemPrompt = baseSystemPrompt
    ? `${executivePreamble}\n\n${baseSystemPrompt}`
    : executivePreamble;

  return {
    compiledSystemPrompt,
    activeCount: selectedToInject.length,
    injectedMemories: selectedToInject,
    totalChars: executivePreamble.length,
  };
}



// =====================================================================
// SECTION 1.4: DYNAMIC RULE EXTRACTION & DETECTION
// =====================================================================

export interface DetectedRule {
  triggerPhrase: string;
  ruleContent: string;
  suggestedCategory: MemoryCategory;
}

export function detectNewRules(userPromptText: string): DetectedRule | null {
  if (!userPromptText || typeof userPromptText !== 'string') return null;

  // Split prompt into individual sentences or logical lines to catch embedded rules anywhere in a broader prompt
  const sentences = userPromptText
    .split(/(?:\n+|\. |\? |\! |; )/)
    .map((s) => s.trim())
    .filter((s) => s.length > 5);

  // Flexible intent patterns to match inside any sentence or clause
  const ruleIntentPatterns: Array<{
    triggerRegex: RegExp;
    category: MemoryCategory;
  }> = [
    // 1. Thresholds / Benchmarks / Limits (e.g. "our SLA limit is 48 hours")
    {
      triggerRegex: /(?:(?:our|the|my|system|department|company)\s+)?(?:threshold|limit|benchmark|cutoff|sla|doa|target|allowance|max|maximum|min|minimum)\s+(?:is|shall\s+be|must\s+be|should\s+be|set\s+to|=|at)\s+(.+)/i,
      category: 'Threshold',
    },
    {
      triggerRegex: /(?:set|cap|limit)\s+(?:the|our|my)?\s*(?:threshold|limit|sla|benchmark|cutoff)\s+(?:to|at)\s+(.+)/i,
      category: 'Threshold',
    },
    // 2. Memory / Policy Directives (e.g. "please remember that ...")
    {
      triggerRegex: /(?:remember\s+(?:that|this|to|when)?|keep\s+in\s+mind\s+(?:that)?|note\s+(?:that|this)?|take\s+note\s+(?:that|of)?|make\s+a\s+note\s+(?:that|to)?|don'?t\s+forget\s+(?:that|to)?|(?:store|save|record)\s+this\s+(?:rule|policy|setting|preference)?|our\s+policy\s+is\s+that)\s*[:\-,]?\s*(.+)/i,
      category: 'Policy Rule',
    },
    // 3. Formatting / Style Preferences (e.g. "always format output as ...")
    {
      triggerRegex: /(?:(?:always|preferably|strictly|from\s+now\s+on)\s+(?:format|output|write|generate|use|structure|render|style|provide)|(?:my|our)\s+(?:preference|style|format)\s+is|prefer\s+(?:to|using|formatting))\s*[:\-,]?\s*(.+)/i,
      category: 'User Preference',
    },
    // 4. Corrections & Overrides (e.g. "no actually that's wrong, ...")
    {
      triggerRegex: /(?:(?:no|wait|actually|hold\s+on)[,\s]+(?:actually|that'?s\s+wrong|that\s+is\s+wrong|incorrect|to\s+clarify)?|(?:correction|update|revision|amendment|fix):|(?:in\s+fact|as\s+a\s+matter\s+of\s+fact)[,\s]+|that'?s\s+incorrect|that\s+is\s+incorrect|that'?s\s+a\s+mistake)\s*[:\-,]?\s*(.+)/i,
      category: 'Correction',
    },
  ];

  for (const sentence of sentences) {
    for (const pattern of ruleIntentPatterns) {
      const match = sentence.match(pattern.triggerRegex);
      if (match && match[1]) {
        let ruleBody = match[1].trim();

        // Strip conversational trailing phrases
        ruleBody = ruleBody
          .replace(/(?:,\s*)?(?:please|thanks|thank\s+you|let\s+me\s+know|ok\??|okay\??|got\s+it\??)$/i, '')
          .trim();

        if (ruleBody.length >= 3) {
          // Clean quotes
          if ((ruleBody.startsWith('"') && ruleBody.endsWith('"')) || (ruleBody.startsWith("'") && ruleBody.endsWith("'"))) {
            ruleBody = ruleBody.slice(1, -1).trim();
          }

          let suggestedCategory = pattern.category;
          if (/threshold|limit|benchmark|sla|doa|variance|tat|cutoff|maximum|minimum/i.test(ruleBody) || /threshold|limit|sla/i.test(sentence)) {
            suggestedCategory = 'Threshold';
          } else if (/format|table|markdown|bullets|style|csv|excel|pdf|slide|tone|language|json/i.test(ruleBody)) {
            suggestedCategory = 'User Preference';
          } else if (/department|organization|company|filinvest|fbsc|team|unit|entity/i.test(ruleBody)) {
            suggestedCategory = 'Domain Context';
          } else if (/correction|wrong|mistake|fix|instead of|replace|update/i.test(ruleBody) && suggestedCategory !== 'Threshold') {
            suggestedCategory = 'Correction';
          }

          return {
            triggerPhrase: sentence,
            ruleContent: ruleBody,
            suggestedCategory,
          };
        }
      }
    }
  }

  return null;
}

// =====================================================================
// SECTION 2.1: TIER 1 — OPFS SHADOW LEDGER
// =====================================================================

const OPFS_DIR = 'sentinel_vault';
const OPFS_JOURNAL = 'journal.json';

export async function syncToOPFS(eventType: string, stateSnapshot: any): Promise<boolean> {
  try {
    if (typeof navigator === 'undefined' || !navigator.storage || !navigator.storage.getDirectory) {
      return false;
    }

    const root = await navigator.storage.getDirectory();
    const vaultDir = await root.getDirectoryHandle(OPFS_DIR, { create: true });
    const journalHandle = await vaultDir.getFileHandle(OPFS_JOURNAL, { create: true });

    let entries: any[] = [];
    try {
      const file = await journalHandle.getFile();
      if (file.size > 0) {
        const text = await file.text();
        entries = JSON.parse(text);
        if (!Array.isArray(entries)) entries = [];
      }
    } catch (_e) {
      entries = [];
    }

    const entry = {
      timestamp: new Date().toISOString(),
      type: eventType,
      snapshot: {
        timestamp: Date.now(),
        themeMode: localStorage.getItem('intel_theme') || 'system',
        folders: stateSnapshot.folders || [],
        sessions: (stateSnapshot.sessions || []).slice(0, 100), // serialize recent sessions
        agents: stateSnapshot.agents || [],
        exportConfig: stateSnapshot.exportConfig || null,
        workDeskDraft: stateSnapshot.workDeskDraft || null,
        memoryRecords: stateSnapshot.memoryRecords || [],
      },
    };

    entries.push(entry);
    // Maintain ring buffer of last 150 transaction snapshots
    if (entries.length > 150) {
      entries = entries.slice(-150);
    }

    const writable = await (journalHandle as any).createWritable();
    await writable.write(JSON.stringify(entries, null, 2));
    await writable.close();
    return true;
  } catch (err) {
    console.warn('[OPFS Shadow Ledger Sync]', err);
    return false;
  }
}

export async function getOPFSJournalSummary(): Promise<{ entryCount: number; lastTimestamp: string | null }> {
  try {
    if (!navigator.storage || !navigator.storage.getDirectory) {
      return { entryCount: 0, lastTimestamp: null };
    }
    const root = await navigator.storage.getDirectory();
    const vaultDir = await root.getDirectoryHandle(OPFS_DIR, { create: false });
    const journalHandle = await vaultDir.getFileHandle(OPFS_JOURNAL, { create: false });
    const file = await journalHandle.getFile();
    if (file.size === 0) return { entryCount: 0, lastTimestamp: null };
    const text = await file.text();
    const entries = JSON.parse(text);
    if (!Array.isArray(entries) || entries.length === 0) return { entryCount: 0, lastTimestamp: null };
    const last = entries[entries.length - 1];
    return {
      entryCount: entries.length,
      lastTimestamp: last?.timestamp || null,
    };
  } catch (_e) {
    return { entryCount: 0, lastTimestamp: null };
  }
}

/**
 * Automated Boot Recovery:
 * If IndexedDB/localStorage is empty (e.g. corporate IT cache wipe),
 * silently restore state from OPFS shadow ledger journal.
 */
export async function checkAndRecoverFromOPFS(): Promise<any | null> {
  try {
    if (!navigator.storage || !navigator.storage.getDirectory) return null;
    const root = await navigator.storage.getDirectory();
    let vaultDir: any;
    try {
      vaultDir = await root.getDirectoryHandle(OPFS_DIR, { create: false });
    } catch (_e) {
      return null;
    }

    const journalHandle = await vaultDir.getFileHandle(OPFS_JOURNAL, { create: false });
    const file = await journalHandle.getFile();
    if (file.size === 0) return null;

    const text = await file.text();
    const entries = JSON.parse(text);
    if (!Array.isArray(entries) || entries.length === 0) return null;

    const latest = entries[entries.length - 1];
    if (!latest || !latest.snapshot) return null;

    console.info('[Sentinel OPFS Recovery] Found shadow ledger with', entries.length, 'entries. Restoring latest state...');
    return latest.snapshot;
  } catch (err) {
    console.warn('[Sentinel OPFS Recovery Error]', err);
    return null;
  }
}

// =====================================================================
// SECTION 2.2: TIER 2 — HOT-DIRECTORY & FILE AUTO-SYNC (FILE SYSTEM ACCESS API)
// =====================================================================

export interface AutoSyncCapabilities {
  directoryPicker: boolean;
  saveFilePicker: boolean;
  openFilePicker: boolean;
}

export function isFileSystemAccessSupported(): AutoSyncCapabilities {
  if (typeof window === 'undefined') {
    return { directoryPicker: false, saveFilePicker: false, openFilePicker: false };
  }
  return {
    directoryPicker: typeof (window as any).showDirectoryPicker === 'function',
    saveFilePicker: typeof (window as any).showSaveFilePicker === 'function',
    openFilePicker: typeof (window as any).showOpenFilePicker === 'function',
  };
}

/**
 * Creates and saves `sentinel_vault_latest.json` in a directory selected by the user.
 * Uses showSaveFilePicker to prompt the native OS file dialog.
 */
export async function createAndSaveHotFileHandle(initialPayload?: any): Promise<{ handle: any; targetName: string; type: 'file' } | null> {
  if (typeof window === 'undefined' || !(window as any).showSaveFilePicker) {
    throw new Error('Native File Save API (showSaveFilePicker) is not supported in this browser.');
  }

  const fileHandle = await (window as any).showSaveFilePicker({
    suggestedName: 'sentinel_vault_latest.json',
    types: [
      {
        description: 'Sentinel Vault Shadow Mirror (.json)',
        accept: {
          'application/json': ['.json'],
        },
      },
    ],
  });

  if (!fileHandle) return null;

  hotDirHandleCached = fileHandle;
  const targetName = fileHandle.name || 'sentinel_vault_latest.json';

  // If initial payload provided, write it immediately
  if (initialPayload) {
    try {
      const writable = await fileHandle.createWritable();
      await writable.write(JSON.stringify(initialPayload, null, 2));
      await writable.close();
    } catch (writeErr) {
      console.warn('[Sentinel Auto-Sync] Initial write error on create:', writeErr);
    }
  }

  // Store in IndexedDB system_config
  const db = await openMemoryVaultDB();
  const tx = db.transaction(STORE_CONFIG, 'readwrite');
  const store = tx.objectStore(STORE_CONFIG);
  store.put({
    key: 'hot_directory_handle',
    handle: fileHandle,
    name: targetName,
    kind: 'file',
    updatedAt: new Date().toISOString(),
  });

  try {
    localStorage.setItem('sentinel_hot_dir_name', targetName);
    localStorage.setItem('sentinel_hot_dir_kind', 'file');
    localStorage.setItem('sentinel_hot_dir_active', 'true');
  } catch (_e) {}

  return { handle: fileHandle, targetName, type: 'file' };
}

/**
 * Connects to an existing `sentinel_vault_latest.json` file selected by the user.
 * Uses showOpenFilePicker to prompt the native OS file open dialog.
 */
export async function openExistingHotFileHandle(): Promise<{ handle: any; targetName: string; type: 'file' } | null> {
  if (typeof window === 'undefined' || !(window as any).showOpenFilePicker) {
    throw new Error('Native File Open API (showOpenFilePicker) is not supported in this browser.');
  }

  const [fileHandle] = await (window as any).showOpenFilePicker({
    multiple: false,
    types: [
      {
        description: 'Sentinel Vault Shadow Mirror (.json)',
        accept: {
          'application/json': ['.json'],
        },
      },
    ],
  });

  if (!fileHandle) return null;

  hotDirHandleCached = fileHandle;
  const targetName = fileHandle.name || 'sentinel_vault_latest.json';

  // Store in IndexedDB system_config
  const db = await openMemoryVaultDB();
  const tx = db.transaction(STORE_CONFIG, 'readwrite');
  const store = tx.objectStore(STORE_CONFIG);
  store.put({
    key: 'hot_directory_handle',
    handle: fileHandle,
    name: targetName,
    kind: 'file',
    updatedAt: new Date().toISOString(),
  });

  try {
    localStorage.setItem('sentinel_hot_dir_name', targetName);
    localStorage.setItem('sentinel_hot_dir_kind', 'file');
    localStorage.setItem('sentinel_hot_dir_active', 'true');
  } catch (_e) {}

  return { handle: fileHandle, targetName, type: 'file' };
}

/**
 * Maps an entire directory / folder using showDirectoryPicker or native Electron desktop dialog.
 */
export async function setHotDirectoryHandle(): Promise<{ handle: any; folderName: string; type: 'directory' } | null> {
  // Dual-mode desktop branch (Electron)
  if (typeof window !== 'undefined' && (window as any).sentinelDesktop?.selectDirectory) {
    const selectedPath = await (window as any).sentinelDesktop.selectDirectory();
    if (!selectedPath) return null;
    const folderName = selectedPath.split(/[\\/]/).pop() || selectedPath;
    try {
      localStorage.setItem('sentinel_desktop_sync_dir', selectedPath);
      localStorage.setItem('sentinel_hot_dir_name', folderName);
      localStorage.setItem('sentinel_hot_dir_kind', 'directory');
      localStorage.setItem('sentinel_hot_dir_active', 'true');
    } catch (_e) {}
    hotDirHandleCached = { kind: 'desktop_dir', path: selectedPath, name: folderName };
    return { handle: hotDirHandleCached, folderName, type: 'directory' };
  }

  if (typeof window === 'undefined' || !(window as any).showDirectoryPicker) {
    throw new Error('File System Access API (showDirectoryPicker) is not supported in this browser environment.');
  }

  const dirHandle = await (window as any).showDirectoryPicker({
    mode: 'readwrite',
    id: 'sentinel_hot_dir',
    startIn: 'documents',
  });

  if (!dirHandle) return null;

  hotDirHandleCached = dirHandle;
  const folderName = dirHandle.name || 'Mapped Folder';

  // Store in IndexedDB system_config
  const db = await openMemoryVaultDB();
  const tx = db.transaction(STORE_CONFIG, 'readwrite');
  const store = tx.objectStore(STORE_CONFIG);
  store.put({
    key: 'hot_directory_handle',
    handle: dirHandle,
    name: folderName,
    kind: 'directory',
    updatedAt: new Date().toISOString(),
  });

  try {
    localStorage.setItem('sentinel_hot_dir_name', folderName);
    localStorage.setItem('sentinel_hot_dir_kind', 'directory');
    localStorage.setItem('sentinel_hot_dir_active', 'true');
  } catch (_e) {}

  return { handle: dirHandle, folderName, type: 'directory' };
}

/**
 * Clears and disconnects active hot-directory or file handle.
 */
export async function clearHotDirectoryHandle(): Promise<void> {
  hotDirHandleCached = null;
  try {
    const db = await openMemoryVaultDB();
    const tx = db.transaction(STORE_CONFIG, 'readwrite');
    const store = tx.objectStore(STORE_CONFIG);
    store.delete('hot_directory_handle');
  } catch (_e) {}
  try {
    localStorage.removeItem('sentinel_hot_dir_name');
    localStorage.removeItem('sentinel_hot_dir_kind');
    localStorage.setItem('sentinel_hot_dir_active', 'false');
  } catch (_e) {}
}

export async function getHotDirectoryHandle(): Promise<any | null> {
  if (hotDirHandleCached) return hotDirHandleCached;

  try {
    const db = await openMemoryVaultDB();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_CONFIG, 'readonly');
      const store = tx.objectStore(STORE_CONFIG);
      const req = store.get('hot_directory_handle');
      req.onsuccess = () => {
        if (req.result && req.result.handle) {
          hotDirHandleCached = req.result.handle;
          resolve(req.result.handle);
        } else {
          resolve(null);
        }
      };
      req.onerror = () => resolve(null);
    });
  } catch (_e) {
    return null;
  }
}

/**
 * Unified synchronizer: writes `sentinel_vault_latest.json` directly to the mapped
 * FileSystemFileHandle or FileSystemDirectoryHandle.
 */
export async function syncToHotDirectory(vaultPayload: any): Promise<boolean> {
  try {
    // Dual-mode desktop branch (Electron)
    if (typeof window !== 'undefined' && (window as any).sentinelDesktop?.writeFile) {
      const baseDir = localStorage.getItem('sentinel_desktop_sync_dir') || '.';
      const targetPath = baseDir.replace(/[\\/]$/, '') + '/sentinel_vault_latest.json';
      const jsonString = JSON.stringify(vaultPayload, null, 2);
      await (window as any).sentinelDesktop.writeFile(targetPath, jsonString);
      console.info('[Sentinel Auto-Sync] Synced updated sentinel_vault_latest.json directly via Electron Desktop Bridge.');
      return true;
    }

    const handle = await getHotDirectoryHandle();
    if (!handle) return false;

    // Check / request readwrite permission
    if (typeof handle.queryPermission === 'function') {
      let perm = await handle.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') {
        if (typeof handle.requestPermission === 'function') {
          perm = await handle.requestPermission({ mode: 'readwrite' });
        }
        if (perm !== 'granted') {
          console.warn('[Sentinel Auto-Sync] Permission not granted to target handle.');
          return false;
        }
      }
    }

    const jsonString = JSON.stringify(vaultPayload, null, 2);

    // Case 1: Target handle is a direct File handle
    if (handle.kind === 'file' || (!handle.getFileHandle && typeof handle.createWritable === 'function')) {
      const writable = await handle.createWritable();
      await writable.write(jsonString);
      await writable.close();
      console.info('[Sentinel Auto-Sync] Synced updated sentinel_vault_latest.json directly to file handle.');
      return true;
    }

    // Case 2: Target handle is a Directory handle
    const fileHandle = await handle.getFileHandle('sentinel_vault_latest.json', { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(jsonString);
    await writable.close();

    console.info('[Sentinel Auto-Sync] Synced updated sentinel_vault_latest.json into directory handle.');
    return true;
  } catch (err) {
    console.warn('[Sentinel Hot-Directory Sync Error]', err);
    return false;
  }
}

/**
 * Direct file download fallback for browsers or frames that do not support File System Access API
 */
export function downloadFallbackSyncFile(vaultPayload: any): void {
  const blob = new Blob([JSON.stringify(vaultPayload, null, 2)], { type: 'application/json' });
  triggerBlobDownload(blob, 'sentinel_vault_latest.json');
}

// =====================================================================
// SECTION 2.3: TIER 3 — MODULAR COMPONENT EXPORTS
// =====================================================================

function triggerBlobDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

/**
 * 1. Full Disaster Recovery Snapshot (.json)
 */
export async function exportDisasterRecoverySnapshot(state: any) {
  const memories = await getPersistentMemories();
  const themeMode = localStorage.getItem('intel_theme') || localStorage.getItem('theme') || 'system';
  const lic = state.license || (function () {
    try {
      const raw = localStorage.getItem('intel_license');
      return raw ? JSON.parse(raw) : null;
    } catch (_e) {
      return null;
    }
  })();

  const docTitleInput = document.getElementById('workDeskDocTitle') as HTMLInputElement | null;
  const editor = document.getElementById('workDeskEditor') as HTMLTextAreaElement | null;

  const payload = {
    snapshotType: 'disaster_recovery',
    formatVersion: '2026.2',
    app: 'Sentinel Studio™ — Enterprise Capability Hub (ERI × EPEx)',
    architect: 'Patrick Jay F. Tanap, CPA, CICA',
    exportedAt: new Date().toISOString(),
    hardwareAuthHash: state.deviceFingerprint || lic?.signature || 'PJFT-SENTINEL-VERIFIED',
    license: lic,
    themePreference: themeMode,
    persistent_memory: memories,
    sessions: state.sessions || [],
    folders: state.folders || [],
    custom_agents: state.agents || [],
    exportConfig: state.exportConfig || null,
    workDeskDraft: {
      title: docTitleInput ? docTitleInput.value : 'Untitled Document',
      content: editor ? editor.value : '',
      mode: state.workDeskMode || 'word',
    },
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const filename = `Sentinel_Disaster_Recovery_Snapshot_${new Date().toISOString().slice(0, 10)}.json`;
  triggerBlobDownload(blob, filename);
}

/**
 * 2. Export Memory Ledger (.mem)
 */
export async function exportMemoryLedger() {
  const memories = await getPersistentMemories();
  const payload = {
    snapshotType: 'memory_ledger',
    formatVersion: '2026.2',
    app: 'Sentinel Studio™ — Persistent Memory Engine',
    exportedAt: new Date().toISOString(),
    totalRules: memories.length,
    persistent_memory: memories,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const filename = `Sentinel_Memory_Ledger_${new Date().toISOString().slice(0, 10)}.mem`;
  triggerBlobDownload(blob, filename);
}

/**
 * 3. Export Custom Agents (.agents)
 */
export function exportCustomAgents(agents: any[]) {
  const payload = {
    snapshotType: 'custom_agents',
    formatVersion: '2026.2',
    app: 'Sentinel Studio™ — Agent Studio Routine Catalog',
    exportedAt: new Date().toISOString(),
    totalAgents: agents.length,
    custom_agents: agents,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const filename = `Sentinel_Custom_Agents_${new Date().toISOString().slice(0, 10)}.agents`;
  triggerBlobDownload(blob, filename);
}

/**
 * 4. Export Chat History (.chats)
 */
export function exportChatHistory(sessions: any[], folders: any[]) {
  const payload = {
    snapshotType: 'chat_history',
    formatVersion: '2026.2',
    app: 'Sentinel Studio™ — Conversation Transcripts',
    exportedAt: new Date().toISOString(),
    totalSessions: sessions.length,
    sessions,
    folders,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const filename = `Sentinel_Chat_History_${new Date().toISOString().slice(0, 10)}.chats`;
  triggerBlobDownload(blob, filename);
}

// =====================================================================
// SECTION 2.4: TIER 4 — UNIVERSAL INTELLIGENT RESTORE ENGINE
// =====================================================================

export interface RestoreResult {
  success: boolean;
  type: 'disaster_recovery' | 'memory' | 'agents' | 'chats' | 'license' | 'unknown';
  message: string;
  data?: any;
  needsPrompt?: boolean;
}

export async function parseAndInspectRestorePayload(fileText: string): Promise<RestoreResult> {
  let parsed: any;
  try {
    parsed = JSON.parse(fileText);
  } catch (err: any) {
    return {
      success: false,
      type: 'unknown',
      message: `Invalid JSON format: ${err.message}`,
    };
  }

  // 1. License Key (.lic)
  if (parsed.status === 'approved' && (parsed.signature || parsed.licensee)) {
    return {
      success: true,
      type: 'license',
      message: `Hardware verification license for "${parsed.licensee || 'Authorized User'}" detected.`,
      data: parsed,
    };
  }

  // 2. Memory Ledger (.mem or payload containing persistent_memory only)
  if (parsed.snapshotType === 'memory_ledger' || (parsed.persistent_memory && !parsed.sessions && !parsed.custom_agents)) {
    return {
      success: true,
      type: 'memory',
      message: `Memory Ledger detected with ${(parsed.persistent_memory || []).length} rules.`,
      data: parsed.persistent_memory || [],
    };
  }

  // 3. Custom Agents (.agents or payload containing custom_agents only)
  if (parsed.snapshotType === 'custom_agents' || (parsed.custom_agents && !parsed.sessions && !parsed.persistent_memory)) {
    return {
      success: true,
      type: 'agents',
      message: `Agent Routine Catalog detected with ${(parsed.custom_agents || []).length} agents.`,
      data: parsed.custom_agents || [],
    };
  }

  // 4. Chat History (.chats or payload containing sessions only)
  if (parsed.snapshotType === 'chat_history' || (parsed.sessions && !parsed.persistent_memory && !parsed.custom_agents)) {
    return {
      success: true,
      type: 'chats',
      message: `Chat History detected with ${(parsed.sessions || []).length} conversations.`,
      data: {
        sessions: parsed.sessions || [],
        folders: parsed.folders || [],
      },
    };
  }

  // 5. Full Disaster Recovery Snapshot (contains persistent_memory AND (sessions or custom_agents))
  if (
    parsed.snapshotType === 'disaster_recovery' ||
    (parsed.persistent_memory && parsed.sessions) ||
    (parsed.sessions && parsed.custom_agents)
  ) {
    return {
      success: true,
      type: 'disaster_recovery',
      message: `Full Disaster Recovery Snapshot detected.`,
      data: parsed,
      needsPrompt: true,
    };
  }

  return {
    success: false,
    type: 'unknown',
    message: 'Unrecognized backup structure. Payload does not match Sentinel schema specifications.',
  };
}

export async function mergeMemoryRecords(incoming: MemoryRecord[]): Promise<number> {
  const db = await openMemoryVaultDB();
  const existing = await getPersistentMemories();
  const existingMap = new Map<string, MemoryRecord>();
  existing.forEach((m) => {
    existingMap.set(m.id, m);
    // Also map content normalized to avoid duplicates
    existingMap.set(m.content.trim().toLowerCase(), m);
  });

  let mergedCount = 0;
  const tx = db.transaction(STORE_MEMORIES, 'readwrite');
  const store = tx.objectStore(STORE_MEMORIES);

  const now = new Date().toISOString();
  for (const item of incoming) {
    if (!item.content) continue;
    const key = item.content.trim().toLowerCase();
    if (!existingMap.has(item.id) && !existingMap.has(key)) {
      const newRec: MemoryRecord = {
        ...item,
        id: item.id || `mem_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
        createdAt: item.createdAt || now,
        updatedAt: now,
      };
      store.put(newRec);
      mergedCount++;
    }
  }

  await new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });

  return mergedCount;
}

export async function overwriteAllMemories(incoming: MemoryRecord[]): Promise<void> {
  const db = await openMemoryVaultDB();
  const tx = db.transaction(STORE_MEMORIES, 'readwrite');
  const store = tx.objectStore(STORE_MEMORIES);
  store.clear();
  for (const item of incoming) {
    store.put(item);
  }
  await new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}
