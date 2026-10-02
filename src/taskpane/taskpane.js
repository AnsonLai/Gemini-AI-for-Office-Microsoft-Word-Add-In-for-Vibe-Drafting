import { requestGemini, requestGeminiJson, geminiEndpoint } from './modules/chat/gemini-client.js';
import { createLazyModuleLoader } from './modules/utils/lazy-module-loader.js';
import { postTaskpaneStartupProfile } from './modules/diagnostics/startup-profile-reporter.js';
/*
 * Gemini AI for Office - Task Pane Implementation
 * Author: Anson Lai
 * Location: Vancouver, Canada
 * Description: Word add-in integrating Google Gemini AI for document editing and analysis
 */

/* global document, Office, Word, localStorage, __TASKPANE_STARTUP_PROFILE__ */

import { marked } from 'marked';
import "./taskpane.css";

import {
  registerChatUiHandlers,
  setupScrollToBottom,
  createTypingIndicator,
  shakeInput,
  addMessageToChat,
  updateSystemMessage,
  addRetryButton,
  hideAllRetryButtons,
  removeMessage
} from './modules/chat/chat-ui.js';
import {
  maintainHistoryWindow,
  sanitizeHistory
} from './modules/chat/chat-history.js';
import { runDocumentAgentLoop } from './modules/chat/document-agent-loop.js';
import { createToolRegistry, getToolDeclarations } from './modules/chat/tool-registry.js';
import { buildCanonicalContextText } from './modules/chat/canonical-context.js';
import { getModelProfile } from './modules/config/model-profiles.js';
import {
  saveCheckpoint,
  getCheckpoint,
  importCheckpoints,
  formatAutoCheckpointLabel,
  migrateLegacyCheckpoints
} from './modules/storage/checkpoint-store.js';

const TASKPANE_STARTUP_PROFILE_ENABLED = __TASKPANE_STARTUP_PROFILE__ === true;
const TASKPANE_MODULE_EVALUATED_MS = typeof performance !== 'undefined' ? performance.now() : null;

// Configure marked for GFM (GitHub Flavored Markdown) with tables, breaks, etc.
marked.setOptions({
  gfm: true,           // Enable GitHub Flavored Markdown
  breaks: true,        // Convert \n to <br>
});

// ==================== CONFIGURATION CONSTANTS ====================

const RELEASE_MARKER = "v2.1.0.3";
const DEFAULT_AUTHOR = "Gemini AI";
const GLANCE_COLLAPSED_STORAGE_KEY = "glanceCollapsed";

globalThis.__GEMINI_TASKPANE_RELEASE__ = RELEASE_MARKER;

function collectSameOriginScriptResources() {
  const entries = performance.getEntriesByType?.('resource') || [];
  return entries.flatMap(entry => {
    try {
      const resourceUrl = new URL(entry.name, window.location.href);
      if (resourceUrl.origin !== window.location.origin || !/\.m?js$/i.test(resourceUrl.pathname)) return [];
      return [{
        name: resourceUrl.pathname,
        initiatorType: entry.initiatorType,
        durationMs: Number(entry.duration.toFixed(2)),
        encodedBodySize: Number(entry.encodedBodySize) || 0,
        transferSize: Number(entry.transferSize) || 0
      }];
    } catch {
      return [];
    }
  });
}

async function reportTaskpaneStartupProfile(info, officeReadyMs, usableUiMs) {
  if (!TASKPANE_STARTUP_PROFILE_ENABLED || typeof window === 'undefined' || typeof performance === 'undefined') return;

  const navigation = performance.getEntriesByType?.('navigation')?.[0];
  const scriptResources = collectSameOriginScriptResources();
  const initialScriptNames = new Set(scriptResources.map(resource => resource.name));
  const measuredBytes = property => scriptResources.reduce((total, resource) => total + resource[property], 0);
  const navigationMs = value => Number.isFinite(value) ? Number(value.toFixed(2)) : null;
  const record = {
    status: 'pending',
    scenario: 'automaticGlanceSuppressed',
    providerCalls: false,
    release: RELEASE_MARKER,
    host: info?.host || null,
    platform: Office?.context?.platform || null,
    navigationStartEpochMs: Number.isFinite(performance.timeOrigin) ? Math.round(performance.timeOrigin) : null,
    moduleEvaluatedMs: TASKPANE_MODULE_EVALUATED_MS,
    officeReadyMs,
    usableUiMs,
    domInteractiveMs: navigationMs(navigation?.domInteractive),
    domContentLoadedMs: navigationMs(navigation?.domContentLoadedEventEnd),
    loadEventEndMs: navigationMs(navigation?.loadEventEnd),
    initialScriptEncodedBytes: measuredBytes('encodedBodySize'),
    initialScriptTransferBytes: measuredBytes('transferSize'),
    scriptResources
  };

  try {
    if (info?.host !== Office.HostType.Word) throw new Error('Taskpane startup profile requires Word host.');
    const firstToolModuleLoadStart = performance.now();
    await loadAgenticTools();
    record.firstToolModuleLoadMs = Number((performance.now() - firstToolModuleLoadStart).toFixed(2));
    record.toolModuleInitialized = agenticToolsInitialized;
    record.lazyScriptResources = collectSameOriginScriptResources()
      .filter(resource => !initialScriptNames.has(resource.name));
    record.status = agenticToolsInitialized ? 'passed' : 'failed';
  } catch {
    record.status = 'failed';
    record.toolModuleInitialized = false;
    record.lazyScriptResources = collectSameOriginScriptResources()
      .filter(resource => !initialScriptNames.has(resource.name));
    record.toolModuleLoadFailed = true;
  }

  window.__TASKPANE_STARTUP_PROFILE__ = record;
  console.info('[TaskpaneStartupProfile]', record);

  void postTaskpaneStartupProfile(record);
}

// Safety settings for Gemini API (disable all safety blocks)
const SAFETY_SETTINGS_BLOCK_NONE = [
  { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
  { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
];

// Search and text limits
const SEARCH_LIMITS = {
  MAX_LENGTH: 100,           // Max search string length for comments/highlights
  MAX_LENGTH_MODIFY: 80,     // Max search string length for modify_text operations
  SUFFIX_LENGTH: 60,         // Suffix length for range expansion
  RETRY_LENGTH: 30           // Fallback shorter search length for retries
};

// Document processing limits
const DOCUMENT_LIMITS = {
  MAX_WORDS: 30000,          // Approx 40 pages, ~40k tokens
  MAX_LOOPS: 6,              // Maximum tool execution loops
  MAX_NO_PROGRESS_TOOL_LOOPS: 2, // Stop when the same mutation tool cycle keeps applying 0 changes
  TOKEN_MULTIPLIER: 1.33     // Words to tokens conversion factor
};

// Auto-checkpoint throttle: skip a new auto-snapshot if one was taken very
// recently (e.g. multiple tools in one turn should not each snapshot).
const CHECKPOINT_THROTTLE_MS = 15000;
let lastAutoCheckpointAt = 0;
let lastAutoCheckpointId = -1;
let checkpointMigrationDone = false;

// API generation limits
const API_LIMITS = {
  MAX_OUTPUT_TOKENS: 48000   // Maximum tokens for AI response output
};

// Timeout limits for API calls
const TIMEOUT_LIMITS = {
  FETCH_TIMEOUT_MS: 60000,        // 60s timeout per individual API call
  TOTAL_REQUEST_TIMEOUT_MS: 180000 // 3 min total timeout for entire request (including tool loops)
};

// Global abort controller for cancelling requests
let currentRequestController = null;
/**
 * Extracts enhanced document context with rich formatting metadata.
 * Returns an object with enhanced paragraph notation and section mapping.
 * 
 * Format: [P#|Style|ListInfo|TableInfo|SectionInfo] Text
 * Examples:
 *   [P1|Normal] Regular paragraph
 *   [P2|Heading1] Chapter heading
 *   [P3|ListNumber|L1:0|§] 1. Section header (starts section 1)
 *   [P4|Normal|§1] Body text belonging to section 1
 *   [P5|Normal|T:1,2] Table cell at row 1, column 2
 */
async function extractEnhancedDocumentContext(context) {
  const body = context.document.body;
  const paragraphs = body.paragraphs;

  // Load all relevant paragraph properties
  paragraphs.load("items");
  await context.sync();

  // Load detailed properties for each paragraph
  for (const para of paragraphs.items) {
    para.load("text, style, listItemOrNullObject, parentTableOrNullObject, parentTableCellOrNullObject");
  }
  await context.sync();

  // Load list details for paragraphs that are list items
  for (const para of paragraphs.items) {
    if (!para.listItemOrNullObject.isNullObject) {
      para.listItemOrNullObject.load("level, listString");
    }
    if (!para.parentTableCellOrNullObject.isNullObject) {
      para.parentTableCellOrNullObject.load("rowIndex, cellIndex");
    }
  }

  await context.sync();

  // Build enhanced paragraph data
  const enhancedParagraphs = [];
  let currentSection = null;      // Current section number (e.g., "1", "2")
  let currentSubSection = null;   // Current subsection (e.g., "1.1", "2.3")
  let sectionCounter = 0;         // Tracks top-level sections
  let lastListLevel = -1;         // Tracks list nesting level
  let sectionStack = [];          // Stack for tracking nested sections

  for (let i = 0; i < paragraphs.items.length; i++) {
    const para = paragraphs.items[i];
    const text = para.text || "";
    const style = para.style || "Normal";

    // Build metadata parts
    const metaParts = [style];

    // Check if paragraph is a list item
    let isListItem = false;
    let listLevel = -1;
    let listString = "";

    if (!para.listItemOrNullObject.isNullObject) {
      isListItem = true;
      listLevel = para.listItemOrNullObject.level || 0;
      listString = para.listItemOrNullObject.listString || "";

      // Determine list type from style name
      const isNumbered = style.toLowerCase().includes("number") ||
        style.toLowerCase().includes("list number") ||
        /^\d+[.)]/.test(listString);
      const listType = isNumbered ? "ListNumber" : "ListBullet";

      // Replace style with more specific list type
      metaParts[0] = listType;

      // Add list ID and level (using a simple counter-based ID)
      metaParts.push(`L:${listLevel}`);
    }

    // Check if paragraph is in a table
    let isInTable = false;
    if (!para.parentTableCellOrNullObject.isNullObject) {
      isInTable = true;
      const rowIndex = para.parentTableCellOrNullObject.rowIndex || 0;
      const cellIndex = para.parentTableCellOrNullObject.cellIndex || 0;
      metaParts.push(`T:${rowIndex},${cellIndex}`);
    }

    // Section detection for legal contract patterns
    let sectionMarker = "";

    if (isListItem && !isInTable) {
      // This list item could be a section header
      // Detect section headers: list items at level 0 or items that start new sections

      if (listLevel === 0) {
        // Top-level list item = new section
        sectionCounter++;
        currentSection = String(sectionCounter);
        currentSubSection = null;
        sectionStack = [currentSection];
        sectionMarker = "§";  // Mark as section header
        lastListLevel = listLevel;
      } else if (listLevel > lastListLevel) {
        // Nested list item = subsection
        const parentSection = sectionStack[sectionStack.length - 1] || currentSection;
        const subNum = sectionStack.length;
        currentSubSection = `${parentSection}.${listLevel}`;
        sectionStack.push(currentSubSection);
        sectionMarker = "§";  // Also mark as subsection header
        lastListLevel = listLevel;
      } else if (listLevel <= lastListLevel && listLevel > 0) {
        // Same or shallower nested level - pop stack and create new subsection
        while (sectionStack.length > listLevel + 1) {
          sectionStack.pop();
        }
        const parentSection = sectionStack[0] || currentSection;
        currentSubSection = `${parentSection}.${listLevel}`;
        sectionStack[listLevel] = currentSubSection;
        sectionMarker = "§";
        lastListLevel = listLevel;
      }

      if (sectionMarker) {
        metaParts.push(sectionMarker);
      }
    } else if (!isListItem && !isInTable && currentSection) {
      // Non-list paragraph following a section header = section body
      const belongsTo = currentSubSection || currentSection;
      metaParts.push(`§${belongsTo}`);
    }

    // Build the enhanced notation
    const metaString = metaParts.join("|");
    const enhancedLine = `[P${i + 1}|${metaString}] ${text}`;

    enhancedParagraphs.push({
      index: i + 1,
      text: text,
      style: style,
      isListItem: isListItem,
      listLevel: listLevel,
      isInTable: isInTable,
      section: currentSection,
      subSection: currentSubSection,
      isSectionHeader: sectionMarker === "§",
      meta: metaString,
      enhancedLine: enhancedLine
    });
  }

  return {
    paragraphs: enhancedParagraphs,
    formattedText: enhancedParagraphs.map(p => p.enhancedLine).join("\n"),
    sectionCount: sectionCounter
  };
}

/** Prefer the canonical (engine) paragraph text for the model view; fall back to Word text. */
function canonicalContextOr(enhancedParagraphs, sourceBaseline, wordFormattedText) {
  if (!enhancedParagraphs) return wordFormattedText;
  const canonicalText = buildCanonicalContextText(enhancedParagraphs, sourceBaseline);
  if (canonicalText == null) {
    const comparable = value => String(value ?? '').replace(/\s+/g, ' ').trim();
    const firstTextMismatch = enhancedParagraphs.findIndex((paragraph, offset) => comparable(sourceBaseline?.[offset]?.exactText) !== comparable(paragraph.text));
    console.warn('Canonical context unavailable: Word and source paragraphs do not align; using Word text.', {
      wordParagraphs: enhancedParagraphs.length,
      sourceParagraphs: Array.isArray(sourceBaseline) ? sourceBaseline.length : null,
      firstTextMismatch: firstTextMismatch < 0 ? null : firstTextMismatch + 1,
      extraSourceTail: Array.isArray(sourceBaseline)
        ? sourceBaseline.slice(enhancedParagraphs.length).map(paragraph => ({ index: paragraph.index, empty: !comparable(paragraph.exactText) }))
        : null
    });
    return wordFormattedText;
  }
  return canonicalText;
}

/** Capture a fresh model view and canonical source baseline for the next model turn. */
async function captureRefreshedDocumentContext(wordOperationSupportPromise) {
  let formattedText = '';
  let sourceBaseline = [];
  let enhancedParagraphs = null;

  await Word.run(async (context) => {
    const body = context.document.body;
    try {
      const enhancedContext = await extractEnhancedDocumentContext(context);
      formattedText = enhancedContext.formattedText;
      enhancedParagraphs = enhancedContext.paragraphs;
    } catch (error) {
      console.warn('Enhanced context refresh failed; using document text', error?.code || error?.name);
      body.load('text');
      await context.sync();
      formattedText = body.text || '';
    }

    const sourceOoxml = body.getOoxml();
    await context.sync();
    const [wordOperationRunner] = await wordOperationSupportPromise;
    sourceBaseline = wordOperationRunner.captureWordSourceBaseline(sourceOoxml.value);
  });

  if (!Array.isArray(sourceBaseline) || sourceBaseline.length === 0) {
    throw new Error('The current document source baseline is unavailable.');
  }
  return { formattedText: canonicalContextOr(enhancedParagraphs, sourceBaseline, formattedText), sourceBaseline };
}

let chatHistory = [];

Office.onReady((info) => {
  const officeReadyMs = TASKPANE_STARTUP_PROFILE_ENABLED && typeof performance !== 'undefined'
    ? Number(performance.now().toFixed(2))
    : null;
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    // Show main view by default
    showMainView();

    // Add event listener for the chat send button (Fast)
    document.getElementById("send-button").onclick = () => sendChatMessage('fast');

    // Add event listener for the THINK button (Slow)
    document.getElementById("think-button").onclick = () => sendChatMessage('slow');

    // Add Enter key support for chat (Shift+Enter for new line)
    document.getElementById("chat-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        if (e.shiftKey) {
          // Shift+Enter: New line (default behavior)
          return;
        }
        e.preventDefault();
        if (e.ctrlKey || e.metaKey) {
          // Ctrl+Enter or Cmd+Enter: Thinking chat (slow)
          sendChatMessage('slow');
        } else {
          // Enter: Regular chat (fast)
          sendChatMessage('fast');
        }
      }
    });

    // Add event listeners for settings UI
    document.getElementById("settings-button").onclick = showSettingsView;
    document.getElementById("save-api-key").onclick = saveApiKey;
    document.getElementById("back-to-main").onclick = showMainView;

    // Add event listener for refresh chat button
    document.getElementById("refresh-chat-button").onclick = refreshChat;

    // Add event listener for Glance refresh
    document.getElementById("refresh-glance-button").onclick = runGlanceChecks;
    document.getElementById("toggle-glance-button").onclick = () => {
      const container = document.getElementById("glance-container");
      if (!container) return;
      const shouldCollapse = !container.classList.contains("collapsed");
      saveGlanceCollapsedState(shouldCollapse);
      applyGlanceCollapsedState(shouldCollapse);
    };

    // Add event listener for Add Glance Card
    document.getElementById("add-glance-card-button").onclick = () => {
      const settings = loadGlanceSettings();
      settings.push({
        id: 'q' + Date.now(),
        title: 'New Question',
        question: 'What would you like to check?'
      });
      saveGlanceSettings(settings);
      renderGlanceSettings();
    };

    // Check for API key on load
    if (!loadApiKey()) {
      showWelcomeScreen();
    } else if (TASKPANE_STARTUP_PROFILE_ENABLED) {
      // Keep the normal Glance view visible while the local startup profile avoids provider calls.
      renderGlanceMain();
    } else {
      // Run Glance checks if key exists
      renderGlanceMain();
      runGlanceChecks();
    }

    // Accordion Event Listeners
    setupAccordion("glance-settings-header", "glance-settings-content");
    setupAccordion("advanced-settings-header", "advanced-settings-content");

    // Scroll-to-bottom button setup
    setupScrollToBottom();

    // Add event listener for refresh author button
    document.getElementById("refresh-author-button").onclick = async () => {
      const author = await fetchDocumentAuthor();
      if (author) {
        document.getElementById("redline-author-input").value = author;
        saveRedlineAuthor(author);
      }
    };

    // Add event listeners for Redline settings
    document.getElementById("redline-toggle").onchange = (e) => {
      saveRedlineSetting(e.target.checked);
    };

    document.getElementById("redline-author-input").oninput = (e) => {
      saveRedlineAuthor(e.target.value);
    };

    // Update checkpoint status on load (internal only now)
    // updateCheckpointStatus(); // UI removed, but we can keep tracking internally if needed, or just remove this call.
  }
  if (TASKPANE_STARTUP_PROFILE_ENABLED && typeof performance !== 'undefined') {
    void reportTaskpaneStartupProfile(info, officeReadyMs, Number(performance.now().toFixed(2)));
  }
});

function showWelcomeScreen() {
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.innerHTML = ""; // Clear existing messages

  const welcomeContainer = document.createElement("div");
  welcomeContainer.className = "welcome-container";

  welcomeContainer.innerHTML = `
    <div class="welcome-header">
      <h2>Get Started in 30 Seconds</h2>
    </div>
    <div class="welcome-step">
      <div class="step-number">1</div>
      <div class="step-content">
        <p>Go to <a href="https://aistudio.google.com/app/api-keys" target="_blank">Google AI Studio</a>.</p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">2</div>
      <div class="step-content">
        <p>Click <strong>Create API key</strong> (top left).</p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">3</div>
      <div class="step-content">
        <p>Select your project (or create new) and copy the key string starting with <code style="color: #ff0000ff;">AIza...</code></p>
      </div>
    </div>
    <div class="welcome-step">
      <div class="step-number">4</div>
      <div class="step-content">
        <p>Click the <strong>Gear Icon</strong> <span style="font-size: 1.2em;">&#9881;</span> at the top right corner to enter your key.</p>
      </div>
    </div>
    <div class="welcome-note">
      <p style="text-align: right;">The free tier is <em>plenty</em> for personal use.</p>
    </div>

    <hr class="welcome-divider">

    <div class="welcome-header">
      <h2 >Features</h2>
    </div>

    <div class="feature-explanation">
      <h3>Document Tools</h3>
      <p>Chat with an assistant who can access to tools that can <strong>edit text</strong>, <strong>search Google</strong>, <strong>highlight key info</strong>, and <strong>leave comments</strong>.  These tools allow the assistant to interact with your document naturally and help you with your tasks.</p>
    </div>

    <div class="feature-explanation">
      <h3>Glance Checks</h3>
      <p>Set up custom criteria (like <em>Grammar</em> or <em>Factual Accuracy</em>) to automatically check every document you open.  You can customize these questions in Settings.</p>
    </div>

    <div class="feature-explanation">
      <h3>System Prompts</h3>
      <p>Customize how the AI behaves. You can tell it to be a <em>Grade 10 student working on an English paper</em> or an <em>associate lawyer at a New York law firm specializing in contracts</em>.  Give it context and instructions you think would be helpful.</p>
    </div>

    <div class="feature-explanation">
      <h3>Model Choices</h3>
      <p><strong>Fast Model:</strong> This model is used for regular chats and is great for quick edits and simple questions.  It is fast and cheap.</p>
      <p><strong>Slow Model:</strong> This model is used when you select "Think".  It provides deep analysis and basic online searches.  It is slower and more expensive, but provides more thorough results.</p>
    </div>

    <div class="welcome-footer">
      <p><em>If you have any questions, please reach out to us at <a href="mailto:support@reference.legal">support@reference.legal</a>.</em></p>
    </div>
  `;

  chatMessages.appendChild(welcomeContainer);
}

// --- Settings & View Management ---

function switchView(hideId, showId) {
  const hideEl = document.getElementById(hideId);
  const showEl = document.getElementById(showId);

  if (!hideEl || !showEl) return;

  // Fade out current
  hideEl.classList.add("view-hidden");
  hideEl.classList.remove("view-container"); // Ensure it doesn't conflict

  setTimeout(() => {
    hideEl.style.display = "none";
    showEl.style.display = "block";

    // Force reflow
    void showEl.offsetWidth;

    // Fade in new
    showEl.classList.remove("view-hidden");
    showEl.classList.add("view-container");
  }, 200); // Match CSS transition speed
}

function showSettingsView() {
  document.getElementById("settings-button").style.display = "none";
  document.getElementById("refresh-chat-button").style.display = "none";

  switchView("main-view", "settings-view");

  // Load current key into input
  const currentKey = loadApiKey();
  if (currentKey) {
    document.getElementById("api-key-input").value = currentKey;
  }
  // Load current models
  const currentFastModel = loadModel('fast');
  if (currentFastModel) {
    document.getElementById("model-select-fast").value = currentFastModel;
  }
  const currentSlowModel = loadModel('slow');
  if (currentSlowModel) {
    document.getElementById("model-select-slow").value = currentSlowModel;
  }
  // Load current system message
  const currentSystemMessage = loadSystemMessage();
  if (currentSystemMessage) {
    document.getElementById("system-message-input").value = currentSystemMessage;
  }
  // Render Glance settings
  renderGlanceSettings();

  // Load redline setting
  const redlineEnabled = loadRedlineSetting();
  document.getElementById("redline-toggle").checked = redlineEnabled;

  // Load redline author setting
  const redlineAuthor = loadRedlineAuthor();
  document.getElementById("redline-author-input").value = redlineAuthor;
}

function showMainView() {
  document.getElementById("settings-button").style.display = "block";
  document.getElementById("refresh-chat-button").style.display = "block";

  switchView("settings-view", "main-view");

  renderGlanceMain();
}


function refreshChat() {
  // Cancel any ongoing request
  if (currentRequestController) {
    currentRequestController.abort();
    currentRequestController = null;
    console.log("Active request cancelled by refresh.");
  }

  // Immediately unlock UI (in case it was locked by an active request)
  const chatInput = document.getElementById("chat-input");
  const sendButton = document.getElementById("send-button");
  const thinkButton = document.getElementById("think-button");

  if (chatInput) {
    chatInput.disabled = false;
    chatInput.value = "";
    chatInput.focus();
  }
  if (sendButton) sendButton.disabled = false;
  if (thinkButton) thinkButton.disabled = false;

  // Clear chat history
  chatHistory = [];

  // Clear the chat messages UI
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.innerHTML = "";

  // Add the welcome message back
  const welcomeMessage = document.createElement("div");
  welcomeMessage.className = "chat-message system";
  welcomeMessage.textContent = "Welcome! Ask me to assist you in editing this document.";
  chatMessages.appendChild(welcomeMessage);

  // Add a system message confirming the refresh
  addMessageToChat("System", "Chat history cleared. Starting new conversation.");
}

function saveApiKey() {
  const apiKey = document.getElementById("api-key-input").value;
  const fastModel = document.getElementById("model-select-fast").value;
  const slowModel = document.getElementById("model-select-slow").value;
  const systemMessage = document.getElementById("system-message-input").value;
  const redlineEnabled = document.getElementById("redline-toggle").checked;
  const redlineAuthor = document.getElementById("redline-author-input").value;

  if (apiKey && apiKey.trim() !== "") {
    localStorage.setItem("geminiApiKey", apiKey);
    localStorage.setItem("geminiModelFast", fastModel);
    localStorage.setItem("geminiModelSlow", slowModel);
    localStorage.setItem("geminiSystemMessage", systemMessage);
    saveRedlineSetting(redlineEnabled);
    saveRedlineAuthor(redlineAuthor);
    // Glance settings are saved automatically on change
    showMainView();
    addMessageToChat("System", "Settings saved successfully.");
    // Re-run checks with new settings
    runGlanceChecks();
  } else {
    addMessageToChat("System", "API Key cannot be empty.");
  }
}

function loadApiKey() {
  // First check localStorage (user-provided key takes precedence)
  const storedKey = localStorage.getItem("geminiApiKey");
  if (storedKey && storedKey.trim() !== "") {
    return storedKey;
  }
}

function loadModel(type = 'fast') {
  const key = type === 'slow' ? "geminiModelSlow" : "geminiModelFast";
  const storedModel = localStorage.getItem(key);
  if (storedModel && storedModel.trim() !== "") {
    return storedModel;
  }
  // Defaults
  return type === 'slow' ? "gemini-2.5-pro" : "gemini-flash-latest";
}

function loadSystemMessage() {
  const storedMessage = localStorage.getItem("geminiSystemMessage");
  if (storedMessage && storedMessage.trim() !== "") {
    return storedMessage;
  }
  return "Example: You are assisting an undergraduate student with their academic paper. You must be specific, precise, and double-check all your advice and suggested changes. Maintain a cheerful and helpful tone.";
}

function loadRedlineSetting() {
  const storedSetting = localStorage.getItem("redlineEnabled");
  return storedSetting !== null ? storedSetting === "true" : true; // Default to true (enabled)
}

function saveRedlineSetting(enabled) {
  localStorage.setItem("redlineEnabled", enabled.toString());
}

function loadRedlineAuthor() {
  const storedAuthor = localStorage.getItem("redlineAuthor");
  if (storedAuthor && storedAuthor.trim() !== "") {
    return storedAuthor;
  }
  return DEFAULT_AUTHOR; // Unified default fallback
}

function saveRedlineAuthor(author) {
  if (author !== undefined && author !== null) {
    localStorage.setItem("redlineAuthor", author.toString());
  }
}

async function setChangeTrackingForAi(context, redlineEnabled, sourceLabel = "AI") {
  let originalMode = null;
  let changed = false;
  let available = false;

  try {
    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    available = true;
    originalMode = doc.changeTrackingMode;
    const desiredMode = redlineEnabled ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;

    if (originalMode !== desiredMode) {
      doc.changeTrackingMode = desiredMode;
      await context.sync();
      changed = true;
    }
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: unavailable`, error);
  }

  return { available, originalMode, changed };
}

async function restoreChangeTracking(context, trackingState, sourceLabel = "AI") {
  if (!trackingState || !trackingState.available || !trackingState.changed || trackingState.originalMode === null) {
    return;
  }

  try {
    context.document.changeTrackingMode = trackingState.originalMode;
    await context.sync();
  } catch (error) {
    console.warn(`[ChangeTracking] ${sourceLabel}: restore failed`, error);
  }
}

const agenticToolDependencies = {
  getRequestSignal: () => currentRequestController?.signal,
  loadApiKey,
  loadModel,
  loadSystemMessage,
  loadRedlineSetting,
  loadRedlineAuthor,
  setChangeTrackingForAi,
  restoreChangeTracking,
  SEARCH_LIMITS,
  SAFETY_SETTINGS_BLOCK_NONE,
  API_LIMITS
};
let agenticToolsInitialized = false;
const loadWordOperationSupport = createLazyModuleLoader(
  () => Promise.all([
    import('./modules/docx-redline-js-integration/word-operation-runner.js'),
    import('@ansonlai/docx-redline-js')
  ]),
  ([, docxLibrary]) => {
    docxLibrary.setPlatform(Office?.context?.platform);
  }
);
const loadAgenticTools = createLazyModuleLoader(
  async () => {
    const [module] = await Promise.all([
      import('./modules/commands/agentic-tools.js'),
      loadWordOperationSupport()
    ]);
    return module;
  },
  module => {
    module.initAgenticTools(agenticToolDependencies);
    agenticToolsInitialized = true;
  }
);

/**
 * Fetches the document's author from Word properties.
 */
async function fetchDocumentAuthor() {
  try {
    let author = "";
    await Word.run(async (context) => {
      const properties = context.document.properties;
      properties.load("lastAuthor, author");
      await context.sync();

      // Use lastAuthor if available, otherwise author
      author = properties.lastAuthor || properties.author || "";
    });
    return author;
  } catch (error) {
    console.warn("Could not fetch document author:", error);
    return "";
  }
}

function loadGlanceSettings() {
  const stored = localStorage.getItem("glanceSettings");
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch (e) {
      console.error("Error parsing glance settings", e);
    }
  }
  // Default fallback
  return [
    { id: 'q1', title: 'Grammar & Spelling', question: 'Are there any glaring spelling or grammatical issues?' },
    { id: 'q2', title: 'Factual Accuracy', question: 'Is this document factually accurate?' }
  ];
}

function saveGlanceSettings(settings) {
  localStorage.setItem("glanceSettings", JSON.stringify(settings));
}

function loadGlanceCollapsedState() {
  return localStorage.getItem(GLANCE_COLLAPSED_STORAGE_KEY) === "true";
}

function saveGlanceCollapsedState(isCollapsed) {
  localStorage.setItem(GLANCE_COLLAPSED_STORAGE_KEY, isCollapsed.toString());
}

function applyGlanceCollapsedState(isCollapsed = loadGlanceCollapsedState()) {
  const container = document.getElementById("glance-container");
  const toggleButton = document.getElementById("toggle-glance-button");
  if (!container || !toggleButton) return;

  container.classList.toggle("collapsed", isCollapsed);
  toggleButton.setAttribute("aria-expanded", (!isCollapsed).toString());
  toggleButton.setAttribute("title", isCollapsed ? "Show Glance results" : "Hide Glance results");
}

function setupAccordion(headerId, contentId) {
  const header = document.getElementById(headerId);
  const content = document.getElementById(contentId);

  if (header && content) {
    header.onclick = () => {
      const isOpen = content.classList.contains("open");

      if (isOpen) {
        content.classList.remove("open");
        header.classList.remove("active");
        // Wait for transition then hide (optional, but keep display block for anim)
        // We rely on max-height: 0 hiding it
      } else {
        content.classList.add("open");
        header.classList.add("active");
      }
    };
  }
}


function renderGlanceMain() {
  const list = document.getElementById("glance-list");
  const container = document.getElementById("glance-container");
  list.innerHTML = "";
  const settings = loadGlanceSettings();

  if (settings.length === 0) {
    if (container) container.style.display = "none";
    return;
  }

  if (container) container.style.display = "block";
  applyGlanceCollapsedState();

  settings.forEach(item => {
    const div = document.createElement("div");
    div.className = "glance-item";
    div.id = `glance-item-${item.id}`;
    div.innerHTML = `
      <div class="glance-header">
        <span id="glance-indicator-${item.id}" class="glance-indicator gray"></span>
        <span class="glance-title">${item.title}</span>
      </div>
      <p id="glance-summary-${item.id}" class="glance-summary">Waiting for analysis...</p>
    `;
    list.appendChild(div);
  });
}

function renderGlanceSettings() {
  const list = document.getElementById("glance-settings-list");
  list.innerHTML = "";
  const settings = loadGlanceSettings();

  settings.forEach((item, index) => {
    const card = document.createElement("div");
    card.className = "glance-settings-card";
    card.dataset.index = index;
    card.dataset.id = item.id;

    // Slimmer layout: Drag handle on left, inputs stacked but compact
    card.innerHTML = `
      <div class="glance-card-header-row">
        <input type="text" class="ms-TextField-field glance-title-input" value="${item.title}" placeholder="Title">
        <span class="drag-handle" title="Drag to reorder">☰</span>
        <button class="delete-card-btn" title="Delete">✕</button>
      </div>
      <textarea class="ms-TextField-field glance-question-input" placeholder="Question (e.g. Is the grammar correct?)" rows="2">${item.question}</textarea>
    `;

    // Event Listeners
    card.querySelector(".delete-card-btn").onclick = () => {
      settings.splice(index, 1);
      saveGlanceSettings(settings);
      renderGlanceSettings();
    };

    const titleInput = card.querySelector(".glance-title-input");
    titleInput.onchange = (e) => {
      settings[index].title = e.target.value;
      saveGlanceSettings(settings);
    };

    const questionInput = card.querySelector(".glance-question-input");
    questionInput.onchange = (e) => {
      settings[index].question = e.target.value;
      saveGlanceSettings(settings);
    };

    // Drag Events - Attach start/end to HANDLE only
    const handle = card.querySelector('.drag-handle');
    handle.draggable = true;
    handle.addEventListener('dragstart', handleDragStart);
    handle.addEventListener('dragend', handleDragEnd);

    // Drop targets are still the CARDS
    card.addEventListener('dragover', handleDragOver);
    card.addEventListener('drop', handleDrop);
    card.addEventListener('dragenter', handleDragEnter);
    card.addEventListener('dragleave', handleDragLeave);

    list.appendChild(card);
  });
}

// Drag and Drop Handlers
let dragSrcEl = null;

function handleDragStart(e) {
  const card = this.closest('.glance-settings-card');
  card.style.opacity = '0.4';
  dragSrcEl = card;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/html', card.innerHTML);
}

function handleDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  return false;
}

function handleDragToggleClass(e, addClass) {
  const card = e.target.closest('.glance-settings-card');
  if (card) {
    card.classList.toggle('over', addClass);
  }
}

function handleDragEnter(e) {
  handleDragToggleClass(e, true);
}

function handleDragLeave(e) {
  handleDragToggleClass(e, false);
}

function handleDrop(e) {
  e.stopPropagation();

  const targetCard = e.target.closest('.glance-settings-card');

  if (dragSrcEl !== targetCard && targetCard) {
    const list = document.getElementById("glance-settings-list");
    const items = Array.from(list.children);
    const srcIndex = items.indexOf(dragSrcEl);
    const destIndex = items.indexOf(targetCard);

    const settings = loadGlanceSettings();
    const [movedItem] = settings.splice(srcIndex, 1);
    settings.splice(destIndex, 0, movedItem);

    saveGlanceSettings(settings);
    renderGlanceSettings();
  }
  return false;
}

function handleDragEnd(e) {
  const card = this.closest('.glance-settings-card');
  if (card) card.style.opacity = '1';

  const items = document.querySelectorAll('.glance-settings-card');
  items.forEach(function (item) {
    item.classList.remove('over');
  });
}

async function runGlanceChecks() {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) return;

  const settings = loadGlanceSettings();
  if (settings.length === 0) return;

  // Update UI to showing loading
  settings.forEach(item => {
    const indicator = document.getElementById(`glance-indicator-${item.id}`);
    const summary = document.getElementById(`glance-summary-${item.id}`);
    if (indicator) indicator.className = "glance-indicator gray";
    if (summary) summary.innerText = "Checking...";
  });

  try {
    let docText = "";
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      docText = body.text;
    });

    const model = loadModel('fast'); // Use fast model for glance checks

    // Prepare prompt for dynamic checks
    let questionsPrompt = "";
    settings.forEach((item, index) => {
      questionsPrompt += `Question ${index + 1} (ID: "${item.id}"): ${item.question}\n`;
    });

    const prompt = `
      Analyze the following document text and answer the following questions.
      Return the result as a JSON object where keys are the Question IDs (e.g., "q1", "q2").
      For each question, provide:
      - "status": "green" (no issues/good), "yellow" (minor issues/caution), or "red" (major issues/bad).
      - "summary": A very brief summary (max 10 words).

      IMPORTANT: Return ONLY the JSON object. Do not include any markdown formatting (like \`\`\`json), conversational text, or explanations.

      Questions:
      ${questionsPrompt}

      Document Text:
      """${docText}""" 
    `;

    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      tools: [{ google_search: {} }],
      safetySettings: SAFETY_SETTINGS_BLOCK_NONE
    };

    const result = await requestGemini({ model, apiKey: geminiApiKey, payload, signal: currentRequestController?.signal, timeoutMs: TIMEOUT_LIMITS.FETCH_TIMEOUT_MS });
    const candidate = result.candidates[0];
    let text = candidate.content.parts[0].text;

    // Robust JSON Extraction: Find the first '{' and the last '}'
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      text = jsonMatch[0];
    } else {
      // Fallback cleanup if regex fails (though regex is preferred)
      text = text.replace(/^```json\s*/, "").replace(/^```\s*/, "").replace(/```$/, "").trim();
    }

    const json = JSON.parse(text);

    // Update UI
    settings.forEach(item => {
      const res = json[item.id];
      if (res) {
        const indicator = document.getElementById(`glance-indicator-${item.id}`);
        const summary = document.getElementById(`glance-summary-${item.id}`);
        if (indicator) {
          indicator.className = `glance-indicator ${res.status}`;
          // Add pulse animation
          indicator.classList.add("pulse");
          setTimeout(() => indicator.classList.remove("pulse"), 500);
        }
        if (summary) summary.innerText = res.summary;
      }
    });


  } catch (error) {
    console.error("Glance check failed:", error);
    settings.forEach(item => {
      const summary = document.getElementById(`glance-summary-${item.id}`);
      if (summary) summary.innerText = "Error running check.";
    });
  }
}

// --- Checkpoint Management (IndexedDB-backed; see modules/storage/checkpoint-store.js) ---

/**
 * One-time migration of any legacy localStorage checkpoints into IndexedDB.
 * Runs at most once per session; clears the legacy key afterward.
 */
async function ensureCheckpointMigration() {
  if (checkpointMigrationDone) return;
  checkpointMigrationDone = true; // set first so a failure doesn't retry every save
  try {
    const legacyJson = localStorage.getItem("docCheckpoints");
    if (!legacyJson) return;
    const legacy = JSON.parse(legacyJson);
    const records = migrateLegacyCheckpoints(legacy);
    if (records.length > 0) {
      await importCheckpoints(records);
      console.log(`Migrated ${records.length} checkpoint(s) from localStorage to IndexedDB.`);
    }
    localStorage.removeItem("docCheckpoints");
  } catch (e) {
    console.warn("Checkpoint migration skipped:", e);
  }
}

/**
 * Capture the current document state as a checkpoint.
 * @param {boolean} silent - suppress chat status messages
 * @param {string|null} toolName - when set, this is an auto-checkpoint taken
 *   before a mutating tool; it is labeled `auto:<toolName>:<ISO>` and throttled.
 * @returns {Promise<number>} the new checkpoint id (>=1), or -1 on failure/skip
 */
async function createCheckpoint(silent = false, toolName = null) {
  const isAuto = !!toolName;

  // Throttle auto-checkpoints so multi-tool turns don't snapshot repeatedly.
  // Returning the last auto id keeps per-message revert buttons pointing at a
  // valid recent pre-edit state.
  if (isAuto && (Date.now() - lastAutoCheckpointAt) < CHECKPOINT_THROTTLE_MS) {
    return lastAutoCheckpointId;
  }

  if (!silent) {
    addMessageToChat("System", "Saving checkpoint...");
  }

  try {
    await ensureCheckpointMigration();

    // 'ooxml.value' is a base64 string of the entire document body.
    let ooxmlValue = "";
    await Word.run(async (context) => {
      const ooxml = context.document.body.getOoxml();
      await context.sync();
      ooxmlValue = ooxml.value;
    });

    const label = isAuto ? formatAutoCheckpointLabel(toolName) : "manual";
    const id = await saveCheckpoint(label, ooxmlValue);

    if (isAuto) {
      lastAutoCheckpointAt = Date.now();
      lastAutoCheckpointId = id;
    }
    if (!silent) {
      addMessageToChat("System", "Checkpoint saved.");
    }
    return id;
  } catch (error) {
    // Never block the edit on a checkpoint failure.
    console.error("Error saving checkpoint:", error);
    if (!silent) {
      addMessageToChat("Error", `Could not save checkpoint. ${error.message}`);
    }
    return -1;
  }
}

async function restoreCheckpoint(id) {
  let record = null;
  try {
    record = await getCheckpoint(id);
  } catch (e) {
    console.error("Error loading checkpoint:", e);
  }

  if (!record || !record.ooxml) {
    addMessageToChat("Error", "Invalid or missing checkpoint.");
    return;
  }

  const msgElement = addMessageToChat("System", "Reverting to checkpoint...");

  try {
    await Word.run(async (context) => {
      // Disable Track Changes to avoid "Delete All + Insert All" redlines
      const doc = context.document;
      doc.load("changeTrackingMode");
      await context.sync();

      const originalMode = doc.changeTrackingMode;
      if (originalMode !== Word.ChangeTrackingMode.off) {
        doc.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
      }

      context.document.body.clear(); // Clear the current document body
      context.document.body.insertOoxml(record.ooxml, "Replace");
      await context.sync();

      // Restore the original track-changes mode after reverting.
      if (originalMode !== Word.ChangeTrackingMode.off) {
        doc.changeTrackingMode = originalMode;
        await context.sync();
      }

      updateSystemMessage(msgElement, "Reverted successfully.");
    });
  } catch (error) {
    console.error("Error reverting checkpoint:", error);
    updateSystemMessage(msgElement, "Error: Could not revert checkpoint.");
  }
}

registerChatUiHandlers({
  onCancelRequest: () => {
    if (currentRequestController) {
      currentRequestController.abort();
      console.log('User cancelled request');
    }
  },
  onRestoreCheckpoint: restoreCheckpoint
});

// --- Chat Feature ---

async function sendChatMessage(modelType = 'fast', messageOverride = null) {
  const chatInput = document.getElementById("chat-input");
  const sendButton = document.getElementById("send-button");
  const thinkButton = document.getElementById("think-button");
  const userMessage = messageOverride || chatInput.value;

  if (userMessage.trim() === "") {
    shakeInput();
    return;
  }

  // Hide any existing retry buttons since conversation is continuing
  hideAllRetryButtons();

  // Reset tool execution tracker for this request
  const requestRecords = [];

  // Sanitize history to remove any hanging function calls from interrupted sessions
  chatHistory = sanitizeHistory(chatHistory);

  // Set up abort controller for this request (allows user cancellation)
  const requestController = new AbortController();
  currentRequestController = requestController;
  const requestStartTime = Date.now();

  // Resolve the model profile up front so it is available in both the request
  // loop and the outer catch block (e.g. for the timeout/throttle message).
  const modelProfile = getModelProfile(loadModel(modelType));

  // Lock UI
  chatInput.disabled = true;
  sendButton.disabled = true;
  if (thinkButton) thinkButton.disabled = true;

  // Display user message
  addMessageToChat("User", userMessage);
  chatInput.value = "";

  // Show loading indicator with typing dots and cancel button (yellow for slow, teal for fast)
  const dotColor = modelType === 'slow' ? 'yellow' : 'teal';
  const loadingMsg = createTypingIndicator(dotColor, true); // true = include cancel button
  const chatMessages = document.getElementById("chat-messages");
  chatMessages.appendChild(loadingMsg);
  chatMessages.scrollTop = chatMessages.scrollHeight;




  try {
    // --- Get Document Context ---
    const wordOperationSupportPromise = loadWordOperationSupport();
    wordOperationSupportPromise.catch(() => {});
    let docText = "";
    let docSourceBaseline = [];
    let docComments = [];
    let docRedlines = [];
    let docSelection = "";

    await Word.run(async (context) => {
      const body = context.document.body;

      // --- STAGE 1: Critical Text Retrieval ---
      // Fetch current selection & basic text first
      const selection = context.document.getSelection();
      selection.load("text");

      // We'll try enhanced extraction first as it's the gold standard
      let enhancedParagraphs = null;
      try {
        const enhancedContext = await extractEnhancedDocumentContext(context);
        docText = enhancedContext.formattedText;
        enhancedParagraphs = enhancedContext.paragraphs;
        console.log(`Enhanced context extracted: ${enhancedContext.paragraphs.length} paragraphs`);
      } catch (enhancedError) {
        console.warn("Enhanced context failed, falling back to simple text", enhancedError);
        // Fallback
        body.load("text");
        await context.sync();
        docText = body.text;
      }

      // Keep an opaque canonical snapshot alongside the model's human-readable context.
      // Execution compares this baseline before preparing any proposed redline batch.
      try {
        const sourceOoxml = body.getOoxml();
        await context.sync();
        const [wordOperationRunner] = await wordOperationSupportPromise;
        docSourceBaseline = wordOperationRunner.captureWordSourceBaseline(sourceOoxml.value);
        // Show the model the same paragraph text that edits are applied to.
        docText = canonicalContextOr(enhancedParagraphs, docSourceBaseline, docText);
      } catch (error) {
        console.warn('Canonical document targeting context unavailable', error?.code || error?.name);
      }
      docSelection = selection.text;

      // Sync to ensure we captured text/selection before trying risky features
      await context.sync();

      // --- STAGE 2: Optional Rich Data (Comments/Redlines) ---
      // These are prone to failure in older Word versions or specific environments
      try {
        const isWordApi14 = Office.context.requirements.isSetSupported("WordApi", "1.4");
        const isWordApi16 = Office.context.requirements.isSetSupported("WordApi", "1.6");

        if (isWordApi14) {
          const comments = context.document.comments;
          comments.load("items/content, items/authorName, items/creationDate");

          let trackedChanges = null;
          if (isWordApi16) {
            try {
              trackedChanges = body.getTrackedChanges();
              trackedChanges.load("items/type, items/text, items/author, items/date");
            } catch (e) { console.warn("Tracked changes not supported (API available but failed)", e); }
          } else {
            console.log("Tracked changes not supported (WordApi 1.6 required)");
          }

          await context.sync(); // syncing specifically for comments/redlines

          // Process optional data
          if (comments && comments.items) {
            docComments = comments.items.map(c => `[Comment by ${c.authorName} on ${c.creationDate}]: ${c.content}`);
          }
          if (trackedChanges && trackedChanges.items) {
            docRedlines = trackedChanges.items.map(tc => `[${tc.type} by ${tc.author} on ${tc.date}]: "${tc.text}"`);
          }
        } else {
          console.log("Optional rich data (comments/redlines) not supported (WordApi 1.4 required)");
        }

      } catch (optionalDataError) {
        if (optionalDataError.name === "RichApi.Error" && optionalDataError.code === "ApiNotFound") {
          console.warn("Could not fetch comments or redlines (API not found despite support check), proceeding with text only.");
        } else {
          console.warn("Could not fetch comments or redlines (API error), proceeding with text only:", optionalDataError);
        }
      }

    });
    // --- Check Document Size ---
    if (currentRequestController !== requestController) return;
    const wordCount = docText.split(/\s+/).length;
    const estimatedTokens = Math.ceil(wordCount * DOCUMENT_LIMITS.TOKEN_MULTIPLIER);

    if (wordCount > DOCUMENT_LIMITS.MAX_WORDS) {
      removeMessage(loadingMsg);
      addMessageToChat("System", `Document is too large to process (approx. ${estimatedTokens} tokens). Please reduce the document size or select a smaller section.`);

      // Re-enable UI
      chatInput.disabled = false;
      sendButton.disabled = false;
      if (thinkButton) thinkButton.disabled = false;

      return;
    }

    // --- Call Gemini API ---
    const geminiApiKey = loadApiKey();
    if (!geminiApiKey) {
      removeMessage(loadingMsg);
      addMessageToChat("Error", "Please set your Gemini API key in the Settings (click the \u2699 icon in the top right).");
      return;
    }

    const geminiModel = loadModel(modelType);
    const apiUrl = geminiEndpoint(geminiModel, geminiApiKey);

    let contextString = "";
    if (docSelection && docSelection.trim() !== "") {
      contextString += `User Highlighted Text:\n"""${docSelection}"""\n\n`;
    }
    if (docText) {
      contextString += `Context from the current document:\n"""${docText}"""\n\n`;
    }
    if (docComments.length > 0) {
      contextString += `Comments in the document:\n${docComments.join("\n")}\n\n`;
    }
    if (docRedlines.length > 0) {
      contextString += `Tracked Changes (Redlines) in the document:\n${docRedlines.join("\n")}\n\n`;
    }

    const prompt = contextString
      ? `${contextString}User Question:\n${userMessage}`
      : userMessage;

    // Add to history
    chatHistory.push({ role: "user", parts: [{ text: prompt }] });

    // Maintain rolling window - but ensure we don't break function call/response pairs
    if (chatHistory.length > 10) {
      chatHistory = maintainHistoryWindow(chatHistory, 10);
    }

    const tools = getToolDeclarations();

    const systemInstruction = {
      parts: [
        {
          text: loadSystemMessage() + `\\n\\nDOCUMENT CONTEXT FORMAT:
The document content uses enhanced paragraph markers with formatting metadata:
- [P#|Style] - Normal paragraphs with their style (e.g., [P1|Normal], [P2|Heading1])
- [P#|ListNumber|L:level|§] - Numbered list item at nesting level, § means it's a section header
- [P#|ListBullet|L:level] - Bullet list item at nesting level
- [P#|Normal|§N] - Normal paragraph belonging to section N (follows a section header)
- [P#|Normal|T:row,col] - Paragraph inside a table cell at row,col position

IMPORTANT: The [P#] tags, [T:row,col] tags, and other metadata are for YOUR internal reasoning and tool usage only. 
NEVER reference "P14", "P15", "Paragraph 14", etc. in your response to the user. The user does not see these numbers and they will be confusing (especially for table cells which the user does not count as paragraphs).
Instead, refer to locations continuously and naturally, e.g., "I updated the Introduction," "I fixed the second item in the list," "I modified the table row," or "I updated the highlighted text."

TOOL SELECTION GUIDANCE:
- Use the [P#] identifiers when calling tools, but never speak them to the user.
- For simple text edits within a paragraph: use \`apply_redlines\`
- For editing contiguous lists (adding/removing/reordering items): prefer \`edit_list\` to preserve formatting
- For converting non-contiguous headers (like "1. PURPOSE", "2. DEFINITION" with body text between them) to a proper numbered list: use \`convert_headers_to_list\`
- For editing tables: prefer \`edit_table\` to preserve structure
- For editing legal contract sections (numbered headers + body paragraphs): prefer \`edit_section\`
- For word count, character count, or text statistics of selection/paragraphs: use \`get_selection_stats\`
- The § marker indicates section structure - paragraphs marked §N belong to section N

IMPORTANT: You have access to tools. You can chat and respond normally to questions. However, when the user asks for an action that involves manipulating the document, you should HEAVILY FAVOR using the corresponding tool rather than just describing the action.

CRITICAL: For plain text edits and inline formatting within existing paragraphs, use \`apply_redlines\`. For structural list/table/section edits, use the dedicated tools (\`edit_list\`, \`convert_headers_to_list\`, \`edit_table\`, \`edit_section\`).
CRITICAL: When the user asks to convert normal paragraphs into a new table, use \`apply_redlines\`, not \`edit_table\` (which is only for existing tables marked with T:row,col).
CRITICAL: For text-to-table conversions, your \`apply_redlines\` instruction MUST say which full paragraph range is being replaced and MUST require a complete multiline GitHub Markdown table. Example: "Replace P4 through P6 with this two-column markdown table: | Disclosing Party | Receiving Party |\\n|---|---|\\n| [Name of Disclosing Party] | [Name of Receiving Party] |\\n| [Address of Disclosing Party] | [Address of Receiving Party] |". Never request a single pipe-delimited line.
CRITICAL: If the user asks to "Reply to a comment" by "changing textual content", you MUST call BOTH \`apply_redlines\` (to apply the text change) AND \`insert_comment\` (to insert the reply). Call them in the same turn.
NEVER claim to have "added a sentence" or "changed text" if you have only called \`insert_comment\`.
NEVER state that you have taken an action unless you have successfully invoked the corresponding tool.

AFTER executing a tool, DO NOT repeat the content of the document or the changes in your text response. The user can see the changes in the document.

CRITICAL: Do NOT use internal paragraph markers (like [P#] or P#) or internal IDs in your text responses to the user. These are for your internal reasoning and tool calls only. Refer to locations naturally (e.g., "the second paragraph", "the Definitions section", "the paragraph regarding termination").
 
 LIST HANDLING:
 When adding or modifying lists via \`apply_redlines\`, you MUST use specific Markdown syntax so the engine can format them correctly in Word:
 - Unordered: Use '* ' (asterisk space).
 - Ordered: Use '1. ', 'a. ', 'i. ', etc.
 - Multi-level / Outlines: Use exact numbering like '1.1.', '1.1.1.' or '2.1. ' if that is the intent.
 - Indentation: Sub-items MUST be indented by 4 spaces.
 - Do NOT use generic bullets ('-') if you want specific numbering. The engine relies on your markers (e.g., '1.1.') to detect the list type.

 TABLE HANDLING:
 - Existing table edits: use \`edit_table\` only when the target paragraph has a T:row,col marker.
 - New table creation from normal text: use \`apply_redlines\`.
 - In \`apply_redlines\` instructions for tables, name the contiguous source range and require multiline markdown table syntax.
 - Correct table shape: "| Header A | Header B |\\n|---|---|\\n| Cell A | Cell B |".
 - Incorrect table shape: "Header A|Header B|Cell A|Cell B".
 - If source blocks have multiple lines, preserve them as additional table rows. Do not use HTML tags such as <br> inside table cells.`,
        },
      ],
    };

    // The portable loop owns exchanges and stop decisions; this adapter owns UI/Word I/O.
    const loopState = { history: chatHistory, docText, docSourceBaseline, records: requestRecords };
    try {
      await runDocumentAgentLoop({
        state: loopState,
        registry: createToolRegistry(),
        loadRegistry: async () => {
          const toolsModule = await loadAgenticTools();
          return createToolRegistry(toolsModule.createAgenticTools({
            ...agenticToolDependencies, getRequestSignal: () => requestController.signal,
          }));
        },
        requestModel: history => callGeminiWithRetry(apiUrl, {
          contents: history,
          systemInstruction,
          tools,
          safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
          generationConfig: { maxOutputTokens: modelProfile.maxOutputTokens },
        }, modelProfile.retries, 1000, requestController.signal),
        refreshContext: () => captureRefreshedDocumentContext(wordOperationSupportPromise),
        createCheckpoint,
        signal: requestController.signal,
        requestStartTime,
        maxLoops: DOCUMENT_LIMITS.MAX_LOOPS,
        maxNoProgressToolLoops: DOCUMENT_LIMITS.MAX_NO_PROGRESS_TOOL_LOOPS,
        totalTimeoutMs: TIMEOUT_LIMITS.TOTAL_REQUEST_TIMEOUT_MS,
        onEvent(event) {
          if (currentRequestController !== requestController) return;
          if (event.type === 'turn_start') {
            console.log(`Starting chat loop iteration ${event.turn}`);
          } else if (event.type === 'call_recovered') {
            console.log('Recovered malformed tool call:', event.name);
          } else if (event.type === 'tool_start') {
            updateSystemMessage(loadingMsg, event.progressText);
          } else if (event.type === 'tool_end') {
            if (event.display.show) {
              updateSystemMessage(loadingMsg, event.completionText,
                event.completionWithCheckpoint ? event.checkpointIndex : undefined);
            }
            const messages = document.getElementById('chat-messages');
            if (loadingMsg && messages) messages.appendChild(loadingMsg);
          } else if (event.type === 'response') {
            if (requestRecords.length === 0) removeMessage(loadingMsg);
            addMessageToChat('Gemini', event.text);
          } else if (event.type === 'stop') {
            if (event.reason === 'cancelled') {
              removeMessage(loadingMsg);
              addMessageToChat('System', event.message);
            } else {
              const throttleWarning = event.reason === 'timeout' && modelProfile.previewThrottleWarning
                ? '\n\nThis model is in preview and your access has likely been throttled. Please go into settings and revert to Gemini 2.5.'
                : '';
              updateSystemMessage(loadingMsg, event.message + throttleWarning);
            }
          } else if (event.type === 'mutation_failure') {
            console.warn(`[LoopGuard] No-progress mutation loop ${event.failures}/${DOCUMENT_LIMITS.MAX_NO_PROGRESS_TOOL_LOOPS}`);
          }
        },
      });
    } finally {
      // Preserve actual outcomes even when the provider or an injected dependency throws.
      if (currentRequestController === requestController) chatHistory = loopState.history;
    }

  } catch (error) {
    if (currentRequestController !== requestController) return;
    console.error("Error calling Gemini API:", error);

    // Handle user cancellation specifically
    if (error.message === 'Request cancelled by user') {
      removeMessage(loadingMsg);
      addMessageToChat("System", "Request cancelled.");
    } else {
      // Only remove loadingMsg if no tools were executed (meaning it's still a "Thinking" message)
      if (requestRecords.length === 0) {
        removeMessage(loadingMsg);

        // Cleanup history for failed requests (timeout or error)
        if (chatHistory.length > 0 && chatHistory[chatHistory.length - 1].role === "user") {
          console.log("Discarding failed request from history");
          chatHistory.pop();
        }
      }

      let errorMessage = error.message ? `Sorry, I couldn't get a response. Error: ${error.message}` : `Sorry, I couldn't get a response. Error: ${String(error)}`;

      // Override error message for timeouts. Only suggest reverting to 2.5 for
      // models flagged as preview/throttled.
      if (error.message && (error.message.includes("timed out") || error.message.includes("timeout"))) {
        errorMessage = modelProfile.previewThrottleWarning
          ? "This model is in preview and has likely been throttled. Please go into settings and revert to Gemini 2.5."
          : "The request timed out. The AI is taking longer than usual. Please try again.";
      }

      const errorMsgEl = addMessageToChat("Error", errorMessage);

      // Add retry button if it's the specific missing content error
      if (error.message && error.message.includes("Gemini response was missing content.parts")) {
        addRetryButton(errorMsgEl, userMessage);
      }
    }
  } finally {
    // A refreshed/superseded request cannot overwrite the next request's UI.
    if (currentRequestController === requestController) {
      currentRequestController = null;
      chatInput.disabled = false;
      sendButton.disabled = false;
      if (thinkButton) thinkButton.disabled = false;
      chatInput.focus();
    }
  }
}

// Retries provider transport only. Tool execution remains in the caller.
async function callGeminiWithRetry(url, payload, retries = 3, backoff = 1000, signal = currentRequestController?.signal) {
  return requestGeminiJson(url, payload, {
    signal,
    timeoutMs: TIMEOUT_LIMITS.FETCH_TIMEOUT_MS,
    maxAttempts: retries,
    backoffMs: backoff,
  });
}
