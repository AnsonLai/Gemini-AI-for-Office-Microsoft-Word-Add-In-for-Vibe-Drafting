import { requestGemini, requestGeminiJson, geminiEndpoint } from './modules/chat/gemini-client.js';
import { populateModelSetting, readModelSetting } from './modules/config/model-settings.js';
import { saveApiKeySetting, saveModelSettingFrom, saveSystemMessageSetting } from './modules/config/settings-save.js';
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
  addConfirmButtons,
  hideAllRetryButtons,
  removeCancelButton,
  removeMessage
} from './modules/chat/chat-ui.js';
import {
  maintainHistoryWindow,
  sanitizeHistory
} from './modules/chat/chat-history.js';
import { runDocumentAgentLoop } from './modules/chat/document-agent-loop.js';
import { landedEditMessage, summarizeRequestOutcome, waitForSettle, wroteOrAttempted } from './modules/chat/request-outcome.js';
import { createToolRegistry, getToolDeclarations } from './modules/chat/tool-registry.js';
import { createWordDocumentSourceReader } from './modules/chat/word-document-source.js';
import { createWebSearch } from './modules/chat/web-search.js';
import { DOCUMENT_EDITING_GUIDE, summarizeDocumentSource } from './modules/chat/document-editing-guide.js';
import { getModelProfile } from './modules/config/model-profiles.js';
import { glanceResponseText, parseGlanceResults } from './modules/glance/glance-results.js';
import {
  saveCheckpoint,
  getCheckpoint,
  updateCheckpoint,
  pruneCheckpoints,
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

const RELEASE_MARKER = "v2.3.0.1";
const DEFAULT_AUTHOR = "Gemini AI";
const GLANCE_COLLAPSED_STORAGE_KEY = "glanceCollapsed";
// Off by default: an automatic check sends the whole document to Gemini unasked.
const GLANCE_AUTORUN_STORAGE_KEY = "glanceAutoRun";

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
    await loadDocumentTools();
    record.firstToolModuleLoadMs = Number((performance.now() - firstToolModuleLoadStart).toFixed(2));
    record.toolModuleInitialized = documentToolsLoaded;
    record.lazyScriptResources = collectSameOriginScriptResources()
      .filter(resource => !initialScriptNames.has(resource.name));
    record.status = documentToolsLoaded ? 'passed' : 'failed';
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

// Document processing limits
const DOCUMENT_LIMITS = {
  MAX_WORDS: 30000,          // Approx 40 pages, ~40k tokens
  MAX_LOOPS: 6,              // Maximum tool execution loops
  MAX_NO_PROGRESS_TOOL_LOOPS: 2, // Stop when the same mutation tool cycle keeps applying 0 changes
  TOKEN_MULTIPLIER: 1.33     // Words to tokens conversion factor
};

let checkpointMigrationDone = false;

// Timeout limits for API calls
const TIMEOUT_LIMITS = {
  // Thinking models can take over a minute on one turn (KB-13).
  FETCH_TIMEOUT_MS: 120000,       // 2 min timeout per individual API call
  TOTAL_REQUEST_TIMEOUT_MS: 300000, // 5 min total timeout for entire request (including tool loops)
  REFRESH_SETTLE_TIMEOUT_MS: 15000  // How long refresh waits for a cancelled request to finish
};

// Global abort controller for cancelling requests
let currentRequestController = null;
// Settles when the current chat request has finished, including any Word write.
// Its value is { wrote, checkpointIndex }: whether an edit landed and the Revert target.
let currentRequestSettled = Promise.resolve();
let revertInProgress = false;
// The document URL can change when Word reopens/switches documents. Fresh part
// tokens remain authoritative for edits/Undo within this taskpane instance.
const documentInstanceId = `${Date.now()}-${Math.random()}`;
const getDocumentId = () => `${documentInstanceId}:${Office.context.document.url || ''}`;
async function captureRefreshedDocumentContext(readSource) {
  const source = await readSource();
  return { formattedText: summarizeDocumentSource(source) };
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
    document.getElementById("back-to-main").onclick = leaveSettings;

    // Every Settings field saves as soon as it changes.
    document.getElementById("api-key-input").addEventListener("change", onApiKeyChange);
    document.getElementById("system-message-input").addEventListener("change", onSystemMessageChange);
    for (const type of ['fast', 'slow']) {
      document.getElementById(`model-select-${type}`).addEventListener("change", () => onModelChange(type, true));
      document.getElementById(`model-custom-${type}`).addEventListener("change", () => onModelChange(type, false));
    }

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

    document.getElementById("glance-autorun-toggle").onchange = (e) => {
      saveGlanceAutoRun(e.target.checked);
    };

    // Check for API key on load
    if (!loadApiKey()) {
      showWelcomeScreen();
    } else {
      renderGlanceMain();
      // The local startup profile never makes provider calls.
      if (loadGlanceAutoRun() && !TASKPANE_STARTUP_PROFILE_ENABLED) runGlanceChecks();
    }

    // Drop other panes' stale checkpoints (full document copies) early.
    void ensureCheckpointMigration();

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
      <p>Chat with an assistant that can <strong>edit text</strong> as tracked changes, <strong>format and highlight</strong> key info, <strong>leave comments</strong>, and <strong>search the web</strong> with Google when a question needs facts the document does not contain.</p>
    </div>

    <div class="feature-explanation">
      <h3>Glance Checks</h3>
      <p>Set up custom criteria (like <em>Grammar</em> or <em>Factual Accuracy</em>) and check a document with one click. You can customize these questions in Settings, and choose to run them automatically when the pane opens.</p>
    </div>

    <div class="feature-explanation">
      <h3>System Prompts</h3>
      <p>Customize how the AI behaves. You can tell it to be a <em>Grade 10 student working on an English paper</em> or an <em>associate lawyer at a New York law firm specializing in contracts</em>.  Give it context and instructions you think would be helpful.</p>
    </div>

    <div class="feature-explanation">
      <h3>Model Choices</h3>
      <p><strong>Fast Model:</strong> This model is used for regular chats and is great for quick edits and simple questions.  It is fast and cheap.</p>
      <p><strong>Slow Model:</strong> This model is used when you select "Think". It reasons more deeply, so it is slower and more expensive, but gives more thorough results. Both models can search the web.</p>
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

// The settings that Glance depends on, as they were when Settings opened.
let settingsOnOpen = null;
const settingsSnapshot = () => JSON.stringify([
  loadApiKey() || '', localStorage.getItem("glanceSettings"), loadGlanceAutoRun(),
  loadModel('fast'), loadModel('slow'), loadSystemMessage()
]);

function showSettingsView() {
  settingsOnOpen = settingsSnapshot();
  document.getElementById("settings-button").style.display = "none";
  document.getElementById("refresh-chat-button").style.display = "none";
  showSettingsStatus();

  switchView("main-view", "settings-view");

  // Load current key into input
  const currentKey = loadApiKey();
  if (currentKey) {
    document.getElementById("api-key-input").value = currentKey;
  }
  // Load current models
  const currentFastModel = loadModel('fast');
  if (currentFastModel) {
    populateModelSetting('fast', currentFastModel);
  }
  const currentSlowModel = loadModel('slow');
  if (currentSlowModel) {
    populateModelSetting('slow', currentSlowModel);
  }
  // Load current system message
  const currentSystemMessage = loadSystemMessage();
  if (currentSystemMessage) {
    document.getElementById("system-message-input").value = currentSystemMessage;
  }
  // Render Glance settings
  renderGlanceSettings();
  document.getElementById("glance-autorun-toggle").checked = loadGlanceAutoRun();

  // Load redline setting
  const redlineEnabled = loadRedlineSetting();
  document.getElementById("redline-toggle").checked = redlineEnabled;

  // Load redline author setting
  const redlineAuthor = loadRedlineAuthor();
  document.getElementById("redline-author-input").value = redlineAuthor;
}

// Fields have already saved; leaving only returns to the main view and refreshes what depends on them.
function leaveSettings() {
  const changed = settingsOnOpen !== null && settingsSnapshot() !== settingsOnOpen;
  settingsOnOpen = null;
  showSettingsStatus();
  showMainView();
  // With the key removed, an empty chat goes back to the first-run welcome.
  if (!loadApiKey() && !document.querySelector("#chat-messages .chat-message")) showWelcomeScreen();
  if (changed && loadGlanceAutoRun()) runGlanceChecks();
}

function showMainView() {
  document.getElementById("settings-button").style.display = "block";
  document.getElementById("refresh-chat-button").style.display = "block";

  switchView("settings-view", "main-view");

  renderGlanceMain();
}


async function refreshChat() {
  // A revert can't be stopped once Word is writing, and its outcome must be shown.
  if (revertInProgress) {
    addMessageToChat("System", "Wait for the revert to finish before clearing the chat.");
    return;
  }
  // Cancel any ongoing request and wait until it has settled: a Word write
  // already committing can't be stopped, and nothing may start alongside it.
  let cancelledOutcome = null;
  if (currentRequestController) {
    const settled = currentRequestSettled;
    const refreshButton = document.getElementById("refresh-chat-button");
    currentRequestController.abort();
    console.log("Active request cancelled by refresh.");
    if (refreshButton) refreshButton.disabled = true;
    let settle;
    try {
      settle = await waitForSettle(settled, TIMEOUT_LIMITS.REFRESH_SETTLE_TIMEOUT_MS);
    } finally {
      if (refreshButton) refreshButton.disabled = false;
    }
    if (settle.timedOut) {
      // Word may still be writing: keep the chat and the locked inputs, and let the request finish itself.
      addMessageToChat("System", "Word is still finishing the last request. The chat was not cleared; try again in a moment, or close and reopen the pane if this continues.");
      return;
    }
    cancelledOutcome = settle.value;
  }

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
  // The cancelled request's receipt is gone with the old chat, so say that its edit landed.
  if (cancelledOutcome?.wrote) {
    addMessageToChat("System", landedEditMessage(cancelledOutcome.checkpointIndex), cancelledOutcome.checkpointIndex);
  }
}

// Settings messages show inside Settings: the chat is hidden while it is open.
let settingsStatusTimer = null;
function showSettingsStatus(message = "", { saved = false } = {}) {
  const status = document.getElementById("settings-status");
  if (!status) return;
  clearTimeout(settingsStatusTimer);
  status.textContent = message;
  status.hidden = !message;
  status.classList.toggle("saved", saved);
  // A confirmation fades after a moment; an error stays until the next change.
  if (message && saved) settingsStatusTimer = setTimeout(() => showSettingsStatus(), 2000);
}

// An empty API key removes the stored key, which is the no-key state.
function onApiKeyChange(e) {
  const { stored } = saveApiKeySetting(localStorage, e.target.value);
  showSettingsStatus(stored ? "Saved." : "API key removed.", { saved: true });
}

function onSystemMessageChange(e) {
  saveSystemMessageSetting(localStorage, e.target.value);
  showSettingsStatus("Saved.", { saved: true });
}

function onModelChange(type, fromSelect) {
  // Choosing "Custom model ID..." waits for the id to be typed.
  if (fromSelect && document.getElementById(`model-select-${type}`).value === 'custom'
    && !document.getElementById(`model-custom-${type}`).value.trim()) return;
  const result = saveModelSettingFrom(localStorage, type, readModelSetting);
  if (result.ok) showSettingsStatus("Saved.", { saved: true });
  else showSettingsStatus(result.error);
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
  return type === 'slow' ? "gemini-4-carbon" : "gemini-flash-latest";
}

function loadSystemMessage() {
  const storedMessage = localStorage.getItem("geminiSystemMessage");
  if (storedMessage && storedMessage.trim() !== "") {
    return storedMessage;
  }
  return "Assist with precise professional document wording. Preserve meaning and existing presentation unless the user requests a supported change.";
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

let documentToolsLoaded = false;
const loadWordOperationSupport = createLazyModuleLoader(
  () => Promise.all([
    import('./modules/docx-redline-js-integration/word-operation-runner.js'),
    import('@ansonlai/docx-redline-js'),
    import('./modules/chat/document-inspection-session.js'),
    import('./modules/docx-redline-js-integration/word-source-identity.js')
  ]),
  ([, docxLibrary]) => {
    docxLibrary.setPlatform(Office?.context?.platform);
  }
);
const loadDocumentTools = createLazyModuleLoader(
  async () => {
    const [module] = await Promise.all([
      import('./modules/chat/document-tools.js'),
      loadWordOperationSupport()
    ]);
    return module;
  },
  module => {
    documentToolsLoaded = true;
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

function loadGlanceAutoRun() {
  return localStorage.getItem(GLANCE_AUTORUN_STORAGE_KEY) === "true";
}

function saveGlanceAutoRun(enabled) {
  localStorage.setItem(GLANCE_AUTORUN_STORAGE_KEY, enabled.toString());
}

function countWords(paragraphs) {
  return paragraphs.reduce((count, paragraph) => count + paragraph.exactText.split(/\s+/).filter(Boolean).length, 0);
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


// The questions the cards were last built from. Rebuilding resets every card,
// so returning from Settings keeps results unless the questions changed.
let renderedGlanceSettings = null;

function renderGlanceMain() {
  const list = document.getElementById("glance-list");
  const container = document.getElementById("glance-container");
  const settings = loadGlanceSettings();
  const settingsKey = JSON.stringify(settings);
  if (settingsKey === renderedGlanceSettings) return;
  renderedGlanceSettings = settingsKey;
  // A running check asks the old questions; its answers no longer fit the cards.
  glanceController?.abort();
  glanceController = null;
  list.innerHTML = "";

  if (settings.length === 0) {
    if (container) container.style.display = "none";
    return;
  }

  if (container) container.style.display = "block";
  applyGlanceCollapsedState();

  const idle = "Press refresh to run this check.";
  settings.forEach(item => {
    // Built with textContent: titles are free text and must never be parsed as HTML.
    const div = document.createElement("div");
    div.className = "glance-item";
    div.id = `glance-item-${item.id}`;
    const header = document.createElement("div");
    header.className = "glance-header";
    const indicator = document.createElement("span");
    indicator.id = `glance-indicator-${item.id}`;
    indicator.className = "glance-indicator gray";
    const title = document.createElement("span");
    title.className = "glance-title";
    title.textContent = item.title;
    header.append(indicator, title);
    const summary = document.createElement("p");
    summary.id = `glance-summary-${item.id}`;
    summary.className = "glance-summary";
    summary.textContent = idle;
    div.append(header, summary);
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

    // Static markup only; the saved title and question are set as values below.
    card.innerHTML = `
      <div class="glance-card-header-row">
        <input type="text" class="ms-TextField-field glance-title-input" placeholder="Title">
        <span class="drag-handle" title="Drag to reorder">☰</span>
        <button class="delete-card-btn" title="Delete">✕</button>
      </div>
      <textarea class="ms-TextField-field glance-question-input" placeholder="Question (e.g. Is the grammar correct?)" rows="2"></textarea>
    `;
    card.querySelector(".glance-title-input").value = item.title ?? "";
    card.querySelector(".glance-question-input").value = item.question ?? "";

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
  const list = document.getElementById("glance-settings-list");
  const items = Array.from(list.children);
  const srcIndex = items.indexOf(dragSrcEl);
  const destIndex = items.indexOf(targetCard);

  // Only a card dragged by its handle reorders; dropped text or files never do
  // (indexOf -1 would otherwise move the last question).
  if (srcIndex >= 0 && destIndex >= 0 && srcIndex !== destIndex) {

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
  dragSrcEl = null;

  const items = document.querySelectorAll('.glance-settings-card');
  items.forEach(function (item) {
    item.classList.remove('over');
  });
}

// Only the latest Glance run may update the cards; starting a new run cancels the old one.
let glanceController = null;

async function runGlanceChecks() {
  const settings = loadGlanceSettings();
  if (settings.length === 0) return;
  const showAll = (status, text) => settings.forEach(item => {
    const indicator = document.getElementById(`glance-indicator-${item.id}`);
    const summary = document.getElementById(`glance-summary-${item.id}`);
    if (indicator) indicator.className = `glance-indicator ${status}`;
    if (summary) summary.textContent = text;
  });

  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    showAll("gray", "Add your Gemini API key in Settings to run checks.");
    return;
  }

  glanceController?.abort();
  const controller = new AbortController();
  glanceController = controller;
  const isCurrent = () => glanceController === controller && !controller.signal.aborted;

  showAll("gray", "Checking...");

  try {
    const { inspection } = await createWordDocumentSourceReader({ loadSupport: loadWordOperationSupport, getDocumentId, signal: controller.signal })();
    if (!isCurrent()) return;
    // Glance sends the full document text, so it must stay within the size limit.
    if (countWords(inspection.paragraphs) > DOCUMENT_LIMITS.MAX_WORDS) {
      showAll("gray", "This document is too long for Glance checks.");
      return;
    }
    const docText = inspection.paragraphs.map(paragraph => paragraph.exactText).join('\n');

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

    const result = await requestGemini({ model, apiKey: geminiApiKey, payload, signal: controller.signal, timeoutMs: TIMEOUT_LIMITS.FETCH_TIMEOUT_MS });
    if (!isCurrent()) return;
    const results = parseGlanceResults(glanceResponseText(result), settings.map(item => item.id));

    settings.forEach(item => {
      const indicator = document.getElementById(`glance-indicator-${item.id}`);
      const summary = document.getElementById(`glance-summary-${item.id}`);
      const res = results[item.id];
      if (!res) {
        if (summary) summary.textContent = "No answer for this check.";
        return;
      }
      if (indicator) {
        indicator.className = `glance-indicator ${res.status}`;
        indicator.classList.add("pulse");
        setTimeout(() => indicator.classList.remove("pulse"), 500);
      }
      if (summary) summary.textContent = res.summary;
    });
  } catch (error) {
    if (!isCurrent()) return;
    console.error("Glance check failed:", error);
    showAll("gray", "Error running check.");
  } finally {
    if (glanceController === controller) glanceController = null;
  }
}

// --- Checkpoint Management (IndexedDB-backed; see modules/storage/checkpoint-store.js) ---

/**
 * One-time migration of any legacy localStorage checkpoints into IndexedDB,
 * then removal of other sessions' stale checkpoints (full document copies).
 * Runs at most once per session; clears the legacy key afterward.
 */
async function ensureCheckpointMigration() {
  if (checkpointMigrationDone) return;
  checkpointMigrationDone = true; // set first so a failure doesn't retry every save
  try {
    const legacyJson = localStorage.getItem("docCheckpoints");
    if (legacyJson) {
      const legacy = JSON.parse(legacyJson);
      const records = migrateLegacyCheckpoints(legacy);
      if (records.length > 0) {
        await importCheckpoints(records);
        console.log(`Migrated ${records.length} checkpoint(s) from localStorage to IndexedDB.`);
      }
      localStorage.removeItem("docCheckpoints");
    }
  } catch (e) {
    console.warn("Checkpoint migration skipped:", e);
  }
  try {
    await pruneCheckpoints(documentInstanceId);
  } catch (e) {
    console.warn("Checkpoint pruning skipped:", e);
  }
}

/**
 * Save the exact body Flat OPC prepared for an automatic batch as a checkpoint
 * labeled `auto:<toolName>:<ISO>`.
 * @returns {Promise<number>} the new checkpoint id (>=1), or -1 on failure
 */
async function createAutoCheckpoint(toolName, sourceOoxml) {
  try {
    await ensureCheckpointMigration();
    return await saveCheckpoint(formatAutoCheckpointLabel(toolName), sourceOoxml, { sessionId: documentInstanceId });
  } catch (error) {
    // The direct batch executor refuses a write when this returns -1.
    console.error("Error saving checkpoint:", error);
    return -1;
  }
}

/**
 * Revert the body to the state before checkpoint `id`'s edit. Nothing is
 * written when the document changed after that edit (later edits or the user's
 * typing); the user is asked first. "Revert anyway" carries the state shown at
 * that moment, so anything changed after the question is asked about again.
 */
async function restoreCheckpoint(id, { confirmedToken } = {}) {
  // Never race an in-flight edit: its write could land after the restore.
  if (currentRequestController) {
    addMessageToChat('System', 'Wait for the current request to finish, or cancel it, before reverting.');
    return;
  }
  const restoreController = new AbortController();
  currentRequestController = restoreController;
  revertInProgress = true;
  const controls = ['chat-input', 'send-button', 'think-button'].map(id => document.getElementById(id)).filter(Boolean);
  controls.forEach(control => { control.disabled = true; });
  let msgElement;
  try {
    const record = await getCheckpoint(id);
    if (!record?.ooxml) {
      if (currentRequestController === restoreController) addMessageToChat('Error', 'This checkpoint is no longer available.');
      return;
    }
    if (restoreController.signal.aborted) return;
    msgElement = addMessageToChat('System', 'Reverting to checkpoint...');
    const { restoreBodyCheckpoint } = await import('./modules/chat/checkpoint-restoration.js');
    const result = await restoreBodyCheckpoint(record.ooxml, {
      getDocumentId, signal: restoreController.signal,
      guardToken: confirmedToken ?? (typeof record.afterToken === 'string' ? record.afterToken : null),
    });
    if (currentRequestController === restoreController) {
      if (result?.changedSinceEdit === true) {
        updateSystemMessage(msgElement, confirmedToken
          ? 'The document changed again after you chose to revert. Reverting will also undo those newer changes.'
          : result.message);
        addConfirmButtons(msgElement, {
          confirmLabel: 'Revert anyway',
          cancelLabel: 'Keep document',
          onConfirm: () => {
            // Refused now, so keep the question open to answer once the request ends.
            if (currentRequestController) {
              addMessageToChat('System', 'Wait for the current request to finish, or cancel it, before reverting.');
              return false;
            }
            restoreCheckpoint(id, { confirmedToken: result.currentToken });
          },
          onCancel: () => updateSystemMessage(msgElement, 'Revert cancelled. The document was not changed.'),
        });
      } else if (result?.success === true && result?.verified === true) {
        chatHistory = [];
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) chatMessages.innerHTML = '';
        addMessageToChat('System', result.message || 'Reverted and verified the body checkpoint.');
      } else {
        updateSystemMessage(msgElement, result?.message || 'The checkpoint restore could not be verified. Inspect the document before continuing.');
      }
    }
  } catch (error) {
    console.error('Error reverting checkpoint:', error);
    if (currentRequestController === restoreController) {
      if (msgElement) updateSystemMessage(msgElement, 'Error: Could not revert checkpoint.');
      else addMessageToChat('Error', 'Could not load the checkpoint.');
    }
  } finally {
    revertInProgress = false;
    if (currentRequestController === restoreController) {
      currentRequestController = null;
      controls.forEach(control => { control.disabled = false; });
    }
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

  // This request's own prompt, once added; error cleanup removes only this one.
  let promptTurn = null;
  // The last tool result shown in the bubble (an edit receipt or a refusal).
  let shownToolResult = null;
  // Settings changed while a request runs apply from the next request.
  const redlineAuthor = loadRedlineAuthor();
  const generateRedlines = loadRedlineSetting();
  // Refresh waits on this so it never clears the chat while Word may still be writing.
  let settleRequest;
  currentRequestSettled = new Promise(resolve => { settleRequest = resolve; });

  try {
    // --- Get Document Context ---
    const wordOperationSupportPromise = loadWordOperationSupport();
    wordOperationSupportPromise.catch(() => {});
    const readSource = createWordDocumentSourceReader({
      loadSupport: () => wordOperationSupportPromise,
      getDocumentId,
      signal: requestController.signal,
    });
    const documentSource = await readSource();
    const docText = summarizeDocumentSource(documentSource);
    const [, , inspectionModule, wordIdentity] = await wordOperationSupportPromise;
    const inspectionSession = inspectionModule.createDocumentInspectionSession({ readSource, maxTextChars: 4000,
      computeRevisionToken: wordIdentity.computeWordSourceRevisionToken,
      getExcludedParagraphIndex: wordIdentity.detectExportTailIndex,
      ...(globalThis.__WORD_SOURCE_DIAGNOSTICS__ ? { onStaleSource: globalThis.__WORD_SOURCE_DIAGNOSTICS__ } : {}),
    });
    const docSelection = documentSource.selectionText.slice(0, 2000);
    // --- Check Document Size ---
    if (currentRequestController !== requestController) return;
    const wordCount = countWords(documentSource.inspection.paragraphs);
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
    const apiUrl = geminiEndpoint(geminiModel);

    let contextString = "";
    if (docSelection && docSelection.trim() !== "") {
      contextString += `User Highlighted Text:\n"""${docSelection}"""\n\n`;
    }
    if (docText) {
      contextString += `Context from the current document:\n"""${docText}"""\n\n`;
    }
    const prompt = contextString
      ? `${contextString}User Question:\n${userMessage}`
      : userMessage;

    // Add to history
    promptTurn = { role: "user", parts: [{ text: prompt }] };
    chatHistory.push(promptTurn);

    // Maintain rolling window - but ensure we don't break function call/response pairs
    if (chatHistory.length > 10) {
      chatHistory = maintainHistoryWindow(chatHistory, 10);
    }

    const tools = getToolDeclarations();

    const systemInstruction = { parts: [{ text: `Owner drafting preferences:\n${loadSystemMessage()}\n\n${DOCUMENT_EDITING_GUIDE}` }] };

    // The portable loop owns exchanges and stop decisions; this adapter owns UI/Word I/O.
    const loopState = { history: chatHistory, records: requestRecords };
    const requestDeadline = requestStartTime + TIMEOUT_LIMITS.TOTAL_REQUEST_TIMEOUT_MS;
    // search_web runs its own grounded request with the same model, key, deadline and Cancel.
    const searchWeb = createWebSearch({
      request: payload => callGeminiWithRetry(apiUrl, payload, modelProfile.retries, 1000, requestController.signal, requestDeadline, geminiApiKey),
      safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
    });
    try {
      await runDocumentAgentLoop({
        state: loopState,
        registry: createToolRegistry(),
        loadRegistry: async () => {
          const toolsModule = await loadDocumentTools();
          return createToolRegistry({ searchWeb, ...toolsModule.createDocumentTools({
            session: inspectionSession, getDocumentId, getAuthor: () => redlineAuthor,
            getGenerateRedlines: () => generateRedlines,
            persistCheckpoint: source => createAutoCheckpoint('apply_document_edits', source),
            onVerifiedWrite: ({ checkpointIndex, revisionToken }) => Number.isInteger(checkpointIndex) && typeof revisionToken === 'string'
              ? updateCheckpoint(checkpointIndex, { afterToken: revisionToken })
              : undefined,
            signal: requestController.signal,
          }) });
        },
        requestModel: history => callGeminiWithRetry(apiUrl, {
          contents: history,
          systemInstruction,
          tools,
          safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
          generationConfig: { maxOutputTokens: modelProfile.maxOutputTokens },
        }, modelProfile.retries, 1000, requestController.signal, requestDeadline, geminiApiKey),
        refreshContext: () => captureRefreshedDocumentContext(createWordDocumentSourceReader({
          loadSupport: () => wordOperationSupportPromise, getDocumentId,
        })),
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
              shownToolResult = event.completionText;
              updateSystemMessage(loadingMsg, event.completionText,
                event.completionWithCheckpoint ? event.checkpointIndex : undefined);
            } else if (shownToolResult) {
              // A later inspection or search must not replace an edit's receipt.
              updateSystemMessage(loadingMsg, shownToolResult);
            }
            const messages = document.getElementById('chat-messages');
            if (loadingMsg && messages) messages.appendChild(loadingMsg);
          } else if (event.type === 'response') {
            if (!requestRecords.some(record => record.name === 'apply_document_edits')) removeMessage(loadingMsg);
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

    // Any failure after Cancel (including while reading the document) is a cancellation.
    if (requestController.signal.aborted || error.message === 'Request cancelled by user') {
      // Once a tool ran, the bubble holds its receipt and Revert button.
      if (requestRecords.length === 0) {
        removeMessage(loadingMsg);
        if (promptTurn && chatHistory.at(-1) === promptTurn) chatHistory.pop();
      }
      addMessageToChat("System", requestRecords.length === 0 ? "Request cancelled." : "Request cancelled. Inspect the document before retrying.");
    } else {
      // Once an edit was attempted the bubble keeps its receipt (the loop ends those failures itself).
      // Before that, even after read-only tools, the loop has restored the history, so drop this
      // request's prompt and a Retry resends the message cleanly.
      if (!wroteOrAttempted(requestRecords)) {
        removeMessage(loadingMsg);

        // Cleanup history for failed requests (timeout or error), never an earlier request's turn
        if (promptTurn && chatHistory.at(-1) === promptTurn) {
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
    try {
      removeCancelButton(loadingMsg);
      // A refreshed/superseded request cannot overwrite the next request's UI.
      if (currentRequestController === requestController) {
        currentRequestController = null;
        chatInput.disabled = false;
        sendButton.disabled = false;
        if (thinkButton) thinkButton.disabled = false;
        chatInput.focus();
      }
    } finally {
      // Always settle, or a waiting refresh would stay disabled for good.
      settleRequest(summarizeRequestOutcome(requestRecords));
    }
  }
}

// Retries provider transport only. Tool execution remains in the caller.
async function callGeminiWithRetry(url, payload, retries = 3, backoff = 1000, signal = currentRequestController?.signal, deadline = Infinity, apiKey = loadApiKey()) {
  return requestGeminiJson(url, payload, {
    apiKey,
    signal,
    timeoutMs: TIMEOUT_LIMITS.FETCH_TIMEOUT_MS,
    maxAttempts: retries,
    backoffMs: backoff,
    deadline,
  });
}
