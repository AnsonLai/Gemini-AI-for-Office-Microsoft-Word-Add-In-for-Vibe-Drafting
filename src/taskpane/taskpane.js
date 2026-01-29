/*
 * Gemini AI for Office - Task Pane Implementation
 * Author: Anson Lai
 * Location: Vancouver, Canada
 * Description: Word add-in integrating Google Gemini AI for document editing and analysis
 */

/* global document, Office, Word, localStorage */

import { marked } from 'marked';
import { diff_match_patch } from 'diff-match-patch';
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
  validateHistoryPairs,
  sanitizeHistory,
  removeAllFunctionPairs,
  createFreshStartWithContext
} from './modules/chat/chat-history.js';
import {
  initAgenticTools,
  executeRedline,
  executeComment,
  executeHighlight,
  executeNavigate,
  executeResearch,
  executeInsertListItem,
  executeEditList,
  executeConvertHeadersToList,
  executeEditTable,
  executeEditSection
} from './modules/commands/agentic-tools.js';

// Configure marked for GFM (GitHub Flavored Markdown) with tables, breaks, etc.
marked.setOptions({
  gfm: true,           // Enable GitHub Flavored Markdown
  breaks: true,        // Convert \n to <br>
});

// ==================== CONFIGURATION CONSTANTS ====================

const DEFAULT_AUTHOR = "Gemini AI";

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
  TOKEN_MULTIPLIER: 1.33     // Words to tokens conversion factor
};

// Storage quotas
const STORAGE_LIMITS = {
  SAFE_LIMIT: 4500000,       // ~4.5MB safe limit for localStorage
  MIN_PRUNE_COUNT: 5         // Minimum checkpoints to prune when quota exceeded
};

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
      enhancedLine: enhancedLine
    });
  }

  return {
    paragraphs: enhancedParagraphs,
    formattedText: enhancedParagraphs.map(p => p.enhancedLine).join("\n"),
    sectionCount: sectionCounter
  };
}

let chatHistory = [];
let toolsExecutedInCurrentRequest = [];  // Track successful tool executions for recovery

Office.onReady((info) => {
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

initAgenticTools({
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
});

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
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiApiKey}`;

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

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    const result = await response.json();
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

// --- Checkpoint Management ---

function getCheckpoints() {
  const checkpointsJson = localStorage.getItem("docCheckpoints");
  return checkpointsJson ? JSON.parse(checkpointsJson) : [];
}

function saveCheckpoints(checkpoints) {
  const MAX_RETRIES = 10; // Maximum number of retry attempts

  let retries = 0;
  while (retries < MAX_RETRIES) {
    try {
      localStorage.setItem("docCheckpoints", JSON.stringify(checkpoints));
      return true; // Success
    } catch (error) {
      if (error.name === 'QuotaExceededError' && checkpoints.length > 1) {
        // Remove 50% of checkpoints (more aggressive pruning)
        const toRemove = Math.max(1, Math.floor(checkpoints.length / 2));
        checkpoints.splice(0, toRemove);
        console.warn(`QuotaExceededError: Removed ${toRemove} oldest checkpoint(s), ${checkpoints.length} remaining. Retrying...`);
        retries++;
      } else if (error.name === 'QuotaExceededError' && checkpoints.length <= 1) {
        // Can't prune anymore, clear all and give up gracefully
        console.warn("Storage quota exceeded. Clearing all checkpoints.");
        try {
          localStorage.removeItem("docCheckpoints");
        } catch (e) { /* ignore */ }
        return false; // Silently fail rather than throw
      } else {
        // Not a quota error
        console.error("Failed to save checkpoints:", error);
        return false; // Silently fail rather than throw
      }
    }
  }

  // If we've exhausted retries, fail gracefully
  console.warn("Unable to save checkpoint after max retries. Clearing checkpoints.");
  try {
    localStorage.removeItem("docCheckpoints");
  } catch (e) { /* ignore */ }
  return false;
}

// function updateCheckpointStatus() { ... } removed as UI is gone.

async function createCheckpoint(silent = false) {
  if (!silent) {
    addMessageToChat("System", "Saving checkpoint...");
  }
  try {
    return await Word.run(async (context) => {
      const ooxml = context.document.body.getOoxml();
      await context.sync();

      // 'ooxml.value' is a base64 string of the entire document body
      const ooxmlLength = ooxml.value.length;
      console.log(`Checkpoint OOXML length: ${ooxmlLength}`);

      const checkpoints = getCheckpoints();

      // Check for quota issues roughly (5MB limit usually)
      let totalSize = 0;
      checkpoints.forEach(c => totalSize += c.length);
      console.log(`Current total checkpoints size: ${totalSize}`);

      let prunedCount = 0;

      // Prune at least MIN_PRUNE_COUNT checkpoints if we need to prune any, to create a buffer
      while ((totalSize + ooxmlLength > STORAGE_LIMITS.SAFE_LIMIT || (prunedCount > 0 && prunedCount < STORAGE_LIMITS.MIN_PRUNE_COUNT)) && checkpoints.length > 0) {
        const removed = checkpoints.shift(); // Remove oldest
        totalSize -= removed.length;
        prunedCount++;
      }

      if (prunedCount > 0) {
        console.warn(`LocalStorage quota exceeded. Removed ${prunedCount} oldest checkpoint(s).`);
        if (!silent) {
          addMessageToChat("System", `Storage full. Removed ${prunedCount} old checkpoint(s) to make space.`);
        }
      }

      checkpoints.push(ooxml.value);
      saveCheckpoints(checkpoints);

      if (!silent) {
        addMessageToChat("System", `Checkpoint saved. Total: ${checkpoints.length}`);
      }

      // Return the index of the newly created checkpoint (0-based)
      return checkpoints.length - 1;
    });
  } catch (error) {
    console.error("Error saving checkpoint:", error);
    if (!silent) {
      addMessageToChat("Error", `Could not save checkpoint. ${error.message}`);
    }
    return -1;
  }
}


async function restoreCheckpoint(index) {
  const checkpoints = getCheckpoints();
  if (index < 0 || index >= checkpoints.length) {
    addMessageToChat("Error", "Invalid checkpoint index.");
    return;
  }

  const msgElement = addMessageToChat("System", `Reverting to checkpoint #${index + 1}...`);

  const targetCheckpointOoxml = checkpoints[index];

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
      context.document.body.insertOoxml(targetCheckpointOoxml, "Replace");
      await context.sync();

      // Optionally restore track changes, but reverting usually implies going back to a state.
      // If we restore it, we might want to do it cleanly.
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
  toolsExecutedInCurrentRequest = [];

  // Sanitize history to remove any hanging function calls from interrupted sessions
  chatHistory = sanitizeHistory(chatHistory);

  // Set up abort controller for this request (allows user cancellation)
  currentRequestController = new AbortController();
  const requestStartTime = Date.now();

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
    let docText = "";
    let docComments = [];
    let docRedlines = [];
    let docSelection = "";

    try {
      await Word.run(async (context) => {
        const body = context.document.body;

        // Fetch current selection
        const selection = context.document.getSelection();
        selection.load("text");

        // Fetch comments
        const comments = context.document.comments;
        comments.load("content, authorName, creationDate");

        // Fetch tracked changes (redlines)
        let trackedChanges = null;
        try {
          trackedChanges = body.getTrackedChanges();
          trackedChanges.load("type, text, author, date");
        } catch (e) {
          console.warn("Tracked changes API not supported or failed:", e);
        }

        await context.sync();

        // Use enhanced document context extraction for rich formatting info
        try {
          const enhancedContext = await extractEnhancedDocumentContext(context);
          docText = enhancedContext.formattedText;
          console.log(`Enhanced context extracted: ${enhancedContext.paragraphs.length} paragraphs, ${enhancedContext.sectionCount} sections`);
        } catch (enhancedError) {
          console.warn("Enhanced context extraction failed, falling back to simple extraction:", enhancedError);
          // Fallback to simple extraction
          const paragraphs = body.paragraphs;
          paragraphs.load("text");
          await context.sync();
          docText = paragraphs.items.map((p, index) => `[P${index + 1}] ${p.text}`).join("\n");
        }

        docSelection = selection.text;

        if (comments && comments.items) {
          docComments = comments.items.map((c) => {
            return `[Comment by ${c.authorName} on ${c.creationDate}]: ${c.content}`;
          });
        }

        if (trackedChanges && trackedChanges.items) {
          docRedlines = trackedChanges.items.map((tc) => {
            const type = tc.type; // "Inserted" or "Deleted"
            return `[${type} by ${tc.author} on ${tc.date}]: "${tc.text}"`;
          });
        }
      });
    } catch (error) {
      console.warn("Could not fetch document text, comments, or redlines:", error);
    }

    // --- Check Document Size ---
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
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

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

    // Define tools
    const tools = [
      {
        function_declarations: [
          {
            name: "apply_redlines",
            description: "Applies suggested edits to the document. Use this tool whenever the user asks to 'edit text', 'change text', 'modify', 'add', 'delete', 'reword', 'rephrase', 'update', 'bold', 'italicize', 'underline', 'strikethrough', or apply any TEXT FORMATTING to the document.\n\nIMPORTANT - FORMATTING RULES:\n- Bold: **text**\n- Italic: *text*\n- Underline: ++text++\n- Strikethrough: ~~text~~\n\nIMPORTANT - LIST RULES:\n- Use Markdown syntax for lists. \n- For Bullet Lists: Use '* item'. For nested items, indent with 4 spaces (e.g., '    * sub-item').\n- For Numbered Lists: Use '1. item', 'a. item', 'i. item', etc. explicitly. \n- For Nested Numbering: Use '1.1.', '1.1.1.' styles if appropriate. \n- DO NOT use simple hyphens ('-') if you intend to create a structured or numbered list. \n- INDENTATION is critical for sub-levels. Use 2 or 4 spaces.\n\nDo NOT suggest changes in the chat; always use this tool to apply them directly. The edits will be applied under track changes (redlines). NEVER say you have applied edits unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The specific instruction for how to edit the document (e.g., 'Change Lessor to Landlord', 'Fix spelling', 'Reword the introduction').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "insert_comment",
            description: "Inserts comments into the document based on the user's instruction. Use this tool to flag risks, add notes, or review specific sections. NEVER say you have inserted comments unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to comment on and what to say (e.g., 'Flag all risky clauses', 'Comment on the first paragraph').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "highlight_text",
            description: "Highlights text with a colored background marker (like a highlighter pen). ONLY use this tool when the user EXPLICITLY asks to 'highlight' text. Do NOT use this for formatting requests like 'bold', 'italicize', or general emphasis - those should use apply_redlines with markdown syntax instead. Use this tool ONLY for explicit highlight requests like 'highlight all dates in yellow' or 'mark these terms with highlighting'. NEVER say you have highlighted text unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to highlight (e.g., 'Highlight all dates', 'Mark placeholders').",
                },
                color: {
                  type: "STRING",
                  enum: ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "gray25", "gray50", "black", "white"],
                  description: "Optional: highlight color. Default is 'yellow'. Options include: yellow, green, cyan, magenta, blue, red, and dark variants.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "perform_research",
            description: "Performs a Google Search to answer questions that require external knowledge, facts, or up-to-date information. Use this when the user asks for information not in the document.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The search query to send to Google.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "navigate_to_section",
            description: "Navigates to and selects a specific section of the document. Use this when the user asks to go to, scroll to, find, or jump to a particular part of the document (e.g., 'go to the introduction', 'scroll to paragraph 5', 'find the signature block', 'show me the definitions section'). This helps users quickly locate relevant content without manually scrolling.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The navigation instruction describing what section to go to (e.g., 'go to paragraph 3', 'find the table of contents', 'scroll to the conclusion', 'show me where parties are defined').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "edit_list",
            description: "Edit an entire list as a unit. Use this when you need to modify, add, or remove items from a bulleted or numbered list. This preserves list formatting and structure better than apply_redlines. Look for paragraphs with |ListNumber or |ListBullet in the context. For numbered lists, you can specify different numbering styles: '1, 2, 3' (decimal - default), 'a, b, c' (lowerAlpha), 'A, B, C' (upperAlpha), 'i, ii, iii' (lowerRoman), or 'I, II, III' (upperRoman). NEVER say you have edited a list unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                startParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the FIRST item in the list (e.g., 3 for [P3])",
                },
                endParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the LAST item in the list",
                },
                newItems: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "The new list items in order. Each string is one list item text (without bullets/numbers).",
                },
                listType: {
                  type: "STRING",
                  enum: ["bullet", "numbered"],
                  description: "The type of list to create",
                },
                numberingStyle: {
                  type: "STRING",
                  enum: ["decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"],
                  description: "Optional: For numbered lists, the numbering style to use. Default is 'decimal' (1, 2, 3). Options: 'decimal' (1, 2, 3), 'lowerAlpha' (a, b, c), 'upperAlpha' (A, B, C), 'lowerRoman' (i, ii, iii), 'upperRoman' (I, II, III).",
                },
              },
              required: ["startParagraphIndex", "endParagraphIndex", "newItems", "listType"],
            },
          },
          {
            name: "insert_list_item",
            description: "Insert a single list item after a specific paragraph. Use this for surgical additions to an existing list - it inherits the numbering format from the paragraph you insert after. Much better than edit_list when you only need to add one or two items. Do NOT include numbering markers in the text - Word will add them automatically.",
            parameters: {
              type: "OBJECT",
              properties: {
                afterParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index to insert after (e.g., 5 to insert after [P5])",
                },
                text: {
                  type: "STRING",
                  description: "The text content of the new list item (WITHOUT any numbering like '1.' or '1.1.' - Word adds these automatically)",
                },
                indentLevel: {
                  type: "INTEGER",
                  description: "Optional: The indentation level relative to the paragraph you're inserting after. 0 = same level (default), 1 = one level deeper (sub-item), -1 = one level shallower",
                },
              },
              required: ["afterParagraphIndex", "text"],
            },
          },
          {
            name: "edit_table",
            description: "Edit a table as a unit. Use this when you need to modify table content, add/remove rows or columns. This preserves table formatting. Look for paragraphs with |T:row,col in the context. NEVER say you have edited a table unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndex: {
                  type: "INTEGER",
                  description: "Any paragraph index that is part of the table (has T:row,col marker)",
                },
                action: {
                  type: "STRING",
                  enum: ["replace_content", "add_row", "delete_row", "update_cell"],
                  description: "The table operation to perform",
                },
                content: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "For replace_content: 2D array of strings [[row1cells], [row2cells]]. For add_row: array of cell values. For update_cell: single-element array with new text.",
                },
                targetRow: {
                  type: "INTEGER",
                  description: "For add_row/delete_row/update_cell: the 0-based row index",
                },
                targetColumn: {
                  type: "INTEGER",
                  description: "For update_cell: the 0-based column index",
                },
              },
              required: ["paragraphIndex", "action"],
            },
          },
          {
            name: "edit_section",
            description: "Edit a document section as a unit. Use this for legal contracts where numbered/lettered items serve as section headers (marked with §) followed by body text (marked with §N). This preserves the section structure and list numbering. NEVER say you have edited a section unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                sectionHeaderIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the section header (the list item marked with §, e.g., '1. Definitions')",
                },
                newHeaderText: {
                  type: "STRING",
                  description: "Optional: new text for the section header. The list number/letter is automatically preserved.",
                },
                newBodyParagraphs: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new body paragraphs for this section. Each string becomes one paragraph. Omit to keep existing body.",
                },
                preserveSubsections: {
                  type: "BOOLEAN",
                  description: "If true, only edits body text until the next subsection. If false or omitted, replaces entire section including subsections.",
                },
              },
              required: ["sectionHeaderIndex"],
            },
          },
          {
            name: "convert_headers_to_list",
            description: "Convert non-contiguous headers to a numbered list. Use this when headers like '1. PURPOSE', '2. DEFINITION', '3. EXCLUSIONS' have body text between them and need to be converted to a proper auto-numbered list. The tool strips manual numbering and creates a Word list where all headers share continuous numbering. Supports different formats: 1,2,3 or a,b,c or i,ii,iii. NEVER say you have converted headers unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndices: {
                  type: "ARRAY",
                  items: { type: "INTEGER" },
                  description: "Array of 1-based paragraph indices of the headers to convert (e.g., [3, 7, 15] for headers at P3, P7, P15)",
                },
                newHeaderTexts: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new text for each header (without numbers). If omitted, existing text is used with manual numbers stripped.",
                },
                numberingFormat: {
                  type: "STRING",
                  enum: ["arabic", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"],
                  description: "Optional: numbering format. 'arabic' = 1,2,3 (default), 'lowerLetter' = a,b,c, 'upperLetter' = A,B,C, 'lowerRoman' = i,ii,iii, 'upperRoman' = I,II,III",
                },
              },
              required: ["paragraphIndices"],
            },
          },
        ],
      },
    ];

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
- The § marker indicates section structure - paragraphs marked §N belong to section N

IMPORTANT: You have access to tools. You can chat and respond normally to questions. However, when the user asks for an action that involves manipulating the document, you should HEAVILY FAVOR using the corresponding tool rather than just describing the action.

CRITICAL: If the user asks to 'edit text' or make any changes, you MUST use the \`apply_redlines\` tool.
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
 - Do NOT use generic bullets ('-') if you want specific numbering. The engine relies on your markers (e.g., '1.1.') to detect the list type.`,
        },
      ],
    };

    // --- Tool Execution Loop with Multi-Tier Recovery ---
    let loopCount = 0;
    let keepLooping = true;
    let currentRecoveryTier = 0;  // 0=normal, 1=validate pairs, 2=remove all pairs, 3=fresh start, 4=graceful degrade
    const originalUserMessage = prompt;  // Save for Tier 3 recovery

    while (keepLooping && loopCount < DOCUMENT_LIMITS.MAX_LOOPS) {
      loopCount++;
      console.log(`Starting chat loop iteration ${loopCount} (recovery tier: ${currentRecoveryTier})`);

      // Check for user cancellation
      if (currentRequestController && currentRequestController.signal.aborted) {
        console.log('Request cancelled by user during loop');
        removeMessage(loadingMsg);
        addMessageToChat("System", "Request cancelled.");
        keepLooping = false;
        break;
      }

      // Check for overall timeout
      const elapsedTime = Date.now() - requestStartTime;
      if (elapsedTime > TIMEOUT_LIMITS.TOTAL_REQUEST_TIMEOUT_MS) {
        console.warn(`Overall request timeout exceeded: ${elapsedTime}ms`);
        removeMessage(loadingMsg);

        // If some tools executed successfully, show partial success
        if (toolsExecutedInCurrentRequest.length > 0) {
          const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
          const throttleWarning = "\n\nIf you're using Gemini 3, it is in preview and your access has likely been throttled. Please go into settings and revert to Gemini 2.5.";

          if (successMessage) {
            addMessageToChat("System", successMessage + "\n\n*(Request timed out after completing some changes)*" + throttleWarning);
          } else {
            addMessageToChat("Error", "Request timed out. Some changes may have been applied." + throttleWarning);
          }
        } else {
          // Specific message for throttle/timeout
          addMessageToChat("Error", "If you're using Gemini 3, it is in preview and your access has likely been throttled. Please go into settings and revert to Gemini 2.5.");

          // Discard the timed out request from history to allow user to continue clean
          // Remove the last user message we added for this request
          // (The one pushed at `chatHistory.push({ role: "user", parts: [{ text: prompt }] });`)
          // We only remove it if we haven't successfully done tools that we want to keep context for.
          if (chatHistory.length > 0 && chatHistory[chatHistory.length - 1].role === "user") {
            console.log("Discarding timed out request from history");
            chatHistory.pop();
          }
        }
        keepLooping = false;
        break;
      }

      // Prepare payload with current history
      const payload = {
        contents: chatHistory,
        systemInstruction: systemInstruction,
        tools: tools,
        safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
        generationConfig: {
          maxOutputTokens: API_LIMITS.MAX_OUTPUT_TOKENS
        },
      };

      console.log("Sending Chat History to API:", JSON.stringify(chatHistory, null, 2));

      let result;
      try {
        result = await callGeminiWithRetry(apiUrl, payload);
      } catch (apiError) {
        console.error(`API Error on iteration ${loopCount}:`, apiError);

        // Check if this is a function call/response mismatch error
        const isFunctionCallError = apiError.message && (
          apiError.message.includes("function response turn comes immediately after a function call turn") ||
          apiError.message.includes("function call turn comes immediately after a user turn or after a function response turn")
        );

        if (isFunctionCallError) {
          currentRecoveryTier++;
          console.warn(`Function call error detected. Escalating to recovery tier ${currentRecoveryTier}`);

          if (currentRecoveryTier === 1) {
            // Tier 1: Validate and clean history pairs
            console.log("Tier 1: Validating history pairs...");
            const originalLength = chatHistory.length;
            chatHistory = validateHistoryPairs(chatHistory);
            console.log(`History cleaned: ${originalLength} -> ${chatHistory.length} messages`);
            loopCount = 0;  // Reset to retry
            continue;
          } else if (currentRecoveryTier === 2) {
            // Tier 2: Remove ALL function pairs
            console.log("Tier 2: Removing all function call/response pairs...");
            chatHistory = removeAllFunctionPairs(chatHistory);
            console.log(`History after removing function pairs: ${chatHistory.length} messages`);
            loopCount = 0;
            continue;
          } else if (currentRecoveryTier === 3) {
            // Tier 3: Fresh start with original context
            console.log("Tier 3: Creating fresh start with original context...");
            chatHistory = createFreshStartWithContext(originalUserMessage);
            console.log(`History reset to fresh start: ${chatHistory.length} messages`);
            loopCount = 0;
            continue;
          } else {
            // Tier 4: Graceful degradation
            console.log("Tier 4: All recovery attempts failed. Checking for graceful degradation...");
            removeMessage(loadingMsg);

            const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
            if (successMessage) {
              addMessageToChat("System", successMessage + "\n\n*(Conversation refreshed)*");
              // Reset history for next request
              chatHistory = [];
            } else {
              addMessageToChat("Error", "I encountered an issue with the conversation. Please try again.");
            }
            keepLooping = false;
            break;
          }
        }

        // Non-recoverable errors after successful tool execution
        if (loopCount > 1 && toolsExecutedInCurrentRequest.length > 0) {
          console.warn("Stopping loop due to API error after successful tool execution.");
          const successMessage = generateSuccessMessage(toolsExecutedInCurrentRequest);
          if (successMessage) {
            if (loadingMsg) {
              updateSystemMessage(loadingMsg, successMessage + "\n\n*(Conversation refreshed)*");
            } else {
              addMessageToChat("System", successMessage + "\n\n*(Conversation refreshed)*");
            }
            chatHistory = [];
          }
          keepLooping = false;
          break;
        } else {
          throw apiError;
        }
      }

      console.log("Gemini chat raw result:", JSON.stringify(result, null, 2));

      if (!result.candidates || !Array.isArray(result.candidates) || result.candidates.length === 0) {
        throw new Error("Gemini response contained no candidates.");
      }

      const candidate = result.candidates[0];
      let parts = [];
      let content = candidate.content;

      if (content && content.parts && Array.isArray(content.parts)) {
        parts = content.parts;
      } else if (candidate.finishReason === "MALFORMED_FUNCTION_CALL" && candidate.finishMessage) {
        console.warn("Gemini returned MALFORMED_FUNCTION_CALL. Attempting to recover...", candidate.finishMessage);
        const redlineMatch = candidate.finishMessage.match(/apply_redlines\s*\{\s*instruction\s*:\s*(.*)\s*\}/s);
        if (redlineMatch && redlineMatch[1]) {
          const instruction = redlineMatch[1].trim();
          console.log("Recovered instruction:", instruction);
          parts = [{
            functionCall: {
              name: "apply_redlines",
              args: { instruction: instruction }
            }
          }];
          // Ensure content has the proper structure with role
          if (!content || !content.role) {
            content = { role: "model", parts: parts };
          } else {
            content.parts = parts;
          }
        }
      }

      if (parts.length === 0) {
        console.error("Gemini candidate missing content.parts:", candidate);
        throw new Error("Gemini response was missing content.parts (possibly blocked by safety settings or malformed).");
      }

      console.log("Gemini chat content.parts:", parts);

      // Check for ALL function calls in the response
      const functionCallParts = parts.filter((part) => part.functionCall);

      if (functionCallParts.length > 0) {
        // If this is the first loop, remove the "Thinking..." message so we can show tool status
        // Keep loading message visible during tool execution


        // Execute ALL function calls and collect responses
        const functionResponses = [];

        for (const functionCallPart of functionCallParts) {
          const functionCall = functionCallPart.functionCall;
          const args = functionCall.args;
          const instruction = args.instruction;

          // Update loading message status
          if (loadingMsg) {
            const toolFriendlyNames = {
              "apply_redlines": `Applying edits: "${instruction}"...`,
              "insert_comment": `Inserting comments: "${instruction}"...`,
              "highlight_text": `Highlighting text: "${instruction}"...`,
              "perform_research": `Researching: "${instruction}"...`,
              "navigate_to_section": `Navigating to: "${instruction}"...`
            };
            const statusText = toolFriendlyNames[functionCall.name] || "Working...";
            updateSystemMessage(loadingMsg, statusText);
          }


          let toolResult = "";

          if (functionCall.name === "apply_redlines") {
            const checkpointIndex = await createCheckpoint(true);
            const result = await executeRedline(instruction, docText);
            toolResult = result.message;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            // Only show to user if there were actual changes or a true error
            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 edits): ${toolResult}`);
            }

          } else if (functionCall.name === "insert_comment") {
            const checkpointIndex = await createCheckpoint(true);
            const result = await executeComment(instruction, docText);
            toolResult = result.message;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 comments): ${toolResult}`);
            }

          } else if (functionCall.name === "highlight_text") {
            const checkpointIndex = await createCheckpoint(true);
            const highlightColor = args.color || "yellow";
            const result = await executeHighlight(instruction, docText, highlightColor);
            toolResult = result.message;

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: result.showToUser
            });

            if (result.showToUser) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              console.log(`Fallback in progress (0 highlights): ${toolResult}`);
            }

          } else if (functionCall.name === "perform_research") {
            updateSystemMessage(loadingMsg, `Researching: "${instruction}"...`);
            toolResult = await executeResearch(instruction);

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: true
            });

            updateSystemMessage(loadingMsg, `Found search results for: "${instruction}"`);
          } else if (functionCall.name === "navigate_to_section") {
            updateSystemMessage(loadingMsg, `Navigating to: "${instruction}"...`);
            toolResult = await executeNavigate(instruction, docText);

            // Track successful tool execution for recovery
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: instruction,
              result: toolResult,
              success: true
            });

            updateSystemMessage(loadingMsg, `Navigated to: "${instruction}"`);
          } else if (functionCall.name === "edit_list") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing list from P${args.startParagraphIndex} to P${args.endParagraphIndex}...`);

            const result = await executeEditList(
              args.startParagraphIndex,
              args.endParagraphIndex,
              args.newItems,
              args.listType,
              args.numberingStyle
            );
            toolResult = result.message;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_list P${args.startParagraphIndex}-P${args.endParagraphIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "insert_list_item") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Inserting list item after P${args.afterParagraphIndex}...`);

            const result = await executeInsertListItem(
              args.afterParagraphIndex,
              args.text,
              args.indentLevel || 0
            );
            toolResult = result.message;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `insert_list_item after P${args.afterParagraphIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "edit_table") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing table (${args.action})...`);

            const result = await executeEditTable(
              args.paragraphIndex,
              args.action,
              args.content,
              args.targetRow,
              args.targetColumn
            );
            toolResult = result.message;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_table at P${args.paragraphIndex}: ${args.action}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "edit_section") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Editing section at P${args.sectionHeaderIndex}...`);

            const result = await executeEditSection(
              args.sectionHeaderIndex,
              args.newHeaderText,
              args.newBodyParagraphs,
              args.preserveSubsections
            );
            toolResult = result.message;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `edit_section at P${args.sectionHeaderIndex}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          } else if (functionCall.name === "convert_headers_to_list") {
            const checkpointIndex = await createCheckpoint(true);
            updateSystemMessage(loadingMsg, `Converting ${args.paragraphIndices?.length || 0} headers to numbered list...`);

            const result = await executeConvertHeadersToList(
              args.paragraphIndices,
              args.newHeaderTexts,
              args.numberingFormat
            );
            toolResult = result.message;

            // Track successful tool execution
            toolsExecutedInCurrentRequest.push({
              name: functionCall.name,
              instruction: `convert_headers_to_list: ${args.paragraphIndices?.join(', ')}`,
              result: toolResult,
              success: result.success
            });

            if (result.success) {
              updateSystemMessage(loadingMsg, toolResult, checkpointIndex);
            } else {
              updateSystemMessage(loadingMsg, toolResult);
            }
          }

          // Move loading message to bottom after tool output
          if (loadingMsg) {
            const chatMessages = document.getElementById("chat-messages");
            if (chatMessages) chatMessages.appendChild(loadingMsg);
          }

          // Collect this function response

          // Shape this exactly as Gemini expects:
          // functionResponse: {
          //   name: "tool_name",
          //   response: {
          //     name: "tool_name",
          //     content: [ { text: "..." } ]
          //   }
          // }
          functionResponses.push({
            functionResponse: {
              name: functionCall.name,
              response: {
                name: functionCall.name,
                content: [
                  {
                    text: toolResult || ""
                  }
                ]
              }
            }
          });
        }

        // NOW add both the model's function call and the responses to history together
        // This ensures they're added as a complete pair
        chatHistory.push({
          role: "model",
          parts: parts
        });

        chatHistory.push({
          role: "user",
          parts: functionResponses
        });

      } else {
        // Normal text response - this ends the loop
        const aiResponse = parts[0].text;

        // Add model response to history with proper structure
        chatHistory.push({
          role: "model",
          parts: parts
        });

        if (toolsExecutedInCurrentRequest.length === 0) {
          removeMessage(loadingMsg);
        }
        addMessageToChat("Gemini", aiResponse);
        keepLooping = false;
      }
    }

    // Maintain rolling window - but ensure we don't break function call/response pairs
    if (chatHistory.length > 10) {
      chatHistory = maintainHistoryWindow(chatHistory, 10);
    }

  } catch (error) {
    console.error("Error calling Gemini API:", error);

    // Handle user cancellation specifically
    if (error.message === 'Request cancelled by user') {
      removeMessage(loadingMsg);
      addMessageToChat("System", "Request cancelled.");
    } else {
      // Only remove loadingMsg if no tools were executed (meaning it's still a "Thinking" message)
      if (toolsExecutedInCurrentRequest.length === 0) {
        removeMessage(loadingMsg);

        // Cleanup history for failed requests (timeout or error)
        if (chatHistory.length > 0 && chatHistory[chatHistory.length - 1].role === "user") {
          console.log("Discarding failed request from history");
          chatHistory.pop();
        }
      }

      let errorMessage = error.message ? `Sorry, I couldn't get a response. Error: ${error.message}` : `Sorry, I couldn't get a response. Error: ${String(error)}`;

      // Override error message for timeouts
      if (error.message && (error.message.includes("timed out") || error.message.includes("timeout"))) {
        errorMessage = "Gemini 3 is in preview and they have likely been throttled. Please go into settings and revert to Gemini 2.5.";
      }

      const errorMsgEl = addMessageToChat("Error", errorMessage);

      // Add retry button if it's the specific missing content error
      if (error.message && error.message.includes("Gemini response was missing content.parts")) {
        addRetryButton(errorMsgEl, userMessage);
      }
    }
  } finally {
    // Clear the global abort controller
    currentRequestController = null;

    // Unlock UI
    chatInput.disabled = false;
    sendButton.disabled = false;
    if (thinkButton) thinkButton.disabled = false;
    chatInput.focus();
  }
}

// Helper with retry logic and timeout support
async function callGeminiWithRetry(url, payload, retries = 3, backoff = 1000) {
  for (let i = 0; i < retries; i++) {
    // Create abort controller for this specific fetch attempt
    const fetchController = new AbortController();

    // Create timeout that will abort the fetch
    const timeoutId = setTimeout(() => {
      fetchController.abort();
    }, TIMEOUT_LIMITS.FETCH_TIMEOUT_MS);

    try {
      // Also check if the global request controller was aborted (user cancelled)
      if (currentRequestController && currentRequestController.signal.aborted) {
        throw new Error('Request cancelled by user');
      }

      // Listen for global cancellation
      const onGlobalAbort = () => fetchController.abort();
      if (currentRequestController) {
        currentRequestController.signal.addEventListener('abort', onGlobalAbort);
      }

      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal: fetchController.signal
      });

      // Clean up listeners
      clearTimeout(timeoutId);
      if (currentRequestController) {
        currentRequestController.signal.removeEventListener('abort', onGlobalAbort);
      }

      if (!response.ok) {
        const text = await response.text();

        // Check for the specific function call/response error (400 error)
        const isFunctionCallError = response.status === 400 &&
          text.includes("function response turn comes immediately after a function call turn");

        if (isFunctionCallError) {
          // Don't retry this error here - let the caller handle it
          throw new Error(`API failed: ${text}`);
        }

        // Only retry on 5xx errors
        if (response.status >= 500 && response.status < 600) {
          console.warn(`Attempt ${i + 1} failed with ${response.status}: ${text}`);
          if (i === retries - 1) throw new Error(`API failed after ${retries} attempts: ${text}`);
          // Wait before retrying
          await new Promise(r => setTimeout(r, backoff * Math.pow(2, i))); // Exponential backoff
          continue;
        }

        throw new Error(`API failed: ${text}`);
      }

      return await response.json();
    } catch (error) {
      clearTimeout(timeoutId);

      // Check if this was a user cancellation
      if (error.name === 'AbortError' || error.message === 'Request cancelled by user') {
        if (currentRequestController && currentRequestController.signal.aborted) {
          throw new Error('Request cancelled by user');
        }
        // This was a timeout abort
        console.warn(`Attempt ${i + 1} timed out after ${TIMEOUT_LIMITS.FETCH_TIMEOUT_MS / 1000}s`);
        if (i === retries - 1) {
          throw new Error(`Request timed out. The AI is taking longer than usual. Please try again.`);
        }
        await new Promise(r => setTimeout(r, backoff * Math.pow(2, i)));
        continue;
      }

      // If it's the function call error, throw immediately without retry
      if (error.message && error.message.includes("function response turn comes immediately after a function call turn")) {
        throw error;
      }

      if (i === retries - 1) throw error;
      console.warn(`Attempt ${i + 1} failed: ${error.message}`);
      await new Promise(r => setTimeout(r, backoff * Math.pow(2, i)));
    }
  }
}

