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

// OOXML Engine V5.1 - Hybrid Mode for surgical document editing with track changes
import { applyRedlineToOxml, ReconciliationPipeline, wrapInDocumentFragment, preprocessMarkdown, parseTable } from './modules/reconciliation/index.js';

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

// Cached document font - set by detectDocumentFont() before edits
let cachedDocumentFont = "Calibri"; // Safe default for Word

/**
 * Detects and caches the document's font from the first paragraph.
 * Should be called before making edits to ensure font consistency.
 */
async function detectDocumentFont() {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (paragraphs.items.length > 0) {
        const firstPara = paragraphs.items[0];
        firstPara.load("font/name");
        await context.sync();

        if (firstPara.font.name) {
          cachedDocumentFont = firstPara.font.name;
          console.log(`Detected document font: ${cachedDocumentFont}`);
        }
      }
    });
  } catch (error) {
    console.warn("Could not detect document font, using default:", error);
  }
  return cachedDocumentFont;
}

// ==================== HELPER FUNCTIONS ====================

/**
 * Converts markdown to Word-compatible HTML.
 * Ensures proper formatting for Word's HTML parser.
 */
function markdownToWordHtml(markdown) {
  console.log(`[markdownToWordHtml] Processing content: ${markdown.substring(0, 50)}...`);
  if (!markdown) return "";

  // Pre-process underline and strikethrough (marked native GFM might be disabled or fail in some environments)
  const processedMarkdown = markdown
    .replace(/\+\+(.+?)\+\+/g, '<u>$1</u>')
    .replace(/~~(.+?)~~/g, '<s>$1</s>');

  // Parse markdown to HTML using marked library
  let html = marked.parse(processedMarkdown);

  // === TABLE FORMATTING ===
  // Word requires explicit styling for tables to render properly with borders
  if (html.includes('<table>')) {
    html = html.replace(/<table>/g, '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #000;">');
    // Add styling to table cells for better appearance
    html = html.replace(/<th>/g, '<th style="border: 1px solid #000; padding: 8px; background-color: #f0f0f0; font-weight: bold;">');
    html = html.replace(/<td>/g, '<td style="border: 1px solid #000; padding: 8px;">');
  }

  // === ORDERED LIST FORMATTING ===
  // CRITICAL: Word's HTML parser can render <ol> as bullets without explicit styling
  // Adding list-style-type CSS ensures proper numbered list rendering
  if (html.includes('<ol>')) {
    // First, handle any already-styled ordered lists (from nested replacements)
    // to avoid double-replacing them
    html = html.replace(/<ol>/g, '<ol style="list-style-type: decimal; margin-left: 0; padding-left: 40px; margin-bottom: 10px;">');

    // Handle nested ordered lists with different numbering styles
    // Match <ol> tags that are inside <li> elements (nested lists)
    // Use lower-alpha (a, b, c) for first nesting level
    html = html.replace(/<li>([^<]*)<ol style="list-style-type: decimal;/g, '<li>$1<ol style="list-style-type: lower-alpha;');

    // For third-level nesting, use lower-roman (i, ii, iii)
    html = html.replace(/<li>([^<]*)<ol style="list-style-type: lower-alpha;([^>]*)>([^<]*)<li>([^<]*)<ol style="list-style-type: lower-alpha;/g,
      '<li>$1<ol style="list-style-type: lower-alpha;$2>$3<li>$4<ol style="list-style-type: lower-roman;');
  }

  // === UNORDERED LIST FORMATTING ===
  // Ensure <ul> has explicit bullet styling to distinguish from ordered lists
  if (html.includes('<ul>')) {
    html = html.replace(/<ul>/g, '<ul style="list-style-type: disc; margin-left: 0; padding-left: 40px; margin-bottom: 10px;">');

    // Nested unordered lists should use circle then square markers
    html = html.replace(/<li>([^<]*)<ul style="list-style-type: disc;/g, '<li>$1<ul style="list-style-type: circle;');
    html = html.replace(/<li>([^<]*)<ul style="list-style-type: circle;/g, '<li>$1<ul style="list-style-type: square;');
  }

  // === LIST ITEM FORMATTING ===
  // Add spacing to list items for better readability
  html = html.replace(/<li>/g, '<li style="margin-bottom: 5px;">');

  // === TRAILING PARAGRAPH ===
  // Add a paragraph with non-breaking space after lists to ensure proper formatting
  // This fixes the issue where the last list item may not be properly numbered/formatted
  // Using &nbsp; instead of empty paragraph for better Word compatibility
  html = html.replace(/<\/ol>/g, '</ol><p>&nbsp;</p>');
  html = html.replace(/<\/ul>/g, '</ul><p>&nbsp;</p>');

  // === FONT CONSISTENCY ===
  // Wrap with explicit font-family using cached document font
  // This ensures inserted content matches the document's existing font
  html = `<span style="font-family: '${cachedDocumentFont}', Calibri, sans-serif;">${html}</span>`;

  return html;
}

/**
 * Converts markdown to Word-compatible HTML for inline content (no wrapping <p> tags).
 * Use this for modify_text replacements.
 */
function markdownToWordHtmlInline(markdown) {
  if (!markdown) return "";

  // Use parseInline to avoid wrapping in <p> tags for simple text
  // But if there are block elements (lists, tables), use full parse
  const hasBlockElements = /(\n[-*+]\s|\n\d+\.\s|\|.*\|.*\n|^#{1,6}\s)/m.test(markdown);

  if (hasBlockElements) {
    return markdownToWordHtml(markdown);
  }

  // For inline content, pre-process underline/strike and use parseInline
  const processedMarkdown = markdown
    .replace(/\+\+(.+?)\+\+/g, '<u>$1</u>')
    .replace(/~~(.+?)~~/g, '<s>$1</s>');
  return `<span style="font-family: '${cachedDocumentFont}', Calibri, sans-serif;">${marked.parseInline(processedMarkdown)}</span>`;
}

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

// --- Scroll-to-Bottom Button ---
function setupScrollToBottom() {
  const chatMessages = document.getElementById("chat-messages");
  const scrollBtn = document.getElementById("scroll-to-bottom");

  if (!chatMessages || !scrollBtn) return;

  // Show/hide button based on scroll position
  chatMessages.addEventListener("scroll", () => {
    const isNearBottom = chatMessages.scrollHeight - chatMessages.scrollTop - chatMessages.clientHeight < 100;
    scrollBtn.classList.toggle("visible", !isNearBottom);
  });

  // Scroll to bottom on click
  scrollBtn.onclick = () => {
    chatMessages.scrollTo({
      top: chatMessages.scrollHeight,
      behavior: "smooth"
    });
  };
}

// --- Typing Indicator Helper ---
function createTypingIndicator(color = 'teal', showCancelButton = false) {
  const container = document.createElement("div");
  container.className = "chat-message system animate-entry";
  const colorClass = color === 'yellow' ? 'typing-yellow' : 'typing-teal';

  let cancelButtonHtml = '';
  if (showCancelButton) {
    cancelButtonHtml = `
      <button class="cancel-request-btn" title="Cancel request">
        <span class="cancel-icon">✕</span>
      </button>
    `;
  }

  container.innerHTML = `
    <div class="typing-container">
      <span class="typing-indicator ${colorClass}">
        <span class="dot"></span>
        <span class="dot"></span>
        <span class="dot"></span>
      </span>
      ${cancelButtonHtml}
    </div>
  `;

  // Attach cancel button event listener
  if (showCancelButton) {
    const cancelBtn = container.querySelector('.cancel-request-btn');
    if (cancelBtn) {
      cancelBtn.onclick = () => {
        if (currentRequestController) {
          currentRequestController.abort();
          console.log('User cancelled request');
        }
      };
    }
  }

  return container;
}


// --- Shake Input on Error ---
function shakeInput() {
  const chatInput = document.getElementById("chat-input");
  chatInput.classList.add("shake");
  setTimeout(() => {
    chatInput.classList.remove("shake");
  }, 400);
}


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
  sanitizeHistory();

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

function addMessageToChat(sender, message, checkpointIndex = -1) {
  const chatMessages = document.getElementById("chat-messages");
  const messageElement = document.createElement("div");
  // Add base class and specific sender class
  // Add animate-entry for slide-up animation
  messageElement.className = `chat-message ${sender.toLowerCase()} animate-entry`;


  const isSystem = sender === "System" || sender === "Error";

  if (isSystem) {
    renderSystemMessageContent(messageElement, sender, message);
  } else {
    // Render Markdown for user/gemini
    messageElement.innerHTML = `<strong>${sender}:</strong> <div>${marked.parse(message)}</div>`;
  }

  // Add Revert button if a valid checkpoint index is provided
  if (checkpointIndex !== -1) {
    addUndoButton(messageElement, checkpointIndex);
  }

  chatMessages.appendChild(messageElement);
  chatMessages.scrollTop = chatMessages.scrollHeight; // Auto-scroll
  return messageElement; // Return element for potential removal
}

function updateSystemMessage(messageElement, newMessage, checkpointIndex = -1) {
  if (!messageElement) return;

  // Save existing button container before replacing content
  // This preserves the revert button when called with just status text
  const existingBtnContainer = messageElement.querySelector(".revert-btn-container");

  // Update content (this replaces innerHTML, destroying any existing button)
  renderSystemMessageContent(messageElement, "System", newMessage);

  // Update/Add Undo button
  if (checkpointIndex !== -1) {
    // New checkpoint: add fresh button (any saved container is replaced)
    addUndoButton(messageElement, checkpointIndex);
  } else if (existingBtnContainer) {
    // No new checkpoint but had existing button: restore it
    messageElement.appendChild(existingBtnContainer);
  }
}

function renderSystemMessageContent(element, sender, message) {
  const maxLength = 120; // Character limit for system messages
  if (message.length > maxLength) {
    const shortText = message.substring(0, maxLength) + "...";
    const fullText = message;

    element.innerHTML = `<strong>${sender}:</strong> `;

    const textSpan = document.createElement("span");
    textSpan.innerText = shortText;
    element.appendChild(textSpan);

    const toggleBtn = document.createElement("button");
    toggleBtn.innerText = "Show more";
    toggleBtn.className = "system-msg-toggle";
    toggleBtn.onclick = () => {
      if (toggleBtn.innerText === "Show more") {
        textSpan.innerText = fullText;
        toggleBtn.innerText = "Show less";
      } else {
        textSpan.innerText = shortText;
        toggleBtn.innerText = "Show more";
      }
    };
    element.appendChild(toggleBtn);
  } else {
    // Render Markdown inline for System messages
    element.innerHTML = `<strong>${sender}:</strong> <span>${marked.parseInline(message)}</span>`;
  }
}

function addUndoButton(messageElement, checkpointIndex) {
  const buttonContainer = document.createElement("div");
  buttonContainer.className = "revert-btn-container";
  const revertBtn = document.createElement("button");
  revertBtn.innerHTML = "<span>&#8634;</span> Revert changes"; // ↺ clockwise open circle arrow
  revertBtn.className = "revert-checkpoint-btn";
  revertBtn.title = "Undo changes made by this action";
  revertBtn.onclick = () => restoreCheckpoint(checkpointIndex);

  buttonContainer.appendChild(revertBtn);
  messageElement.appendChild(buttonContainer);
}

function addRetryButton(messageElement, originalMessage) {
  const buttonContainer = document.createElement("div");
  buttonContainer.className = "revert-btn-container retry-btn-container";
  const retryBtn = document.createElement("button");
  retryBtn.innerHTML = "<span>&#8635;</span> Retry"; // ↻ counter-clockwise open circle arrow
  retryBtn.className = "revert-checkpoint-btn retry-request-btn";
  retryBtn.title = "Paste message back into input";
  retryBtn.onclick = () => {
    // Paste the message into the input field
    const chatInput = document.getElementById("chat-input");
    if (chatInput) {
      chatInput.value = originalMessage;
      chatInput.focus();
    }
    // Hide the retry button after clicking
    buttonContainer.style.display = "none";
  };

  buttonContainer.appendChild(retryBtn);
  messageElement.appendChild(buttonContainer);
}

function hideAllRetryButtons() {
  const retryContainers = document.querySelectorAll(".retry-btn-container");
  retryContainers.forEach(container => {
    container.style.display = "none";
  });
}

function removeMessage(messageElement) {
  if (messageElement && messageElement.parentNode) {
    messageElement.parentNode.removeChild(messageElement);
  }
}

/**
 * Agentic Tool: Applies redlines based on an instruction using Structural Anchoring.
 */
async function executeRedline(instruction, fullDocumentText) {
  // Check for API key
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    // Detect document font for consistent HTML insertion
    await detectDocumentFont();
    // 1. Build the prompt for the diff generator
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of precise changes to be made, referencing the paragraph numbers.

CRITICAL: Return ONLY valid JSON. Do NOT include explanatory text, notes, or duplicate entries.

Each change must be an object with the following structure:
- "paragraphIndex": The integer number of the paragraph to modify (e.g., 1 for [P1]). For "replace_range", this is the START paragraph.
- "endParagraphIndex": (Only for "replace_range") The integer number of the END paragraph (inclusive).
- "operation": "edit_paragraph", "replace_paragraph", "modify_text", or "replace_range".
- "newContent": (For "edit_paragraph" ONLY) The complete rewritten paragraph content. The system will automatically compute precise word-level changes.
- "content": (For "replace_paragraph" and "replace_range" ONLY) The new content to insert.
- "originalText": (For "modify_text" ONLY) The specific text snippet within the paragraph to find and replace. **MAX 80 characters**.
- "replacementText": (For "modify_text" ONLY) The new text to replace "originalText" with.

**MARKDOWN FORMATTING (VERY IMPORTANT)**:
All content and replacementText values support Markdown formatting. Use these when the user requests formatting:
- **Bold**: Use **text** (double asterisks)
- *Italic*: Use *text* (single asterisks)
- **Underline**: Use ++text++ (double pluses)
- ~~Strikethrough~~: Use ~~text~~ (double tildes)
- ***Bold Italic***: Use ***text*** (triple asterisks)
- **Unordered/Bullet lists**: Use "- item" or "* item" on separate lines. These render as bullet points (•).
- **Ordered/Numbered lists**: Use "1. item", "2. item" on separate lines. These render as 1, 2, 3...
- **Alphabetical lists (A, B, C)**: Use "A. item", "B. item" on separate lines. Use lowercase "a. item" for a, b, c. Use "I.", "II." for roman numerals.
- Line breaks: Use actual newlines (\\n) in the text
- Tables: Use GitHub-style markdown tables:
  | Header 1 | Header 2 |
  |----------|----------|
  | Cell 1   | Cell 2   |
- Headings: Use # for H1, ## for H2, ### for H3

**CRITICAL LIST FORMATTING RULES**:
- **PRESERVE HIERARCHY**: If the document uses nested numbering (1.1, 1.1.1, etc.), ALWAYS use that same hierarchical format in your changes. **Do NOT flatten nested lists** into simple numbered lists (1., 2., 3.) unless specifically asked to restructure the hierarchy.
- **INCLUDE MARKERS**: Always include the correct list marker (e.g., "1.1.1 ") at the start of your \`newContent\` or \`content\` for list items. The system will use these to correctly set the indentation level in Word, and then it will automatically strip them from the final text.
- **NO MIXING**: NEVER mix bullet markers with manual numbering like "• (a)" or "- 1." - this creates malformed output
- **MARKDOWN SYNTAX**: 
  - For bullets: use "- " or "* "
  - For simple numbers: use "1. ", "2. "
  - For hierarchical numbers: use "1.1. ", "1.1.1. "
- **STRIPPING**: When converting existing lists, REMOVE the original markers from your response and use ONLY the markdown syntax described above.

When the user asks for formatted content (bullets, tables, bold, etc.), ALWAYS use the appropriate Markdown syntax.

Rules:
- **PRIORITIZE \`edit_paragraph\`**: This is the NEW preferred method. For ANY text edit (small or large), use \`edit_paragraph\` with the complete rewritten paragraph. The system will automatically compute precise word-level changes using diff-match-patch. This is more reliable than \`modify_text\`.
- Use "edit_paragraph" for ALL text edits: spelling changes, word replacements, sentence rewrites, or even 60% paragraph rewrites. Just provide the full new paragraph content.
- Use "replace_paragraph" only when you need to replace with complex formatted content (lists, tables, headings) that requires HTML insertion.
- Use "modify_text" ONLY as a fallback for very specific surgical edits where you need to target exact substrings.
- **CRITICAL LENGTH LIMIT**: For "modify_text", "originalText" MUST be **80 characters or fewer**. This is a hard limit.
- Use "replace_range" when you need to replace multiple consecutive paragraphs (like converting a bulleted list to a single paragraph).
- For "replace_range", provide ONLY "paragraphIndex", "endParagraphIndex", "operation", and "content". Do NOT include "originalText" or "replacementText".
- For "edit_paragraph", provide ONLY "paragraphIndex", "operation", and "newContent".
- For "modify_text", "originalText" must match EXACTLY text found within that specific paragraph.
- Do NOT include the [P#] marker in any content fields.
- Return ONLY ONE change per unique text location. Do NOT create duplicate entries.

IMPORTANT: This document may contain existing tracked changes. The text shown represents the "accepted" state (as if all changes were accepted). Your changes will be applied as additional tracked changes on top of existing ones.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

Return ONLY the JSON array, nothing else:`;

    // 2. Call Gemini to get the JSON array of changes
    const aiChanges = await callGeminiForDiffs(fullPrompt);

    console.log("AI Suggested Changes (raw):", aiChanges);

    if (!aiChanges || !Array.isArray(aiChanges)) {
      return {
        message: "AI did not return a valid list of changes. Please check the console logs for details.",
        showToUser: false  // Silent error - let the model handle it
      };
    }

    if (aiChanges.length === 0) {
      return {
        message: "AI had no changes to suggest based on the instruction.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    let changesApplied = 0;
    let changeTrackingAvailable = false;
    const redlineEnabled = loadRedlineSetting();

    // 3. Apply changes in Word
    await Word.run(async (context) => {
      // Enable Track Changes only if redline setting is enabled
      let originalChangeTrackingMode = null;
      changeTrackingAvailable = false;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          changeTrackingAvailable = true;
          originalChangeTrackingMode = doc.changeTrackingMode;

          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
          }
        } catch (trackError) {
          console.error("Track Changes not available:", trackError);
          changeTrackingAvailable = false;
        }
      }

      // Load paragraphs to map indices
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // Track the current paragraph count (may change as we add/remove paragraphs)
      let currentParagraphCount = paragraphs.items.length;

      for (const change of aiChanges) {
        try {
          console.log("Processing change:", JSON.stringify(change));

          const pIndex = change.paragraphIndex - 1; // 0-based index

          // Check if this is an insertion at the end (index equals or exceeds paragraph count)
          // We're lenient here - any index beyond current count is treated as an append
          const isInsertAtEnd = pIndex >= currentParagraphCount;

          // Only reject negative indices - positive ones that exceed count are handled as appends
          if (pIndex < 0) {
            console.warn(`Invalid paragraph index (negative): ${change.paragraphIndex}`);
            continue;
          }

          // For out-of-bounds indices, reload paragraphs and check again
          if (pIndex >= paragraphs.items.length) {
            // Reload paragraphs collection to get any newly added ones
            paragraphs.load("items");
            await context.sync();
            currentParagraphCount = paragraphs.items.length;

            // If still out of bounds after reload, treat as append to last paragraph
            if (pIndex >= paragraphs.items.length) {
              console.log(`Paragraph index ${change.paragraphIndex} exceeds count (${paragraphs.items.length}), treating as append`);
            }
          }

          // For insertions at the end, use the last paragraph as reference
          const targetParagraph = (pIndex >= paragraphs.items.length)
            ? paragraphs.items[paragraphs.items.length - 1]
            : paragraphs.items[pIndex];

          if (change.operation === "edit_paragraph") {
            console.log(`Editing Paragraph ${change.paragraphIndex} with DMP`);

            if (!change.newContent) {
              console.warn("No newContent provided for edit_paragraph. Skipping.");
              continue;
            }

            try {
              // If inserting at end, insert new paragraph instead of editing
              if (isInsertAtEnd) {
                console.log(`Inserting new paragraph after paragraph ${paragraphs.items.length}`);
                targetParagraph.insertParagraph(change.newContent, "After");
                await context.sync(); // Sync immediately to ensure tracked changes captures the insertion
                changesApplied++;
              } else {
                // Route through our smart operation router
                await routeChangeOperation(change, targetParagraph, context);
                changesApplied++;
              }
            } catch (error) {
              console.error(`Error editing paragraph ${change.paragraphIndex}:`, error);
              // Fallback to old modify_text approach if DMP fails
            }

          } else if (change.operation === "replace_paragraph") {
            console.log(`Replacing Paragraph ${change.paragraphIndex}`);

            if (change.content === null || change.content === undefined) {
              console.warn("Content is null/undefined for replace_paragraph. Skipping.");
              continue;
            }

            // Normalize content: Convert literal escape sequences to actual characters
            // This handles cases where the AI returns "\\n" as a two-character string instead of actual newlines
            let normalizedContent = normalizeContentEscapes(change.content || "");

            // --- NEW: Detect if target paragraph is already a list item ---
            // If so, we need to preserve its numId/ilvl when replacing content
            let targetIsListItem = false;
            let targetListContext = null;

            if (!isInsertAtEnd) {
              try {
                const targetOoxmlResult = targetParagraph.getOoxml();
                await context.sync();

                // Check for w:numPr in the paragraph's OOXML
                const numPrMatch = targetOoxmlResult.value.match(/<w:numPr>.*?<w:ilvl w:val="(\d+)".*?<w:numId w:val="(\d+)".*?<\/w:numPr>/s);
                if (numPrMatch) {
                  targetIsListItem = true;
                  targetListContext = {
                    ooxml: targetOoxmlResult.value,
                    ilvl: numPrMatch[1],
                    numId: numPrMatch[2]
                  };
                  console.log(`[replace_paragraph] Target P${change.paragraphIndex} is list item: numId=${targetListContext.numId}, ilvl=${targetListContext.ilvl}`);
                }
              } catch (ooxmlError) {
                console.warn("[replace_paragraph] Could not check list context:", ooxmlError);
              }
            }

            // If target is a list item and content is plain text (no list markers), 
            // use OOXML reconciliation to preserve list formatting
            const contentHasListMarkers = /^(\s*)([-*•]|\d+\.|[a-zA-Z]\.|[ivxlcIVXLC]+\.|\d+\.\d+\.?)\s+/m.test(normalizedContent);
            console.log(`[replace_paragraph] contentHasListMarkers: ${contentHasListMarkers}`);

            if (targetIsListItem && !contentHasListMarkers) {
              console.log(`[replace_paragraph] Preserving list context for plain text edit`);

              try {
                const redlineEnabled = loadRedlineSetting();
                const redlineAuthor = loadRedlineAuthor();

                // Get original text for diffing
                const originalText = targetParagraph.text;
                await context.sync();

                // Use OOXML reconciliation to preserve numPr
                const result = await applyRedlineToOxml(
                  targetListContext.ooxml,
                  originalText,
                  normalizedContent,
                  {
                    author: redlineEnabled ? redlineAuthor : undefined,
                    generateRedlines: redlineEnabled
                  }
                );

                if (result.oxml && result.hasChanges) {
                  const doc = context.document;
                  doc.load("changeTrackingMode");
                  await context.sync();

                  const originalMode = doc.changeTrackingMode;
                  if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                    doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                    await context.sync();
                  }

                  try {
                    targetParagraph.insertOoxml(result.oxml, "Replace");
                    await context.sync();
                    console.log("✅ OOXML list-preserving edit successful");
                    changesApplied++;
                  } finally {
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = originalMode;
                      await context.sync();
                    }
                  }
                  continue; // Skip other handlers
                }
              } catch (listPreserveError) {
                console.warn("[replace_paragraph] List preservation failed, falling back:", listPreserveError);
                // Fall through to standard handlers
              }
            }
            // --- END NEW ---

            // Check if this is a list - use OOXML pipeline for proper redlines
            const listData = parseMarkdownList(normalizedContent);
            console.log(`[replace_paragraph] listData result: type=${listData?.type}, items=${listData?.items?.length}`);
            if (listData && listData.type !== 'text') {
              console.log(`Detected ${listData.type} list in replace_paragraph, using OOXML pipeline`);
              try {
                // Get original paragraph info for proper diff/redlines
                // Only get original text if we're REPLACING (not appending)
                let originalTextForDeletion = '';
                if (!isInsertAtEnd) {
                  const originalText = targetParagraph.text;
                  await context.sync();
                  originalTextForDeletion = originalText;
                }

                // Create reconciliation pipeline with redline settings
                const redlineEnabled = loadRedlineSetting();
                const redlineAuthor = loadRedlineAuthor();
                const pipeline = new ReconciliationPipeline({
                  generateRedlines: redlineEnabled,
                  author: redlineAuthor,
                  font: cachedDocumentFont // Ensure font consistency
                });

                // Execute list generation - this creates OOXML with w:ins/w:del track changes
                const result = await pipeline.executeListGeneration(
                  normalizedContent,
                  null, // numberingContext - let pipeline determine
                  null, // originalRunModel - not available here
                  originalTextForDeletion // Only pass original text if replacing, not appending
                );

                console.log(`[ListGen] Generated ${result.ooxml.length} bytes of OOXML, isInsertAtEnd=${isInsertAtEnd}`);

                if (result.ooxml && result.isValid) {
                  // Wrap in document fragment for insertOoxml
                  const wrappedOoxml = wrapInDocumentFragment(result.ooxml, {
                    includeNumbering: true,
                    numberingXml: result.numberingXml // Crucial for A, B, C styles
                  });

                  // Temporarily disable Word's track changes to avoid double-tracking
                  // Our w:ins/w:del ARE the track changes
                  const doc = context.document;
                  doc.load("changeTrackingMode");
                  await context.sync();

                  const originalMode = doc.changeTrackingMode;
                  if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                    doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                    await context.sync();
                  }

                  try {
                    // Use 'After' if appending at end, 'Replace' if replacing existing paragraph
                    const insertMode = isInsertAtEnd ? 'After' : 'Replace';
                    console.log(`[ListGen] Using insert mode: ${insertMode}`);
                    targetParagraph.insertOoxml(wrappedOoxml, insertMode);
                    await context.sync();
                    console.log(`✅ OOXML list generation successful`);
                    changesApplied++;
                  } finally {
                    // Restore track changes mode
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = originalMode;
                      await context.sync();
                    }
                  }
                } else {
                  console.warn('[ListGen] Pipeline returned invalid result, falling back to HTML');
                  const htmlContent = markdownToWordHtml(normalizedContent);
                  const insertLocation = isInsertAtEnd ? "After" : "Replace";
                  targetParagraph.insertHtml(htmlContent, insertLocation);
                  changesApplied++;
                }
              } catch (listError) {
                console.error(`Error in OOXML list generation:`, listError);
                // Fallback to HTML if OOXML fails
                const htmlContent = markdownToWordHtml(normalizedContent);
                const insertLocation = isInsertAtEnd ? "After" : "Replace";
                targetParagraph.insertHtml(htmlContent, insertLocation);
                changesApplied++;
              }
              // Skip the rest of replace_paragraph handling
              continue;
            }

            // Check if this is a table - use OOXML pipeline
            const matchedTable = normalizedContent.includes('|');
            if (matchedTable) {
              const tableData = parseTable(normalizedContent);
              if (tableData.rows.length > 0 || tableData.headers.length > 0) {
                console.log(`Detected table in replace_paragraph, using OOXML pipeline`);
                try {
                  // Create reconciliation pipeline with redline settings
                  const redlineEnabled = loadRedlineSetting();
                  const redlineAuthor = loadRedlineAuthor();
                  const pipeline = new ReconciliationPipeline({
                    generateRedlines: redlineEnabled,
                    author: redlineAuthor
                  });

                  // Execute table generation - this creates OOXML with w:tbl and optional w:ins
                  const result = pipeline.executeTableGeneration(normalizedContent);

                  if (result.ooxml && result.isValid) {
                    // Wrap in document fragment
                    const wrappedOoxml = wrapInDocumentFragment(result.ooxml, {
                      includeNumbering: false
                    });

                    // Disable track changes temporarily
                    const doc = context.document;
                    doc.load("changeTrackingMode");
                    await context.sync();

                    const originalMode = doc.changeTrackingMode;
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                      await context.sync();
                    }

                    try {
                      const insertMode = isInsertAtEnd ? 'After' : 'Replace';
                      console.log(`[TableGen] Using insert mode: ${insertMode}`);
                      targetParagraph.insertOoxml(wrappedOoxml, insertMode);
                      await context.sync();
                      console.log(`✅ OOXML table generation successful`);
                      changesApplied++;
                    } finally {
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = originalMode;
                        await context.sync();
                      }
                    }
                  } else {
                    console.warn('[TableGen] Pipeline failed, falling back to HTML');
                    const htmlContent = markdownToWordHtml(normalizedContent);
                    targetParagraph.insertHtml(htmlContent, isInsertAtEnd ? "After" : "Replace");
                    changesApplied++;
                  }
                } catch (tableError) {
                  console.error(`Error in OOXML table generation:`, tableError);
                  const htmlContent = markdownToWordHtml(normalizedContent);
                  targetParagraph.insertHtml(htmlContent, isInsertAtEnd ? "After" : "Replace");
                  changesApplied++;
                }
                // Skip the rest of replace_paragraph handling
                continue;
              }
            }

            // Convert Markdown to Word-compatible HTML for regular content
            let htmlContent = "";
            try {
              htmlContent = markdownToWordHtml(normalizedContent);
            } catch (markedError) {
              console.error("Error parsing markdown:", markedError);
              htmlContent = normalizedContent; // Fallback to raw text
            }

            // Strip wrapping <p> if present to avoid double paragraphs if Word handles it
            // But only if it's a single simple paragraph (no block elements inside)
            const trimmed = htmlContent.trim();
            const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
              trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
              !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
              !trimmed.includes('<table') && !trimmed.includes('<h');

            if (hasSingleParagraph) {
              htmlContent = trimmed.substring(3, trimmed.length - 4);
            }

            try {
              // If inserting at end, use insertParagraph to add new content after
              if (isInsertAtEnd) {
                console.log(`Inserting new paragraph after paragraph ${paragraphs.items.length}`);
                // Use insertParagraph to add new paragraph after the last one
                const newPara = targetParagraph.insertParagraph(normalizedContent, "After");
                await context.sync(); // Sync immediately to ensure tracked changes captures the insertion
                changesApplied++;
              } else {
                targetParagraph.insertHtml(htmlContent, "Replace");
                changesApplied++;
              }
            } catch (wordError) {
              console.error(`Error replacing paragraph ${change.paragraphIndex}:`, wordError);
            }

          } else if (change.operation === "replace_range") {
            const endIndex = change.endParagraphIndex - 1;
            if (endIndex < 0 || endIndex >= paragraphs.items.length || endIndex < pIndex) {
              console.warn(`Invalid end paragraph index: ${change.endParagraphIndex}`);
              continue;
            }

            console.log(`Replacing Range from P${change.paragraphIndex} to P${change.endParagraphIndex}`);

            try {
              const startPara = paragraphs.items[pIndex];
              const endPara = paragraphs.items[endIndex];

              // Check if we are inside a table - wrap in try/catch for safety
              let startHasTable = false;
              let endHasTable = false;
              try {
                startPara.load("parentTable/id");
                endPara.load("parentTable/id");
                await context.sync();
                startHasTable = !startPara.parentTable.isNullObject;
                endHasTable = !endPara.parentTable.isNullObject;
              } catch (tableCheckError) {
                console.warn("Could not check for table context:", tableCheckError);
                // Continue without table detection
              }

              let targetRange = null;
              let isTableReplacement = false;
              let tableToDelete = null;

              // If both start and end are in the same table
              if (startHasTable && endHasTable) {
                try {
                  const startTable = startPara.parentTable;
                  const endTable = endPara.parentTable;

                  if (startTable.id === endTable.id) {
                    console.log("Detected same table context. Will replace entire table.");
                    // Strategy: Insert AFTER the table, then delete the table.
                    // This avoids GeneralException when replacing complex structures directly.
                    targetRange = startTable.getRange();
                    isTableReplacement = true;
                    tableToDelete = startTable;
                  } else {
                    console.warn("Start and End paragraphs are in DIFFERENT tables. Falling back to standard range expansion.");
                    targetRange = startPara.getRange().expandTo(endPara.getRange());
                  }
                } catch (tableError) {
                  console.warn("Error handling table replacement, falling back to range:", tableError);
                  targetRange = startPara.getRange().expandTo(endPara.getRange());
                }
              } else {
                // Create a range covering both
                targetRange = startPara.getRange().expandTo(endPara.getRange());
              }

              // Use 'content' field for replace_range (not replacementText)
              const contentToParse = change.content || change.replacementText || "";

              if (!contentToParse || contentToParse.trim().length === 0) {
                console.warn("Empty content for replace_range. Skipping.");
                continue;
              }

              // --- NEW: Detect list structures and use OOXML engine for proper numPr ---
              const hasListMarkers = /^(\s*)([-*•]|\d+\.|[a-zA-Z]\.|[ivxlcIVXLC]+\.|\d+\.\d+\.?)\s+/m.test(contentToParse);

              if (hasListMarkers && !isTableReplacement) {
                console.log("[replace_range] Detected list markers, using OOXML reconciliation");

                try {
                  // Get the original text from the range for diffing
                  targetRange.load("text");
                  const originalOoxmlResult = startPara.getOoxml(); // Get OOXML from first paragraph
                  await context.sync();

                  const originalText = targetRange.text || "";
                  const redlineEnabled = loadRedlineSetting();
                  const redlineAuthor = loadRedlineAuthor();

                  // Use the OOXML engine for proper list generation
                  const result = await applyRedlineToOxml(
                    originalOoxmlResult.value,
                    originalText,
                    contentToParse,
                    {
                      author: redlineEnabled ? redlineAuthor : undefined,
                      generateRedlines: redlineEnabled
                    }
                  );

                  if (result.oxml && result.hasChanges) {
                    // Temporarily disable track changes to avoid double-tracking
                    const doc = context.document;
                    doc.load("changeTrackingMode");
                    await context.sync();

                    const originalMode = doc.changeTrackingMode;
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                      await context.sync();
                    }

                    try {
                      targetRange.insertOoxml(result.oxml, "Replace");
                      await context.sync();
                      changesApplied++;
                      console.log("✅ OOXML list reconciliation successful for replace_range");
                    } finally {
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = originalMode;
                        await context.sync();
                      }
                    }
                    continue; // Skip HTML fallback
                  }
                } catch (ooxmlError) {
                  console.warn("[replace_range] OOXML reconciliation failed, falling back to HTML:", ooxmlError);
                  // Fall through to HTML path
                }
              }
              // --- END NEW ---

              // Convert Markdown to Word-compatible HTML (fallback for non-list or table content)
              let htmlContent = "";
              try {
                htmlContent = markdownToWordHtml(contentToParse);
              } catch (markedError) {
                console.error("Error parsing markdown for range:", markedError);
                htmlContent = contentToParse;
              }

              if (isTableReplacement && tableToDelete) {
                // Insert AFTER the table
                if (htmlContent && htmlContent.trim().length > 0) {
                  targetRange.insertHtml(htmlContent, "After");
                }
                // Delete the old table
                tableToDelete.delete();
                changesApplied++;
              } else if (targetRange) {
                // Standard replacement
                try {
                  targetRange.insertHtml(htmlContent, "Replace");
                  changesApplied++;
                } catch (replaceError) {
                  console.warn("Standard insertHtml failed. Trying fallback (Clear + InsertStart).", replaceError);
                  // Fallback: Clear and insert at start
                  try {
                    targetRange.clear(); // Clears content but keeps range
                    targetRange.insertHtml(htmlContent, "Start");
                    changesApplied++;
                  } catch (fallbackError) {
                    console.warn("Fallback (Clear+InsertStart) failed. Trying Nuclear Option (InsertText+InsertHtml).", fallbackError);
                    // Fallback 2: Nuke with text first to reset formatting
                    try {
                      // Replace with a placeholder to reset structure
                      const tempRange = targetRange.insertText(" ", "Replace");
                      tempRange.insertHtml(htmlContent, "Replace");
                      changesApplied++;
                    } catch (nuclearError) {
                      console.error("Replacement failed:", nuclearError);
                    }
                  }
                }
              }
            } catch (rangeError) {
              console.error(`Error replacing range P${change.paragraphIndex}-P${change.endParagraphIndex}:`, rangeError);
            }
          } else if (change.operation === "modify_text") {
            console.log(`Modifying text in Paragraph ${change.paragraphIndex}: "${change.originalText}" -> "${change.replacementText}"`);

            // Safety check for search string length - Word API has strict limits
            const fullOriginalText = change.originalText;
            if (!fullOriginalText || fullOriginalText.length === 0) {
              console.warn(`Empty search text for modify_text in Paragraph ${change.paragraphIndex}. Skipping.`);
              continue;
            }

            // Word's search API has a practical limit of around 80 characters
            const MAX_SEARCH_LENGTH = 80;
            const needsRangeExpansion = fullOriginalText.length > MAX_SEARCH_LENGTH;
            const searchText = needsRangeExpansion
              ? fullOriginalText.substring(0, MAX_SEARCH_LENGTH)
              : fullOriginalText;

            if (needsRangeExpansion) {
              console.warn(`Search text too long (${fullOriginalText.length} chars), using range expansion strategy.`);
            }

            try {
              // Search ONLY within this paragraph
              const searchResults = targetParagraph.search(searchText, { matchCase: true });
              searchResults.load("items");
              await context.sync();

              if (searchResults.items.length > 0) {
                // Apply to first match only when using range expansion (to avoid ambiguity)
                const matchesToProcess = needsRangeExpansion ? [searchResults.items[0]] : searchResults.items;

                for (const item of matchesToProcess) {
                  const replacementText = change.replacementText || "";
                  let htmlReplacement = "";
                  try {
                    // Use inline parsing for modify_text to avoid wrapping in <p> tags
                    // unless the content has block elements
                    htmlReplacement = markdownToWordHtmlInline(replacementText);
                  } catch (markedError) {
                    console.error("Error parsing markdown for modify_text:", markedError);
                    htmlReplacement = replacementText;
                  }

                  // Strip wrapping <p> for simple inline content
                  const trimmed = htmlReplacement.trim();
                  const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
                    trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
                    !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
                    !trimmed.includes('<table') && !trimmed.includes('<h');

                  if (hasSingleParagraph) {
                    htmlReplacement = trimmed.substring(3, trimmed.length - 4);
                  }

                  try {
                    if (needsRangeExpansion) {
                      // Expand the range to cover the full original text length
                      // Strategy: Find a short suffix from the END of the original text,
                      // then expand the range from prefix start to suffix end
                      const foundRange = item.getRange();

                      try {
                        // Take the LAST 60 chars of the original text as our suffix search
                        // This must be short enough for Word's search API
                        const SUFFIX_LENGTH = 60;
                        const suffixStart = Math.max(0, fullOriginalText.length - SUFFIX_LENGTH);
                        const suffixText = fullOriginalText.substring(suffixStart);

                        console.log(`Range expansion: searching for suffix "${suffixText.substring(0, 30)}..." (${suffixText.length} chars)`);

                        if (suffixText.length >= 5 && suffixText.length <= 80) {
                          const suffixResults = targetParagraph.search(suffixText, { matchCase: true });
                          suffixResults.load("items");
                          await context.sync();

                          if (suffixResults.items.length > 0) {
                            // Find the suffix match that comes after our prefix match
                            // by expanding from the found prefix to each suffix candidate
                            let expandedSuccessfully = false;

                            for (const suffixMatch of suffixResults.items) {
                              try {
                                // Expand from found prefix start to suffix end
                                const expandedRange = foundRange.expandTo(suffixMatch.getRange("End"));
                                expandedRange.load("text");
                                await context.sync();

                                // Verify the expanded range roughly matches the original length
                                // Allow some tolerance for whitespace differences
                                const expandedLength = expandedRange.text.length;
                                const originalLength = fullOriginalText.length;
                                const tolerance = Math.max(10, originalLength * 0.1);

                                if (Math.abs(expandedLength - originalLength) <= tolerance) {
                                  console.log(`Expanded range matches: ${expandedLength} chars (original: ${originalLength})`);
                                  // Use insertHtml with "Replace" for atomic replacement (avoids stale range bug)
                                  expandedRange.insertHtml(htmlReplacement || "", "Replace");
                                  changesApplied++;
                                  expandedSuccessfully = true;
                                  break;
                                } else {
                                  console.log(`Expanded range length mismatch: ${expandedLength} vs ${originalLength}, trying next suffix match`);
                                }
                              } catch (expandError) {
                                console.warn("Could not expand to this suffix match:", expandError.message);
                              }
                            }

                            if (!expandedSuccessfully) {
                              // None of the suffix matches worked, fall back to prefix only
                              console.warn("No valid suffix match found, falling back to prefix-only replacement");
                              // Use insertHtml with "Replace" for atomic replacement
                              item.insertHtml(htmlReplacement || "", "Replace");
                              changesApplied++;
                            }
                          } else {
                            // Suffix not found, fall back to just the found range
                            console.warn("Could not find suffix for range expansion, applying to found range only");
                            // Use insertHtml with "Replace" for atomic replacement
                            item.insertHtml(htmlReplacement || "", "Replace");
                            changesApplied++;
                          }
                        } else {
                          // Suffix invalid length, fall back to just the found range
                          console.warn(`Suffix length invalid (${suffixText.length}), applying to found range only`);
                          // Use insertHtml with "Replace" for atomic replacement
                          item.insertHtml(htmlReplacement || "", "Replace");
                          changesApplied++;
                        }
                      } catch (expandError) {
                        console.warn("Range expansion failed, applying to found range only:", expandError.message);
                        // Use insertHtml with "Replace" for atomic replacement
                        item.insertHtml(htmlReplacement || "", "Replace");
                        changesApplied++;
                      }
                    } else {
                      // Standard case: exact match, delete then insert for clean redline
                      // Use insertHtml with "Replace" for atomic replacement
                      item.insertHtml(htmlReplacement || "", "Replace");
                      changesApplied++;
                    }
                  } catch (modifyError) {
                    console.error("Error applying modify_text:", modifyError);
                  }
                }
              } else {
                console.warn(`Could not find text "${searchText}" in Paragraph ${change.paragraphIndex}`);
              }
            } catch (searchError) {
              console.warn(`Search failed for modify_text "${searchText}" in Paragraph ${change.paragraphIndex}:`, searchError.message);

              // Fallback: Try with a shorter search string
              if (searchText.length > 30) {
                const shorterText = searchText.substring(0, 30);
                console.log(`Retrying modify_text with shorter search: "${shorterText}"`);
                try {
                  const retryResults = targetParagraph.search(shorterText, { matchCase: true });
                  retryResults.load("items");
                  await context.sync();

                  if (retryResults.items.length > 0) {
                    const replacementText = change.replacementText || "";
                    let htmlReplacement = markdownToWordHtmlInline(replacementText);
                    const trimmed = htmlReplacement.trim();
                    const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
                      trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
                      !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
                      !trimmed.includes('<table') && !trimmed.includes('<h');

                    if (hasSingleParagraph) {
                      htmlReplacement = trimmed.substring(3, trimmed.length - 4);
                    }
                    // Use insertHtml with "Replace" for atomic replacement
                    retryResults.items[0].insertHtml(htmlReplacement || "", "Replace");
                    changesApplied++;
                  }
                } catch (retryError) {
                  console.warn(`Retry search also failed for modify_text:`, retryError.message);
                }
              }
            }
          }

          // Ensure any queued operations for this change are executed here,
          // so errors are caught per-change instead of bubbling as one big GeneralException.
          await context.sync();
        } catch (changeError) {
          console.error("Error applying change:", changeError);
        }
      }

      // Final sync (should usually be a no-op now, but kept for safety)
      await context.sync();

      // Restore track changes only if we enabled it
      if (redlineEnabled && changeTrackingAvailable && originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
        try {
          context.document.changeTrackingMode = originalChangeTrackingMode;
          await context.sync();
        } catch (restoreError) {
          console.error("Could not restore track changes state:", restoreError);
        }
      }
    });

    console.log(`Total changes applied: ${changesApplied} `);

    if (changesApplied === 0) {
      return {
        message: "Applied 0 edits. The AI's suggestions could not be mapped to the document content.",
        showToUser: false  // Silent fallback - don't clutter the log
      };
    }

    return {
      message: `Successfully applied ${changesApplied} edits${redlineEnabled ? ' with redlines' : ' without redlines'}.`,
      showToUser: true
    };

  } catch (error) {
    console.error("Error in executeRedline:", error);
    return {
      message: `Error applying redlines: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}

// Helper for the Diff generation (specialized prompt)
async function callGeminiForDiffs(prompt) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const jsonSchema = {
    type: "ARRAY",
    items: {
      type: "OBJECT",
      properties: {
        "paragraphIndex": { "type": "INTEGER", "description": "The paragraph number (1-based)" },
        "endParagraphIndex": { "type": "INTEGER", "description": "Only for replace_range: the end paragraph number (inclusive)" },
        "operation": {
          "type": "STRING",
          "enum": ["edit_paragraph", "replace_paragraph", "modify_text", "replace_range"],
          "description": "The type of operation to perform"
        },
        "newContent": { "type": "STRING", "description": "For edit_paragraph only: the complete rewritten paragraph content" },
        "content": { "type": "STRING", "description": "For replace_paragraph and replace_range: the new content" },
        "originalText": { "type": "STRING", "description": "For modify_text only: the text to find (max 80 chars). Split larger edits into multiple operations." },
        "replacementText": { "type": "STRING", "description": "For modify_text only: the replacement text" }
      },
      required: ["paragraphIndex", "operation"]
    }
  };

  const systemInstruction = {
    parts: [
      {
        text: loadSystemMessage(),
      },
    ],
  };

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: systemInstruction,
    safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: API_LIMITS.MAX_OUTPUT_TOKENS,
      responseMimeType: "application/json",
      responseSchema: jsonSchema,
    },
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`API failed: ${err}`);
    }

    const result = await response.json();
    console.log("Gemini diff raw result:", JSON.stringify(result, null, 2));

    if (!result.candidates || !Array.isArray(result.candidates) || result.candidates.length === 0) {
      throw new Error("Gemini diff response contained no candidates.");
    }

    const candidate = result.candidates[0];

    if (!candidate.content || !candidate.content.parts || !Array.isArray(candidate.content.parts) || candidate.content.parts.length === 0) {
      console.error("Gemini diff candidate missing content.parts:", candidate);
      throw new Error("Gemini diff response was missing content.parts (possibly blocked by safety settings).");
    }

    const jsonText = candidate.content.parts[0].text;
    console.log("Gemini diff JSON text:", jsonText);
    return JSON.parse(jsonText);
  } catch (error) {
    console.error("Error getting diffs:", error);
    return null;
  }
}

/**
 * Agentic Tool: Inserts comments based on an instruction using Structural Anchoring.
 */
async function executeComment(instruction, fullDocumentText) {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of comments to be inserted, referencing the paragraph numbers.

Each item must be an object with:
- "paragraphIndex": The integer number of the paragraph to comment on (e.g., 1 for [P1]).
- "textToFind": The specific text snippet within the paragraph to attach the comment to. Must match EXACTLY. CRITICAL: Keep this VERY SHORT - maximum 50 characters or 5-8 words. Use a unique phrase that identifies the location.
- "commentContent": The text of the comment.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON ARRAY OF COMMENTS:`;

    const aiComments = await callGeminiForJSON(fullPrompt, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          "paragraphIndex": { "type": "INTEGER" },
          "textToFind": { "type": "STRING" },
          "commentContent": { "type": "STRING" }
        },
        required: ["paragraphIndex", "textToFind", "commentContent"]
      }
    });
    console.log("AI Suggested Comments:", aiComments);

    if (!aiComments || !Array.isArray(aiComments) || aiComments.length === 0) {
      return {
        message: "AI had no comments to suggest.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    let commentsApplied = 0;

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      for (const item of aiComments) {
        const pIndex = item.paragraphIndex - 1;
        if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

        const targetParagraph = paragraphs.items[pIndex];
        const count = await searchWithFallback(targetParagraph, item.textToFind, context, async (match) => {
          match.insertComment(item.commentContent);
        });
        commentsApplied += count;
      }
    });

    return createToolResult(commentsApplied, 'comments', "Inserted 0 comments. The AI's suggestions could not be mapped to the document content.");

  } catch (error) {
    console.error("Error in executeComment:", error);
    return {
      message: `Error inserting comments: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}

/**
 * Agentic Tool: Highlights text based on an instruction using Structural Anchoring.
 * @param {string} instruction - The instruction for what to highlight
 * @param {string} fullDocumentText - The document content with paragraph anchors
 * @param {string} highlightColor - The default highlight color (default: "Yellow")
 */
async function executeHighlight(instruction, fullDocumentText, highlightColor = "Yellow") {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  // Normalize color to proper case for Word API
  const normalizedColor = highlightColor.charAt(0).toUpperCase() + highlightColor.slice(1).toLowerCase();

  try {
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of highlights to be applied, referencing the paragraph numbers.

Each item must be an object with:
- "paragraphIndex": The integer number of the paragraph (e.g., 1 for [P1]).
- "textToFind": The specific text snippet within the paragraph to highlight. Must match EXACTLY. CRITICAL: Keep this VERY SHORT - maximum 50 characters or 5-8 words. Use a unique phrase that identifies the location.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON ARRAY OF HIGHLIGHTS:`;

    const aiHighlights = await callGeminiForJSON(fullPrompt, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          "paragraphIndex": { "type": "INTEGER" },
          "textToFind": { "type": "STRING" }
        },
        required: ["paragraphIndex", "textToFind"]
      }
    });
    console.log("AI Suggested Highlights:", aiHighlights);

    if (!aiHighlights || !Array.isArray(aiHighlights) || aiHighlights.length === 0) {
      return {
        message: "AI had no highlights to suggest.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    let highlightsApplied = 0;

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      for (const item of aiHighlights) {
        const pIndex = item.paragraphIndex - 1;
        if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

        const targetParagraph = paragraphs.items[pIndex];
        const count = await searchWithFallback(targetParagraph, item.textToFind, context, async (match) => {
          match.font.highlightColor = normalizedColor;
        });
        highlightsApplied += count;
      }
    });

    return createToolResult(highlightsApplied, 'highlights', "Highlighted 0 items. The AI's suggestions could not be mapped to the document content.");

  } catch (error) {
    console.error("Error in executeHighlight:", error);
    return {
      message: `Error highlighting text: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}

/**
 * Agentic Tool: Navigates to and selects a specific section of the document.
 */
async function executeNavigate(instruction, fullDocumentText) {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    const fullPrompt = `You are an expert document navigator. Review the document content (provided with [P#] anchors) based on the user's navigation instruction.
Determine the most relevant paragraph to navigate to and provide navigation details.

Return a JSON object with:
- "paragraphIndex": The integer number of the paragraph to navigate to (e.g., 1 for [P1]).
- "navigationDescription": A brief description of what was found and where the user was taken (e.g., "Navigated to paragraph 3: Introduction section", "Found the signature block at paragraph 15").

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON RESPONSE:`;

    const navigationResult = await callGeminiForJSON(fullPrompt, {
      type: "OBJECT",
      properties: {
        "paragraphIndex": { "type": "INTEGER" },
        "navigationDescription": { "type": "STRING" }
      },
      required: ["paragraphIndex"]
    });
    console.log("AI Navigation Result:", navigationResult);

    if (!navigationResult || !navigationResult.paragraphIndex) {
      return {
        message: "Could not determine where to navigate based on the instruction.",
        showToUser: false
      };
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const pIndex = navigationResult.paragraphIndex - 1;
      if (pIndex < 0 || pIndex >= paragraphs.items.length) {
        throw new Error(`Invalid paragraph index: ${navigationResult.paragraphIndex}`);
      }

      const targetParagraph = paragraphs.items[pIndex];

      // Select the paragraph to navigate to it
      targetParagraph.select();
      await context.sync();
    });

    const description = navigationResult.navigationDescription || `Navigated to paragraph ${navigationResult.paragraphIndex}`;

    return {
      message: description,
      showToUser: true
    };

  } catch (error) {
    console.error("Error in executeNavigate:", error);
    return {
      message: `Error navigating: ${error.message}`,
      showToUser: false
    };
  }
}

// ==================== TOOL EXECUTION HELPERS ====================

/**
 * Validates that prerequisites for tool execution are met (API key exists).
 * @returns {Object} Object with either { apiKey } or { error }
 */
function validateToolPrerequisites() {
  const apiKey = loadApiKey();
  if (!apiKey) {
    return { error: "Error: Please set your Gemini API key in the Settings." };
  }
  return { apiKey };
}

/**
 * Creates a standardized tool execution result object.
 * @param {number} count - Number of items successfully processed
 * @param {string} itemType - Type of item (e.g., "comments", "highlights")
 * @param {string} zeroMessage - Optional custom message for zero count
 * @returns {Object} Result object with { message, showToUser }
 */
function createToolResult(count, itemType, zeroMessage) {
  if (count === 0) {
    return {
      message: zeroMessage || `Applied 0 ${itemType}. The AI's suggestions could not be mapped to the document content.`,
      showToUser: false  // Silent fallback
    };
  }

  const actionVerb = itemType === 'comments' ? 'inserted' : itemType === 'highlights' ? 'highlighted' : 'applied';
  return {
    message: `Successfully ${actionVerb} ${count} ${itemType}.`,
    showToUser: true
  };
}

/**
 * Searches for text within a paragraph with automatic fallback to shorter text on failure.
 * @param {Word.Paragraph} targetParagraph - The paragraph to search within
 * @param {string} searchText - The text to search for
 * @param {Word.RequestContext} context - Word context for sync operations
 * @param {Function} onSuccess - Callback function to execute on each match (receives match object)
 * @returns {Promise<number>} Number of successful operations
 */
async function searchWithFallback(targetParagraph, searchText, context, onSuccess) {
  let operationsCount = 0;

  // Validate and truncate search text
  if (!searchText || searchText.trim().length === 0) {
    return 0;
  }

  if (searchText.length > SEARCH_LIMITS.MAX_LENGTH) {
    searchText = searchText.substring(0, SEARCH_LIMITS.MAX_LENGTH);
  }

  try {
    const searchResults = targetParagraph.search(searchText, { matchCase: false });
    searchResults.load("items");
    await context.sync();

    if (searchResults.items.length > 0) {
      for (const match of searchResults.items) {
        await onSuccess(match);
        operationsCount++;
      }
      return operationsCount;
    }
  } catch (searchError) {
    console.warn(`Search failed for "${searchText}":`, searchError.message);

    // Fallback: Try with shorter text
    if (searchText.length > SEARCH_LIMITS.RETRY_LENGTH) {
      const shorterText = searchText.substring(0, SEARCH_LIMITS.RETRY_LENGTH);
      console.log(`Retrying with shorter search: "${shorterText}"`);

      try {
        const retryResults = targetParagraph.search(shorterText, { matchCase: false });
        retryResults.load("items");
        await context.sync();

        if (retryResults.items.length > 0) {
          await onSuccess(retryResults.items[0]);  // Only use first match for fallback
          return 1;
        }
      } catch (retryError) {
        console.warn(`Retry search also failed:`, retryError.message);
      }
    }
  }

  return 0;
}

// Generic helper for JSON responses
async function callGeminiForJSON(prompt, schema) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const systemInstruction = {
    parts: [{ text: loadSystemMessage() }]
  };

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: systemInstruction,
    safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 48000,
      responseMimeType: "application/json",
      responseSchema: schema,
    },
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`API failed: ${err}`);
    }

    const result = await response.json();
    if (!result.candidates || result.candidates.length === 0) throw new Error("No candidates");
    const candidate = result.candidates[0];
    if (!candidate.content || !candidate.content.parts) throw new Error("No content");

    const jsonText = candidate.content.parts[0].text;
    return JSON.parse(jsonText);
  } catch (error) {
    console.error("Error calling Gemini for JSON:", error);
    return null;
  }
}


async function executeResearch(query) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const tools = [{ google_search: {} }];

  const payload = {
    contents: [{ parts: [{ text: query }] }],
    tools: tools,
    safetySettings: [
      { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ]
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Research API failed: ${err}`);
    }

    const result = await response.json();
    if (!result.candidates || result.candidates.length === 0) return "No results found.";

    const candidate = result.candidates[0];
    if (!candidate.content || !candidate.content.parts) return "No content returned.";

    return candidate.content.parts[0].text;
  } catch (error) {
    console.error("Error in executeResearch:", error);
    return `Error performing research: ${error.message}`;
  }
}

/**
 * Maintains a rolling window of chat history while preserving function call/response pairs
 */
function maintainHistoryWindow(history, maxMessages) {
  if (history.length <= maxMessages) {
    return history;
  }

  // Start from the end and work backwards, keeping complete pairs
  let newHistory = [];
  let i = history.length - 1;

  while (i >= 0 && newHistory.length < maxMessages) {
    const msg = history[i];

    // If this is a function response, we must include its preceding function call
    const isFunctionResponse = msg.role === "user" && msg.parts && msg.parts.some(p => p.functionResponse);

    if (isFunctionResponse && i > 0) {
      const prevMsg = history[i - 1];
      const hasFunctionCall = prevMsg.role === "model" && prevMsg.parts && prevMsg.parts.some(p => p.functionCall);

      if (hasFunctionCall) {
        // Add both the function call and response together
        newHistory.unshift(msg);
        newHistory.unshift(prevMsg);
        i -= 2;
        continue;
      }
    }

    // If this is a function call, check if its response is already included
    const hasFunctionCall = msg.role === "model" && msg.parts && msg.parts.some(p => p.functionCall);

    if (hasFunctionCall && i < history.length - 1) {
      const nextMsg = history[i + 1];
      const hasResponse = nextMsg.role === "user" && nextMsg.parts && nextMsg.parts.some(p => p.functionResponse);

      if (hasResponse && !newHistory.includes(nextMsg)) {
        // Skip this function call since its response isn't in our window
        i--;
        continue;
      }
    }

    newHistory.unshift(msg);
    i--;
  }

  // Final validation: remove any orphaned function calls or responses at the boundaries
  return validateHistoryPairs(newHistory);
}

/**
 * Validates that function calls and responses are properly paired.
 *
 * In addition to enforcing adjacency, this also enforces that:
 * - If a model turn contains N function calls for a given tool name,
 *   the very next user turn must contain N function responses for that
 *   same tool name.
 * - There are no extra function responses for tools that were not called.
 *
 * This mirrors the behaviour described in the Gemini tooling docs and the
 * forum discussion you referenced, and strips out any legacy turns where
 * the counts didn't match (e.g. old code that only returned a single
 * functionResponse for multiple functionCalls).
 */
function validateHistoryPairs(history) {
  const validated = [];

  for (let i = 0; i < history.length; i++) {
    const msg = history[i];
    const parts = msg.parts || [];

    const hasFunctionCall =
      msg.role === "model" && parts.some((p) => p.functionCall);
    const isFunctionResponse =
      msg.role === "user" && parts.some((p) => p.functionResponse);

    // If validated is empty and this is a model turn, skip it
    // (Conversations must start with a user turn)
    if (validated.length === 0 && msg.role === "model") {
      console.warn(
        `Skipping model turn at index ${i} - cannot start history with a model turn.`
      );
      continue;
    }

    // --- Model turn with one or more function calls ---
    if (hasFunctionCall) {
      // CRITICAL: A model turn with function calls can ONLY come after a user turn
      // (either a regular text turn or a function response turn).
      // If the last message in validated is a model turn, this would cause:
      // "function call turn comes immediately after a user turn or after a function response turn" error
      const lastValidated = validated.length > 0 ? validated[validated.length - 1] : null;
      if (lastValidated && lastValidated.role === "model") {
        console.warn(
          `Removing function call at index ${i} - cannot follow another model turn. ` +
          `Last validated turn was role: ${lastValidated.role}. ` +
          `This would cause: "function call turn comes immediately after a user turn or after a function response turn" error.`
        );
        continue;
      }

      const nextMsg = i < history.length - 1 ? history[i + 1] : null;
      if (!nextMsg) {
        console.warn(
          `Removing orphaned function call at index ${i} (no following message).`
        );
        continue;
      }

      const nextParts = nextMsg.parts || [];
      const responseParts =
        nextMsg.role === "user"
          ? nextParts.filter((p) => p.functionResponse)
          : [];

      if (responseParts.length === 0) {
        console.warn(
          `Removing orphaned function call at index ${i} (no function responses in next turn).`
        );
        continue;
      }

      // Count how many times each tool was called in this turn
      const callCounts = {};
      parts.forEach((p) => {
        if (p.functionCall && p.functionCall.name) {
          const name = p.functionCall.name;
          callCounts[name] = (callCounts[name] || 0) + 1;
        }
      });

      // Count how many function responses we have per tool name
      const responseCounts = {};
      responseParts.forEach((p) => {
        const fr = p.functionResponse;
        const name = fr && fr.name;
        if (name) {
          responseCounts[name] = (responseCounts[name] || 0) + 1;
        }
      });

      let mismatch = false;

      // Every called tool must have exactly as many responses
      Object.keys(callCounts).forEach((name) => {
        if (callCounts[name] !== (responseCounts[name] || 0)) {
          mismatch = true;
        }
      });

      // And there must not be responses for tools that were never called
      Object.keys(responseCounts).forEach((name) => {
        if (!callCounts[name]) {
          mismatch = true;
        }
      });

      if (mismatch) {
        console.warn(
          `Removing mismatched function call/response pair at index ${i}. ` +
          `Calls: ${JSON.stringify(callCounts)}, ` +
          `Responses: ${JSON.stringify(responseCounts)}`
        );
        // Drop this model turn, and if the next turn is its response, drop that too.
        if (nextMsg.role === "user" && responseParts.length > 0) {
          i++; // Skip the mismatched response as well
        }
        continue;
      }

      // Pair looks good: keep both the model functionCall turn and the user functionResponse turn
      validated.push(msg);
      validated.push(nextMsg);
      i++; // Skip the response since we already added it
      continue;
    }

    // --- User turn with function responses but no preceding call in validated history ---
    if (isFunctionResponse) {
      const prevMsg = validated.length > 0 ? validated[validated.length - 1] : null;
      const prevParts = prevMsg && prevMsg.parts ? prevMsg.parts : [];
      const prevHasCall =
        prevMsg &&
        prevMsg.role === "model" &&
        prevParts.some((p) => p.functionCall);

      if (!prevHasCall) {
        console.warn(
          `Removing orphaned function response at index ${i} (no preceding function call in validated history).`
        );
        continue;
      }
    }

    // Regular message (no function call/response semantics to enforce)
    validated.push(msg);
  }

  return validated;
}

function sanitizeHistory() {
  if (chatHistory.length === 0) return;

  // Use the validation function to clean up the history
  chatHistory = validateHistoryPairs(chatHistory);
}

/**
 * Tier 2 Recovery: Remove ALL function call/response pairs from history
 * Keeps only regular text messages
 */
function removeAllFunctionPairs(history) {
  return history.filter(msg => {
    const parts = msg.parts || [];
    const hasFunctionCall = parts.some(p => p.functionCall);
    const hasFunctionResponse = parts.some(p => p.functionResponse);
    return !hasFunctionCall && !hasFunctionResponse;
  });
}

/**
 * Tier 3 Recovery: Create fresh start with minimal context
 * Returns new history with just the original user message
 */
function createFreshStartWithContext(originalUserMessage) {
  return [{
    role: "user",
    parts: [{ text: originalUserMessage }]
  }];
}

/**
 * Generate graceful degradation message based on executed tools
 */

/**
 * Extracts "accepted changes" view of text from paragraph OOXML.
 * - Includes text in <w:ins> tags (insertions that would be accepted)
 * - Excludes text in <w:del> tags (deletions that would be removed)
 * - Includes all regular <w:t> text
 * This ensures the AI sees a consistent text representation even when
 * documents have existing tracked changes.
 */

/**
 * Detects if content has block elements (lists, tables, headings)
 * that require HTML insertion instead of word-level diffs
 */
function hasBlockElements(content) {
  if (!content) return false;

  // Check for markdown block elements with improved patterns

  // Detect unordered lists: lines starting with -, *, or + followed by space
  const hasUnorderedList = /^[\s]*[-*+]\s+/m.test(content);

  // Detect ordered lists: lines starting with number(s) followed by period and space
  // Examples: "1. item", "10. item", "  2. item"
  const hasOrderedList = /^[\s]*\d+\.\s+/m.test(content);

  // Detect alphabetical lists: (a), (b), (c) style
  const hasAlphaList = /^[\s]*\([a-z]\)\s+/m.test(content);

  // Detect tables: markdown table syntax with pipes
  const hasTable = /\|.*\|.*\n/.test(content);

  // Detect headings: lines starting with # symbols
  const hasHeading = /^#{1,6}\s/m.test(content);

  // Detect paragraph breaks (multiple consecutive newlines)
  const hasMultipleLineBreaks = content.includes('\n\n');

  const result = hasUnorderedList || hasOrderedList || hasAlphaList || hasTable || hasHeading || hasMultipleLineBreaks;

  // Debug logging to help diagnose issues
  if (result) {
    console.log('Block elements detected:', {
      hasUnorderedList,
      hasOrderedList,
      hasAlphaList,
      hasTable,
      hasHeading,
      hasMultipleLineBreaks,
      contentPreview: content.substring(0, 100)
    });
  }

  return result;
}

/**
 * Checks if text contains inline markdown formatting (bold, italic, code, etc.)
 * Returns true if formatting patterns are detected
 */
function hasInlineMarkdownFormatting(text) {
  if (!text) return false;
  // Check for common inline markdown patterns:
  // **bold**, *italic*, __bold__, _italic_, `code`, ~~strikethrough~~, ++underline++
  // Also check for **...** pattern specifically
  return /(\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_|`.+?`|~~.+?~~|\+\+.+?\+\+)/.test(text);
}

/**
 * Wrapper for preprocessMarkdown that handles empty paragraph cases.
 * Returns clean text and format hints for formatting application.
 * 
 * @param {string} content - Content with markdown formatting
 * @returns {{ cleanText: string, formatHints: Array }}
 */
async function preprocessMarkdownForParagraph(content) {
  try {
    return preprocessMarkdown(content);
  } catch (e) {
    console.error('preprocessMarkdown failed:', e);
    return { cleanText: content, formatHints: [] };
  }
}

/**
 * Applies format hints to specific text ranges in a paragraph using Word's font API.
 * This avoids HTML/OOXML insertion issues in table cells.
 * 
 * @param {Word.Paragraph} paragraph - Target paragraph
 * @param {string} text - The text content
 * @param {Array} formatHints - Array of format hints with start/end/format
 * @param {Word.RequestContext} context - Word context
 */
async function applyFormatHintsToRanges(paragraph, text, formatHints, context) {
  // Load paragraph as range
  const paragraphRange = paragraph.getRange();
  paragraphRange.load('text');
  await context.sync();

  // Get the paragraph text to verify positions
  const paragraphText = paragraphRange.text;

  for (const hint of formatHints) {
    try {
      // Calculate the text to search for based on the hint offsets
      const hintText = text.substring(hint.start, hint.end);
      if (!hintText || hintText.trim().length === 0) continue;

      // Search for the text within the paragraph
      const searchResults = paragraphRange.search(hintText, { matchCase: true, matchWholeWord: false });
      searchResults.load('items');
      await context.sync();

      if (searchResults.items.length > 0) {
        // Apply formatting to the first match
        const targetRange = searchResults.items[0];

        if (hint.format.bold) {
          targetRange.font.bold = true;
        }
        if (hint.format.italic) {
          targetRange.font.italic = true;
        }
        if (hint.format.underline) {
          targetRange.font.underline = Word.UnderlineType.single;
        }
        if (hint.format.strikethrough) {
          targetRange.font.strikeThrough = true;
        }

        await context.sync();
      }
    } catch (formatError) {
      console.warn(`Could not apply formatting to hint at ${hint.start}-${hint.end}:`, formatError);
    }
  }
}

/**
 * Inserts text at a range, using HTML insertion if markdown formatting is detected
 * This ensures **bold**, *italic*, etc. are properly rendered instead of literal
 */

/**
 * Converts text into word tokens represented as unique characters.
 * This allows DMP to diff at word-level instead of character-level.
 * 
 * Words are split on whitespace, but whitespace is preserved as separate tokens.
 * 
 * @param {string} text1 - First text to tokenize
 * @param {string} text2 - Second text to tokenize
 * @returns {{chars1: string, chars2: string, wordArray: string[]}}
 */

/**
 * Converts character-encoded diffs back to actual word diffs.
 * 
 * @param {Array} diffs - DMP diff array with character codes
 * @param {string[]} wordArray - Array mapping char codes to words
 * @returns {Array} - DMP-style diff array with actual words
 */

/**
 * Applies word-level diffs to a paragraph using DMP with contextual search
 * Uses surrounding context to ensure unique matches and proper tracked changes
 * 
 * NEW: Uses word-level diffing instead of character-level to produce cleaner changes.
 * NEW: First checks if changes are purely formatting-related (adding bold, italic, etc.)
 * and applies them surgically using Word's native font API.
 */

/**
 * Converts DMP diff output to search/replace pairs
 * Returns array of changes in reverse order (for right-to-left application)
 */

// ==================== NATIVE WORD API FUNCTIONS ====================

/**
 * Parses markdown list content into structured data
 * Supports numbered lists (1. item) and bullet lists (- item, * item)
 */
function parseMarkdownList(content) {
  if (!content) return null;

  const lines = content.trim().split('\n');
  const items = [];

  for (const line of lines) {
    if (!line.trim()) continue;

    // Unified marker regex matching NumberingService and Pipeline
    const markerRegex = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)(.*)$/;
    const match = line.match(markerRegex);

    if (match) {
      const indent = match[1];
      const marker = match[2].trim();
      const text = match[3];
      const level = Math.floor(indent.length / 2); // 2 spaces per level

      // Determine type based on marker
      const isBullet = /^[-*•]$/.test(marker);
      items.push({
        type: isBullet ? 'bullet' : 'numbered',
        level,
        text: text.trim(),
        marker: marker
      });
      continue;
    }

    // If line doesn't match list pattern, still include as text
    items.push({ type: 'text', level: 0, text: line.trim() });
  }

  if (items.length === 0) return null;

  // Determine primary type (numbered or bullet)
  const hasNumbered = items.some(i => i.type === 'numbered');
  const hasBullet = items.some(i => i.type === 'bullet');
  console.log(`[parseMarkdownList] Processing ${lines.length} lines. hasNumbered=${hasNumbered}, hasBullet=${hasBullet}`);

  return {
    type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
    items: items
  };
}


/**
 * Applies a numbered or bullet list using Word's native list API
 * This is more reliable than HTML insertion
 */
async function applyNativeList(targetParagraph, listData, context) {
  if (!listData || !listData.items || listData.items.length === 0) {
    console.warn('No list items to apply');
    return;
  }

  console.log(`Applying native ${listData.type} list with ${listData.items.length} items`);

  // Determine the built-in style to use
  const listStyle = listData.type === 'numbered'
    ? Word.BuiltInStyleName.listNumber
    : Word.BuiltInStyleName.listBullet;

  // Clear the target paragraph first
  targetParagraph.clear();

  // Apply first item to target paragraph
  const firstItem = listData.items[0];
  targetParagraph.insertText(firstItem.text, Word.InsertLocation.end);
  targetParagraph.styleBuiltIn = listStyle;

  // Set the list level for first item
  targetParagraph.load('listItemOrNullObject');
  await context.sync();

  if (!targetParagraph.listItemOrNullObject.isNullObject) {
    targetParagraph.listItemOrNullObject.level = firstItem.level || 0;
  }

  // Insert remaining items
  let previousPara = targetParagraph;
  for (let i = 1; i < listData.items.length; i++) {
    const item = listData.items[i];

    // Insert new paragraph after previous
    const newPara = previousPara.insertParagraph(item.text, Word.InsertLocation.after);
    newPara.styleBuiltIn = listStyle;

    // Set list level
    newPara.load('listItemOrNullObject');
    await context.sync();

    if (!newPara.listItemOrNullObject.isNullObject) {
      newPara.listItemOrNullObject.level = item.level || 0;
    }

    previousPara = newPara;
  }

  await context.sync();
  console.log(`Successfully applied native list`);
}



/**
 * Normalizes content by converting literal escape sequences to actual characters.
 * This is necessary because AI responses sometimes return "\\n" as a two-character
 * string instead of actual newlines, which breaks markdown parsing.
 */
function normalizeContentEscapes(content) {
  if (!content || typeof content !== 'string') return content;

  // Convert literal \n (two characters) to actual newline
  // Also handle other common escapes
  return content
    .replace(/\\n/g, '\n')      // Literal \n -> newline
    .replace(/\\t/g, '\t')      // Literal \t -> tab
    .replace(/\\r/g, '\r');     // Literal \r -> carriage return
}

/**
 * Routes a change operation to the appropriate method
 * Uses native Word APIs for lists/tables, DMP for text edits
 */
async function routeChangeOperation(change, targetParagraph, context) {
  const originalText = targetParagraph.text;
  let newContent = change.newContent || change.content || "";

  // Normalize content: Convert literal escape sequences to actual characters
  // This handles cases where the AI returns "\\n" as a two-character string instead of actual newlines
  newContent = normalizeContentEscapes(newContent);

  // 1. Empty original text - try native APIs first
  if (!originalText || originalText.trim().length === 0) {
    console.log("Empty paragraph detected");

    // Try to parse as list
    const listData = parseMarkdownList(newContent);
    if (listData && listData.type !== 'text') {
      console.log(`Using native ${listData.type} list API`);
      await applyNativeList(targetParagraph, listData, context);
      return;
    }

    // Try to parse as table - use OOXML Hybrid Mode even for empty paragraphs
    const matchedTable = newContent.includes('|');
    if (matchedTable) {
      const tableData = parseTable(newContent);
      if (tableData.rows.length > 0 || tableData.headers.length > 0) {
        console.log("Detected table in empty paragraph, using OOXML Hybrid Mode");
        // Fall through to OOXML Engine (Stage 4) which handles empty original text correctly
      }
    } else {
      // ...Existing formatting/HTML check...
    }

    // Check if this is simple text with possible formatting
    // For cells inside tables, prefer OOXML path to avoid nested table issues
    const hasFormatting = hasInlineMarkdownFormatting(newContent);
    if (hasFormatting) {
      console.log("Empty paragraph with formatting - using insertText for simplicity");
      // For empty paragraphs with simple formatting, just insert the text directly
      // The formatting will be applied as markdown symbols which is better than nested tables
      // Use insertText with markdown stripped, then apply formatting separately
      const { cleanText, formatHints } = await preprocessMarkdownForParagraph(newContent);
      targetParagraph.insertText(cleanText, "Replace");
      await context.sync();

      // If there are format hints, apply them using Word's font API
      if (formatHints.length > 0) {
        try {
          await applyFormatHintsToRanges(targetParagraph, cleanText, formatHints, context);
        } catch (formatError) {
          console.warn("Could not apply formatting:", formatError);
        }
      }
      return;
    }

    // Fall back to HTML for other content (no formatting, no tables)
    console.log("Using HTML insertion for empty paragraph");
    const htmlContent = markdownToWordHtml(newContent);
    targetParagraph.insertHtml(htmlContent, "Replace");
    return;
  }

  // 2. Check for structured content types

  // Try to parse as numbered/bullet list
  const listData = parseMarkdownList(newContent);
  if (listData && listData.type !== 'text') {
    console.log(`Detected ${listData.type} list, using native API`);
    await applyNativeList(targetParagraph, listData, context);
    return;
  }

  // NOTE: Table detection removed here. Let applyRedlineToOxml handle tables 
  // via OOXML Hybrid Mode, which handles both existing tables and text-to-table.

  // 3. Check for block elements (headings, mixed content, etc.)
  if (hasBlockElements(newContent)) {
    console.log("Block elements detected, using HTML replacement");
    const htmlContent = markdownToWordHtml(newContent);
    targetParagraph.insertHtml(htmlContent, "Replace");
    return;
  }

  // 4. Use OOXML Engine V5.1 (Hybrid Mode) for proper track changes
  // This modifies the DOM in-place, embedding w:ins/w:del directly in the structure
  console.log("Attempting OOXML Hybrid Mode for text edit");
  const redlineEnabled = loadRedlineSetting();

  // Get original text and paragraph OOXML
  // Get original text (preserve whitespace for exact diffing) and paragraph OOXML
  const paragraphOriginalText = targetParagraph.text;
  const paragraphOoxmlResult = targetParagraph.getOoxml();
  await context.sync();

  console.log("[OxmlEngine] Original text:", paragraphOriginalText.substring(0, 100));
  console.log("[OxmlEngine] Paragraph OOXML length:", paragraphOoxmlResult.value.length);

  // Apply redlines using hybrid engine (DOM manipulation approach)
  const redlineAuthor = loadRedlineAuthor();
  const result = await applyRedlineToOxml(
    paragraphOoxmlResult.value,
    paragraphOriginalText,
    newContent,
    {
      author: redlineEnabled ? redlineAuthor : undefined,
      generateRedlines: redlineEnabled
    }
  );

  if (!result.hasChanges) {
    console.log("[OxmlEngine] No changes detected by engine");
    return;
  }

  // Handle native API formatting for table cells
  if (result.useNativeApi && result.formatHints) {
    console.log("[OxmlEngine] Using native Font API for table cell formatting");
    await applyFormatHintsToRanges(targetParagraph, result.originalText, result.formatHints, context);
    console.log("✅ Native API formatting successful");
    return;
  }

  console.log("[OxmlEngine] Generated OOXML with track changes, length:", result.oxml.length);

  try {
    // The hybrid engine embeds w:ins/w:del directly in the DOM structure
    // We still need to disable Word's track changes to prevent double-tracking
    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    const originalMode = doc.changeTrackingMode;
    console.log(`[OxmlEngine] Current track changes mode: ${originalMode}, redlineEnabled: ${redlineEnabled}`);

    // Disable track changes temporarily - our w:ins/w:del ARE the track changes
    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
      console.log("[OxmlEngine] Temporarily disabling Word track changes for OOXML insertion");
      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    }

    try {
      // Insert the modified OOXML - since it's a paragraph-level replacement,
      // and our DOM already contains the track change markers embedded in the structure,
      // Word should render them as track changes
      targetParagraph.insertOoxml(result.oxml, 'Replace');
      await context.sync();
      console.log("✅ OOXML Hybrid Mode reconciliation successful");
    } finally {
      // Restore track changes mode
      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
        console.log(`[OxmlEngine] Restoring track changes mode to: ${originalMode}`);
        doc.changeTrackingMode = originalMode;
        await context.sync();
      }
    }
  } catch (insertError) {
    console.error("❌ OOXML insertion failed:", insertError.message);
    // Fallback to simple text replacement
    console.log("Falling back to simple text replacement");
    targetParagraph.insertText(newContent, "Replace");
    await context.sync();
  }
}

/**
 * Fallback function for modify_text operations
 * Used when DMP approach fails
 */

/**
 * Execute insert_list_item tool - surgically insert a single list item after a specific paragraph
 * @param {number} afterParagraphIndex - 1-based paragraph index to insert after
 * @param {string} text - The text content (without numbering)
 * @param {number} indentLevel - Relative indent: 0=same, 1=deeper, -1=shallower
 */
async function executeInsertListItem(afterParagraphIndex, text, indentLevel = 0) {
  console.log(`[executeInsertListItem] Insert after P${afterParagraphIndex}: "${text.substring(0, 50)}..." (indent: ${indentLevel})`);

  try {
    await Word.run(async (context) => {
      // Enable track changes if redline setting is enabled
      const redlineEnabled = loadRedlineSetting();
      let originalChangeTrackingMode = null;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          originalChangeTrackingMode = doc.changeTrackingMode;
          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
            console.log("[executeInsertListItem] Track changes enabled");
          }
        } catch (trackError) {
          console.warn("[executeInsertListItem] Could not enable track changes:", trackError);
        }
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const paraIdx = afterParagraphIndex - 1; // Convert to 0-based
      if (paraIdx < 0 || paraIdx >= paragraphs.items.length) {
        throw new Error(`Paragraph index ${afterParagraphIndex} out of range (1-${paragraphs.items.length})`);
      }

      const adjacentPara = paragraphs.items[paraIdx];

      // Read the adjacent paragraph's OOXML to get its numId and ilvl
      const adjacentOoxml = adjacentPara.getOoxml();
      await context.sync();

      const numIdMatch = adjacentOoxml.value.match(/w:numId w:val="(\d+)"/);
      const ilvlMatch = adjacentOoxml.value.match(/w:ilvl w:val="(\d+)"/);

      // Debug: Log the numbering definition info if available
      const lvlTextMatch = adjacentOoxml.value.match(/w:lvlText w:val="([^"]*)"/);
      if (lvlTextMatch) {
        console.log(`[executeInsertListItem] Adjacent lvlText format: "${lvlTextMatch[1]}"`);
      }

      // Log a snippet of the OOXML for debugging numbering structure
      const numPrSection = adjacentOoxml.value.match(/<w:numPr[\s\S]*?<\/w:numPr>/);
      if (numPrSection) {
        console.log(`[executeInsertListItem] Adjacent numPr: ${numPrSection[0]}`);
      }

      if (!numIdMatch) {
        // Adjacent paragraph is not a list item - just insert plain paragraph
        console.log("[executeInsertListItem] Adjacent paragraph is not a list item, inserting plain paragraph");
        adjacentPara.insertParagraph(text, "After");
        await context.sync();

        // Restore tracking mode
        if (redlineEnabled && originalChangeTrackingMode !== null &&
          originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
          context.document.changeTrackingMode = originalChangeTrackingMode;
          await context.sync();
        }
        return;
      }

      const numId = numIdMatch[1];
      const baseIlvl = ilvlMatch ? parseInt(ilvlMatch[1], 10) : 0;
      const newIlvl = Math.max(0, Math.min(8, baseIlvl + indentLevel)); // Clamp to 0-8

      console.log(`[executeInsertListItem] Adjacent numId=${numId}, ilvl=${baseIlvl}, newIlvl=${newIlvl}`);

      // Extract run properties (rPr) from adjacent paragraph to preserve font styling
      // Look for the first w:rPr in the paragraph's OOXML
      let rPrBlock = '';
      const rPrMatch = adjacentOoxml.value.match(/<w:rPr[^>]*>([\s\S]*?)<\/w:rPr>/);
      if (rPrMatch) {
        rPrBlock = rPrMatch[0];
        console.log(`[executeInsertListItem] Extracted rPr from adjacent paragraph`);
      } else {
        // No rPr found, try to at least get the font from pPr
        const fontMatch = adjacentOoxml.value.match(/<w:rFonts[^>]*\/>/);
        if (fontMatch) {
          rPrBlock = `<w:rPr>${fontMatch[0]}</w:rPr>`;
          console.log(`[executeInsertListItem] Extracted rFonts from adjacent paragraph`);
        }
      }

      // Build OOXML for the new list item
      const escapedText = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

      const oxmlPara = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
          <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData>
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
              </Relationships>
            </pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
            <pkg:xmlData>
              <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val="ListParagraph"/>
                      <w:numPr>
                        <w:ilvl w:val="${newIlvl}"/>
                        <w:numId w:val="${numId}"/>
                      </w:numPr>
                    </w:pPr>
                    <w:r>
                      ${rPrBlock}
                      <w:t xml:space="preserve">${escapedText}</w:t>
                    </w:r>
                  </w:p>
                </w:body>
              </w:document>
            </pkg:xmlData>
          </pkg:part>
        </pkg:package>`;

      // Insert the paragraph with text, then apply list formatting
      const insertedPara = adjacentPara.insertParagraph(text, "After");
      await context.sync();

      // Try to apply the same list formatting using Word's list API
      // The insertedPara should inherit some formatting, but we need to set the list explicitly
      try {
        // Load the inserted paragraph to access its list properties
        insertedPara.load("listItem");
        await context.sync();

        // If it has a listItem, we can adjust its level
        if (insertedPara.listItem && !insertedPara.listItem.isNullObject) {
          // The list item exists - try to adjust level
          console.log(`[executeInsertListItem] Inserted paragraph has listItem, adjusting level to ${newIlvl}`);
          insertedPara.listItem.level = newIlvl;
          await context.sync();
        } else {
          // No listItem - need to add it to a list
          // Use the same numId as adjacent paragraph via OOXML
          console.log(`[executeInsertListItem] No listItem found, applying list via OOXML`);

          const paraRange = insertedPara.getRange("Whole");
          paraRange.insertOoxml(oxmlPara, "Replace");
          await context.sync();
        }
      } catch (listError) {
        console.warn(`[executeInsertListItem] Could not apply list format via API: ${listError.message}`);
        // Fallback: try OOXML replacement
        try {
          const paraRange = insertedPara.getRange("Whole");
          paraRange.insertOoxml(oxmlPara, "Replace");
          await context.sync();
        } catch (oxmlError) {
          console.warn(`[executeInsertListItem] OOXML fallback also failed: ${oxmlError.message}`);
        }
      }

      console.log(`[executeInsertListItem] Successfully inserted list item (numId=${numId}, ilvl=${newIlvl})`);


      // Restore original tracking mode if we changed it
      if (redlineEnabled && originalChangeTrackingMode !== null &&
        originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
        context.document.changeTrackingMode = originalChangeTrackingMode;
        await context.sync();
      }
    });

    return {
      success: true,
      message: `Successfully inserted list item after P${afterParagraphIndex}`
    };
  } catch (error) {
    console.error("[executeInsertListItem] Error:", error);
    return {
      success: false,
      message: `Failed to insert list item: ${error.message}`
    };
  }
}

/**
 * Execute edit_list tool - replaces a range of paragraphs with a proper list
 * Uses HTML insertion for reliable list formatting
 * @param {number} startIndex - 1-based paragraph index of first paragraph
 * @param {number} endIndex - 1-based paragraph index of last paragraph
 * @param {string[]} newItems - Array of new list item texts
 * @param {string} listType - "bullet" or "numbered"
 * @param {string} numberingStyle - For numbered lists: "decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"
 */
async function executeEditList(startIndex, endIndex, newItems, listType, numberingStyle) {
  if (!newItems || newItems.length === 0) {
    return { success: false, message: "No list items provided." };
  }

  console.log(`executeEditList: Converting P${startIndex}-P${endIndex} to ${listType} list with ${newItems.length} items`);
  console.log(`[executeEditList] Raw newItems array:`);
  newItems.forEach((item, idx) => {
    console.log(`  [${idx}]: "${item.substring(0, 60)}${item.length > 60 ? '...' : ''}"`);
  });

  try {
    await Word.run(async (context) => {
      // Detect document font for consistent HTML insertion
      await detectDocumentFont();

      // Enable track changes if redline setting is enabled (same pattern as executeRedline)
      const redlineEnabled = loadRedlineSetting();
      let originalChangeTrackingMode = null;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          originalChangeTrackingMode = doc.changeTrackingMode;

          // Force enable track changes
          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
            console.log("Track changes enabled for list edit");
          }
        } catch (trackError) {
          console.warn("Could not enable track changes:", trackError);
        }
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let startIdx = startIndex - 1; // Convert to 0-based
      let endIdx = endIndex - 1;

      // Handle out-of-range paragraph indices gracefully
      // The AI may reference paragraphs that don't exist (e.g., after list expansion)
      const paragraphCount = paragraphs.items.length;

      if (paragraphCount === 0) {
        throw new Error("Document has no paragraphs");
      }

      // If start is beyond document, append at end
      if (startIdx >= paragraphCount) {
        console.log(`Start index ${startIndex} exceeds document (${paragraphCount} paragraphs), treating as append`);
        startIdx = paragraphCount - 1;
        endIdx = paragraphCount - 1;
      }

      // Clamp start to valid range
      if (startIdx < 0) {
        startIdx = 0;
      }

      // Clamp end to valid range
      if (endIdx >= paragraphCount) {
        console.log(`End index ${endIndex} exceeds document (${paragraphCount} paragraphs), clamping to ${paragraphCount}`);
        endIdx = paragraphCount - 1;
      }

      // Ensure start <= end
      if (startIdx > endIdx) {
        startIdx = endIdx;
      }

      console.log(`Adjusted range: P${startIdx + 1} to P${endIdx + 1} (original: ${startIndex} to ${endIndex})`);

      // Get the range covering all paragraphs to replace
      const firstPara = paragraphs.items[startIdx];
      const lastPara = paragraphs.items[endIdx];

      // Try to read the existing list's numId from the first paragraph's OOXML
      let existingNumId = null;
      let existingBaseIlvl = 0;
      try {
        const firstParaOoxml = firstPara.getOoxml();
        await context.sync();

        const numIdMatch = firstParaOoxml.value.match(/w:numId w:val="(\d+)"/);
        const ilvlMatch = firstParaOoxml.value.match(/w:ilvl w:val="(\d+)"/);

        if (numIdMatch) {
          existingNumId = numIdMatch[1];
          existingBaseIlvl = ilvlMatch ? parseInt(ilvlMatch[1], 10) : 0;
          console.log(`[executeEditList] Found existing numId: ${existingNumId}, base ilvl: ${existingBaseIlvl}`);
        }
      } catch (oxmlError) {
        console.warn(`[executeEditList] Could not read existing OOXML:`, oxmlError.message);
      }

      // Get ranges to create a combined range
      const startRange = firstPara.getRange("Start");
      const endRange = lastPara.getRange("End");
      const fullRange = startRange.expandTo(endRange);

      await context.sync();

      // Build HTML list
      const listTag = listType === "numbered" ? "ol" : "ul";

      // Map numbering style to CSS list-style-type
      let cssListStyleType = "disc"; // default for bullet
      if (listType === "numbered") {
        const styleMap = {
          "decimal": "decimal",
          "lowerAlpha": "lower-alpha",
          "upperAlpha": "upper-alpha",
          "lowerRoman": "lower-roman",
          "upperRoman": "upper-roman"
        };
        cssListStyleType = styleMap[numberingStyle] || "decimal";
      }

      const listStyle = `style="list-style-type: ${cssListStyleType}; margin-left: 0; padding-left: 40px;"`;

      // Map numberingStyle to OOXML numFmt values for direct OOXML insertion
      const numFmtMap = {
        "decimal": "decimal",
        "lowerAlpha": "lowerLetter",
        "upperAlpha": "upperLetter",
        "lowerRoman": "lowerRoman",
        "upperRoman": "upperRoman"
      };
      const numFmt = numFmtMap[numberingStyle] || "decimal";

      // Determine template numId when creating a new list (no existing numId)
      // We'll use numId 100+ for custom styles to avoid conflicts
      let templateNumId = existingNumId;
      if (!templateNumId) {
        if (listType === "bullet") {
          templateNumId = "1"; // Default bullet
        } else {
          // For numbered lists, use numId 2 (default decimal) - the numFmt will override display
          templateNumId = "2";
        }
        console.log(`[executeEditList] No existing numId, using template: ${templateNumId}, numFmt: ${numFmt}`);
      }
      // Detect hierarchy from leading whitespace indentation (4 spaces = 1 level)
      // Also strip any leading list markers from items to avoid doubled numbering
      const markersRegex = /^((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)/;

      // Analyze items for hierarchy based on leading whitespace
      const itemsWithLevels = newItems.map(item => {
        // Count leading spaces/tabs
        const indentMatch = item.match(/^(\s*)/);
        const indentSize = indentMatch ? indentMatch[1].length : 0;
        const level = Math.floor(indentSize / 4); // 4 spaces per level

        // Strip leading whitespace
        let stripped = item.trim();

        // Also strip any list markers (1., a., -, etc.)
        const markerMatch = stripped.match(markersRegex);
        if (markerMatch) {
          stripped = stripped.replace(markersRegex, '');
          console.log(`[executeEditList] Stripped marker: "${markerMatch[1].trim()}" from item`);
        }

        console.log(`[executeEditList] Level: ${level}, Text: "${stripped.substring(0, 40)}..."`);

        return { text: stripped.trim(), level };
      });

      // SURGICAL APPROACH: Edit existing paragraphs in place
      // This preserves the document's existing formatting better than bulk replacement
      const existingCount = endIdx - startIdx + 1;
      const newCount = itemsWithLevels.length;

      console.log(`[executeEditList] Surgical mode: ${existingCount} existing → ${newCount} new items`);

      // PHASE 1: Edit existing paragraphs with new text (keeping their style)
      const editLimit = Math.min(existingCount, newCount);
      for (let i = 0; i < editLimit; i++) {
        const para = paragraphs.items[startIdx + i];
        const item = itemsWithLevels[i];

        // Get the paragraph's current text and OOXML for debugging
        para.load("text");
        const paraOoxml = para.getOoxml();
        await context.sync();

        const originalText = para.text.trim();
        console.log(`[executeEditList] P${startIdx + i + 1} BEFORE: "${originalText.substring(0, 50)}..."`);
        console.log(`[executeEditList] P${startIdx + i + 1} NEW: "${item.text.substring(0, 50)}..."`);

        const numIdMatch = paraOoxml.value.match(/w:numId w:val="(\d+)"/);
        const ilvlMatch = paraOoxml.value.match(/w:ilvl w:val="(\d+)"/);
        const currentNumId = numIdMatch ? numIdMatch[1] : (existingNumId || '1');
        const currentIlvl = ilvlMatch ? ilvlMatch[1] : '0';
        const newIlvl = existingBaseIlvl + item.level;

        console.log(`[executeEditList] P${startIdx + i + 1} numId=${currentNumId}, ilvl: ${currentIlvl} → ${newIlvl}`);

        // Build OOXML that preserves the paragraph's numbering but updates text and ilvl
        const escapedText = item.text
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');

        // Use templateNumId when paragraph doesn't have existing numbering
        const numIdForPara = currentNumId === (existingNumId || '1') && !existingNumId ? templateNumId : currentNumId;

        // Build the numbering.xml part if creating a new list with custom style
        let numberingPart = '';
        if (!existingNumId && listType === "numbered" && numFmt !== 'decimal') {
          // Include a complete numbering definition for custom styles
          // Use a high numId to avoid conflicts (100+)
          const customNumId = '100';
          numberingPart = `
            <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
              <pkg:xmlData>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="100">
                    <w:multiLevelType w:val="multilevel"/>
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="${numFmt}"/>
                      <w:lvlText w:val="%1."/>
                      <w:lvlJc w:val="left"/>
                      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="100">
                    <w:abstractNumId w:val="100"/>
                  </w:num>
                </w:numbering>
              </pkg:xmlData>
            </pkg:part>`;
          // Use this custom numId for the paragraph
        }

        const oxmlPara = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>${numberingPart}
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
              <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="${newIlvl}"/>
                          <w:numId w:val="${numberingPart ? '100' : numIdForPara}"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r>
                        <w:t xml:space="preserve">${escapedText}</w:t>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
              </pkg:xmlData>
            </pkg:part>
          </pkg:package>`;

        // Use the paragraph's range for replacement (not the paragraph object)
        const paraRange = para.getRange("Whole");
        paraRange.insertOoxml(oxmlPara, "Replace");
        console.log(`[executeEditList] Replaced P${startIdx + i + 1} range with OOXML`);
      }
      await context.sync();

      // PHASE 2: Insert new paragraphs if more items than existing
      if (newCount > existingCount) {
        console.log(`[executeEditList] Phase 2: Inserting ${newCount - existingCount} new paragraphs`);

        // Reload paragraphs after Phase 1 edits
        paragraphs.load("items");
        await context.sync();

        // Get the last edited paragraph to insert after
        const lastEditedIdx = startIdx + existingCount - 1;
        let insertAfterPara = paragraphs.items[lastEditedIdx];

        console.log(`[executeEditList] Will insert after P${lastEditedIdx + 1}`);

        for (let i = existingCount; i < newCount; i++) {
          const item = itemsWithLevels[i];
          const ilvl = existingBaseIlvl + item.level;
          // Use the same numId as determined earlier for this list
          const numIdForPhase2 = (!existingNumId && listType === "numbered" && numFmt !== 'decimal') ? '100' : (existingNumId || templateNumId);

          console.log(`[executeEditList] Inserting item ${i + 1}: "${item.text.substring(0, 30)}..." at ilvl=${ilvl}`);

          // Build OOXML for the new list item (same approach as Phase 1)
          const escapedText = item.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');

          // Include numbering.xml part for first Phase 2 item if using custom style
          // (subsequent items can reuse the numId that was established)
          let phase2NumberingPart = '';
          if (i === existingCount && !existingNumId && listType === "numbered" && numFmt !== 'decimal') {
            phase2NumberingPart = `
            <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
              <pkg:xmlData>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="100">
                    <w:multiLevelType w:val="multilevel"/>
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="${numFmt}"/>
                      <w:lvlText w:val="%1."/>
                      <w:lvlJc w:val="left"/>
                      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="100">
                    <w:abstractNumId w:val="100"/>
                  </w:num>
                </w:numbering>
              </pkg:xmlData>
            </pkg:part>`;
          }

          const oxmlPara = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
              <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData>
                  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                  </Relationships>
                </pkg:xmlData>
              </pkg:part>${phase2NumberingPart}
              <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                <pkg:xmlData>
                  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:body>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val="ListParagraph"/>
                          <w:numPr>
                            <w:ilvl w:val="${ilvl}"/>
                            <w:numId w:val="${numIdForPhase2}"/>
                          </w:numPr>
                        </w:pPr>
                        <w:r>
                          <w:t xml:space="preserve">${escapedText}</w:t>
                        </w:r>
                      </w:p>
                    </w:body>
                  </w:document>
                </pkg:xmlData>
              </pkg:part>
            </pkg:package>`;

          // Insert OOXML after the last paragraph
          const insertRange = insertAfterPara.getRange("End");
          insertRange.insertOoxml(oxmlPara, "After");
          await context.sync();

          // Reload paragraphs to get the newly inserted one for next iteration
          paragraphs.load("items");
          await context.sync();

          // The new paragraph should be at lastEditedIdx + (i - existingCount + 1)
          const newParaIdx = lastEditedIdx + (i - existingCount + 1);
          if (newParaIdx < paragraphs.items.length) {
            insertAfterPara = paragraphs.items[newParaIdx];
          }

          console.log(`[executeEditList] Inserted new paragraph ${i + 1} via OOXML`);
        }
      }

      // PHASE 3: Delete excess paragraphs if fewer new items
      if (newCount < existingCount) {
        // Delete from the end to avoid index shifting
        for (let i = existingCount - 1; i >= newCount; i--) {
          const paraToDelete = paragraphs.items[startIdx + i];
          paraToDelete.delete();
          console.log(`[executeEditList] Deleted excess P${startIdx + i + 1}`);
        }
        await context.sync();
      }

      // Restore original tracking mode if we changed it
      if (redlineEnabled && originalChangeTrackingMode !== null &&
        originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
        context.document.changeTrackingMode = originalChangeTrackingMode;
        await context.sync();
      }

      console.log(`Successfully replaced paragraphs with ${listType} list`);
    });

    return {
      success: true,
      message: `Successfully created ${listType} list with ${newItems.length} items.`
    };
  } catch (error) {
    console.error("Error in executeEditList:", error);
    return {
      success: false,
      message: `Failed to edit list: ${error.message}`
    };
  }
}

/**
 * Execute convert_headers_to_list tool - converts non-contiguous headers to a numbered list
 * This handles the case where headers like "1. PURPOSE", "2. DEFINITION" have body text between them
 * @param {number[]} paragraphIndices - Array of 1-based paragraph indices of headers to convert
 * @param {string[]} newHeaderTexts - Optional array of new header texts (without numbers)
 * @param {string} numberingFormat - Optional: 'arabic' (default), 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'
 */
async function executeConvertHeadersToList(paragraphIndices, newHeaderTexts, numberingFormat) {
  if (!paragraphIndices || paragraphIndices.length === 0) {
    return { success: false, message: "No paragraph indices provided." };
  }

  // Default to arabic if not specified
  const format = numberingFormat || "arabic";
  console.log(`executeConvertHeadersToList: Converting ${paragraphIndices.length} headers to ${format} numbered list`);

  try {
    await Word.run(async (context) => {
      // Enable track changes if redline setting is enabled
      const redlineEnabled = loadRedlineSetting();
      let originalChangeTrackingMode = null;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          originalChangeTrackingMode = doc.changeTrackingMode;
          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
          }
        } catch (trackError) {
          console.warn("Could not enable track changes:", trackError);
        }
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // Sort indices to process in order
      const sortedIndices = [...paragraphIndices].sort((a, b) => a - b);

      // Validate all indices
      for (const idx of sortedIndices) {
        const pIdx = idx - 1;
        if (pIdx < 0 || pIdx >= paragraphs.items.length) {
          throw new Error(`Invalid paragraph index: ${idx}`);
        }
      }

      // Get the first header paragraph and start a new list
      const firstIdx = sortedIndices[0] - 1;
      const firstPara = paragraphs.items[firstIdx];
      firstPara.load("text");
      await context.sync();

      // Strip manual numbering from the first header if present
      let firstText = firstPara.text || "";
      const numberPattern = /^\s*\d+[.\)]\s*/;
      firstText = firstText.replace(numberPattern, "").trim();

      // Use new text if provided
      if (newHeaderTexts && newHeaderTexts.length > 0) {
        firstText = newHeaderTexts[0];
      }

      // Clear and replace the paragraph content
      firstPara.clear();
      firstPara.insertText(firstText, Word.InsertLocation.start);
      await context.sync();

      // Start a new list on this paragraph
      const list = firstPara.startNewList();
      await context.sync();

      // Load the list to set its numbering format
      list.load("id, levelTypes");
      await context.sync();

      // Map format string to Word.ListNumbering constant
      const numberingMap = {
        "arabic": Word.ListNumbering.arabic,
        "lowerLetter": Word.ListNumbering.lowerLetter,
        "upperLetter": Word.ListNumbering.upperLetter,
        "lowerRoman": Word.ListNumbering.lowerRoman,
        "upperRoman": Word.ListNumbering.upperRoman
      };

      const wordNumbering = numberingMap[format] || Word.ListNumbering.arabic;

      // Set the list to use the specified numbering format
      try {
        list.setLevelNumbering(0, wordNumbering);
        await context.sync();
        console.log(`Set list numbering to ${format}`);
      } catch (numError) {
        console.warn("Could not set level numbering, trying style approach:", numError);
        // Fallback: apply numbered list style
        firstPara.styleBuiltIn = Word.BuiltInStyleName.listNumber;
        await context.sync();
      }

      console.log(`Started new numbered list on paragraph ${sortedIndices[0]}`);

      // For remaining headers, attach them to the same list
      for (let i = 1; i < sortedIndices.length; i++) {
        const pIdx = sortedIndices[i] - 1;
        const para = paragraphs.items[pIdx];
        para.load("text");
        await context.sync();

        // Strip manual numbering
        let paraText = para.text || "";
        paraText = paraText.replace(numberPattern, "").trim();

        // Use new text if provided
        if (newHeaderTexts && newHeaderTexts.length > i) {
          paraText = newHeaderTexts[i];
        }

        // Clear and replace the paragraph content
        para.clear();
        para.insertText(paraText, Word.InsertLocation.start);
        await context.sync();

        // Attach to the list
        try {
          para.attachToList(list.id, 0); // level 0
          await context.sync();
          console.log(`Attached paragraph ${sortedIndices[i]} to list`);
        } catch (attachError) {
          console.warn(`Could not attach paragraph ${sortedIndices[i]}, using style:`, attachError);
          para.styleBuiltIn = Word.BuiltInStyleName.listNumber;
          await context.sync();
        }
      }

      // Restore tracking mode
      if (redlineEnabled && originalChangeTrackingMode !== null &&
        originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
        context.document.changeTrackingMode = originalChangeTrackingMode;
        await context.sync();
      }

      console.log(`Successfully converted ${sortedIndices.length} headers to numbered list`);
    });

    return {
      success: true,
      message: `Successfully converted ${paragraphIndices.length} headers to a numbered list.`
    };
  } catch (error) {
    console.error("Error in executeConvertHeadersToList:", error);
    return {
      success: false,
      message: `Failed to convert headers to list: ${error.message}`
    };
  }
}

/**
 * Execute edit_table tool - performs table operations
 * @param {number} paragraphIndex - 1-based index of any paragraph in the table
 * @param {string} action - "replace_content", "add_row", "delete_row", "update_cell"
 * @param {Array} content - Content for the operation
 * @param {number} targetRow - Target row index (0-based)
 * @param {number} targetColumn - Target column index (0-based)
 */
async function executeEditTable(paragraphIndex, action, content, targetRow, targetColumn) {
  try {
    await Word.run(async (context) => {
      // Enable track changes if redline setting is enabled
      const redlineEnabled = loadRedlineSetting();
      let originalChangeTrackingMode = null;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          originalChangeTrackingMode = doc.changeTrackingMode;
          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
          }
        } catch (trackError) {
          console.warn("Could not enable track changes:", trackError);
        }
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const pIdx = paragraphIndex - 1;
      if (pIdx < 0 || pIdx >= paragraphs.items.length) {
        throw new Error(`Invalid paragraph index: ${paragraphIndex}`);
      }

      const targetPara = paragraphs.items[pIdx];
      targetPara.load("parentTableOrNullObject");
      await context.sync();

      if (targetPara.parentTableOrNullObject.isNullObject) {
        throw new Error(`Paragraph ${paragraphIndex} is not inside a table`);
      }

      const table = targetPara.parentTableOrNullObject;
      table.load("rowCount, rows");
      await context.sync();

      if (action === "replace_content") {
        // content should be 2D array [[row1cells], [row2cells], ...]
        if (!content || !Array.isArray(content)) {
          throw new Error("replace_content requires a 2D array of content");
        }

        // Load all rows and cells
        table.rows.load("items");
        await context.sync();

        for (let r = 0; r < content.length && r < table.rows.items.length; r++) {
          const row = table.rows.items[r];
          row.cells.load("items");
          await context.sync();

          for (let c = 0; c < content[r].length && c < row.cells.items.length; c++) {
            const cell = row.cells.items[c];
            const cellBody = cell.body;
            cellBody.clear();
            cellBody.insertText(content[r][c], Word.InsertLocation.start);
          }
        }
        await context.sync();

      } else if (action === "add_row") {
        // content should be array of cell values for the new row
        if (!content || !Array.isArray(content)) {
          throw new Error("add_row requires an array of cell values");
        }

        const insertAt = targetRow !== undefined ? targetRow : table.rowCount;
        const newRow = table.addRows(Word.InsertLocation.end, 1, [content]);
        await context.sync();

      } else if (action === "delete_row") {
        if (targetRow === undefined) {
          throw new Error("delete_row requires targetRow");
        }

        table.rows.load("items");
        await context.sync();

        if (targetRow < 0 || targetRow >= table.rows.items.length) {
          throw new Error(`Invalid row index: ${targetRow}`);
        }

        table.rows.items[targetRow].delete();
        await context.sync();

      } else if (action === "update_cell") {
        if (targetRow === undefined || targetColumn === undefined) {
          throw new Error("update_cell requires targetRow and targetColumn");
        }
        if (!content || content.length === 0) {
          throw new Error("update_cell requires content");
        }

        table.rows.load("items");
        await context.sync();

        if (targetRow < 0 || targetRow >= table.rows.items.length) {
          throw new Error(`Invalid row index: ${targetRow}`);
        }

        const row = table.rows.items[targetRow];
        row.cells.load("items");
        await context.sync();

        if (targetColumn < 0 || targetColumn >= row.cells.items.length) {
          throw new Error(`Invalid column index: ${targetColumn}`);
        }

        const cell = row.cells.items[targetColumn];
        const cellBody = cell.body;
        cellBody.clear();
        cellBody.insertText(content[0], Word.InsertLocation.start);
        await context.sync();

      } else {
        throw new Error(`Unknown table action: ${action}`);
      }
    });

    return {
      success: true,
      message: `Successfully performed table operation: ${action}`
    };
  } catch (error) {
    console.error("Error in executeEditTable:", error);
    return {
      success: false,
      message: `Failed to edit table: ${error.message}`
    };
  }
}

/**
 * Execute edit_section tool - edits a legal contract section
 * @param {number} sectionHeaderIndex - 1-based index of the section header paragraph
 * @param {string} newHeaderText - Optional new text for the header (preserves numbering)
 * @param {string[]} newBodyParagraphs - Optional new body paragraphs
 * @param {boolean} preserveSubsections - Whether to preserve subsections
 */
async function executeEditSection(sectionHeaderIndex, newHeaderText, newBodyParagraphs, preserveSubsections) {
  try {
    let editCount = 0;

    await Word.run(async (context) => {
      // Enable track changes if redline setting is enabled
      const redlineEnabled = loadRedlineSetting();
      let originalChangeTrackingMode = null;

      if (redlineEnabled) {
        try {
          const doc = context.document;
          doc.load("changeTrackingMode");
          await context.sync();

          originalChangeTrackingMode = doc.changeTrackingMode;
          if (originalChangeTrackingMode !== Word.ChangeTrackingMode.trackAll) {
            doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
            await context.sync();
          }
        } catch (trackError) {
          console.warn("Could not enable track changes:", trackError);
        }
      }

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const headerIdx = sectionHeaderIndex - 1;
      if (headerIdx < 0 || headerIdx >= paragraphs.items.length) {
        throw new Error(`Invalid section header index: ${sectionHeaderIndex}`);
      }

      // Load properties to understand section structure
      for (const para of paragraphs.items) {
        para.load("text, listItemOrNullObject");
      }
      await context.sync();

      // Load list levels
      for (const para of paragraphs.items) {
        if (!para.listItemOrNullObject.isNullObject) {
          para.listItemOrNullObject.load("level");
        }
      }
      await context.sync();

      const headerPara = paragraphs.items[headerIdx];

      // Check that header is a list item (section header)
      if (headerPara.listItemOrNullObject.isNullObject) {
        throw new Error(`Paragraph ${sectionHeaderIndex} is not a section header (not a list item)`);
      }

      const headerLevel = headerPara.listItemOrNullObject.level || 0;

      // Find the end of this section (next list item at same or higher level)
      let sectionEndIdx = paragraphs.items.length - 1;
      for (let i = headerIdx + 1; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        if (!para.listItemOrNullObject.isNullObject) {
          const level = para.listItemOrNullObject.level || 0;
          if (level <= headerLevel) {
            // Found next section at same or higher level
            sectionEndIdx = i - 1;
            break;
          } else if (preserveSubsections) {
            // Found a subsection - stop here if preserving
            sectionEndIdx = i - 1;
            break;
          }
        }
      }

      // Update header text if provided
      if (newHeaderText !== undefined && newHeaderText !== null) {
        // Extract the list number/letter prefix from current text
        const currentText = headerPara.text || "";
        const numberMatch = currentText.match(/^(\d+\.?\s*|\([a-z]\)\s*|[a-z]\.\s*|[ivxlcdm]+\.\s*)/i);

        if (numberMatch) {
          // Preserve the numbering prefix
          headerPara.insertText(numberMatch[1] + newHeaderText, Word.InsertLocation.replace);
        } else {
          headerPara.insertText(newHeaderText, Word.InsertLocation.replace);
        }
        editCount++;
      }

      // Replace body paragraphs if provided
      if (newBodyParagraphs && newBodyParagraphs.length > 0) {
        // Delete existing body paragraphs (from end to start)
        for (let i = sectionEndIdx; i > headerIdx; i--) {
          paragraphs.items[i].delete();
        }
        await context.sync();

        // Insert new body paragraphs after header
        let insertAfter = headerPara;
        for (const bodyText of newBodyParagraphs) {
          const newPara = insertAfter.insertParagraph(bodyText, Word.InsertLocation.after);
          insertAfter = newPara;
          editCount++;
        }
      }

      await context.sync();
    });

    if (editCount === 0) {
      return {
        success: true,
        message: "No changes were specified for the section."
      };
    }

    return {
      success: true,
      message: `Successfully edited section at P${sectionHeaderIndex} (${editCount} changes).`
    };
  } catch (error) {
    console.error("Error in executeEditSection:", error);
    return {
      success: false,
      message: `Failed to edit section: ${error.message}`
    };
  }
}
