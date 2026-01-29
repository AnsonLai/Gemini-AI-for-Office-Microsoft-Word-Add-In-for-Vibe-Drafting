/* global document */

import { marked } from 'marked';

let onCancelRequest = null;
let onRestoreCheckpoint = null;

function registerChatUiHandlers(handlers = {}) {
  if (handlers.onCancelRequest) {
    onCancelRequest = handlers.onCancelRequest;
  }
  if (handlers.onRestoreCheckpoint) {
    onRestoreCheckpoint = handlers.onRestoreCheckpoint;
  }
}

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
        if (onCancelRequest) {
          onCancelRequest();
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
  revertBtn.onclick = () => {
    if (onRestoreCheckpoint) {
      onRestoreCheckpoint(checkpointIndex);
    }
  };

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

export {
  registerChatUiHandlers,
  setupScrollToBottom,
  createTypingIndicator,
  shakeInput,
  addMessageToChat,
  updateSystemMessage,
  addRetryButton,
  hideAllRetryButtons,
  removeMessage
};
