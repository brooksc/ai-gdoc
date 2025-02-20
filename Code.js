/**
 * Wed Feb 19 20:55:39 PST 2025
 */
/**
 * AI Editor for Google Docs
 * This script allows users to highlight text, add "AI: [instruction]" comments,
 * and process them using local Ollama models.
 */

// Comment state markers
const COMMENT_STATE = {
  ACCEPTED: '[STATE:ACCEPTED]',
  REJECTED: '[STATE:REJECTED]',
  PROCESSING: '[STATE:PROCESSING]'
};

// Retry configuration
const RETRY_CONFIG = {
  MAX_ATTEMPTS: 3,
  INITIAL_DELAY_MS: 1000,
  MAX_DELAY_MS: 5000
};

/**
 * Sleep for a given number of milliseconds
 * 
 * @param {Number} ms - Milliseconds to sleep
 * @return {Promise} Promise that resolves after the delay
 */
function sleep(ms) {
  return new Promise(resolve => Utilities.sleep(ms));
}

/**
 * Calculate exponential backoff delay
 * 
 * @param {Number} attempt - Current attempt number (0-based)
 * @return {Number} Delay in milliseconds
 */
function calculateBackoffDelay(attempt) {
  const delay = Math.min(
    RETRY_CONFIG.INITIAL_DELAY_MS * Math.pow(2, attempt),
    RETRY_CONFIG.MAX_DELAY_MS
  );
  // Add jitter to prevent thundering herd
  return delay + (Math.random() * 1000);
}

/**
 * Retry updating a comment with exponential backoff
 * 
 * @param {String} fileId - Document ID
 * @param {String} commentId - Comment ID
 * @param {String} content - Comment content
 * @param {Object} options - Additional options (resolved, etc)
 * @return {Object} Result object with success status and error if any
 */
function retryCommentUpdate(fileId, commentId, content, options = {}) {
  let lastError = null;
  
  for (let attempt = 0; attempt < RETRY_CONFIG.MAX_ATTEMPTS; attempt++) {
    try {
      Logger.log("aiedit-debug: Updating comment (attempt " + (attempt + 1) + ")");
      Logger.log("aiedit-debug: Comment params:", { content, resolved: options.resolved });
      
      // Use update instead of insert for existing comments
      Drive.Comments.update({
        content: content,
        resolved: !!options.resolved  // Ensure boolean
      }, fileId, commentId);
      
      Logger.log("aiedit-debug: Comment update successful");
      return { success: true };
    } catch (error) {
      lastError = error;
      Logger.log("aiedit-debug: Comment update failed: " + error.message);
      
      if (attempt < RETRY_CONFIG.MAX_ATTEMPTS - 1) {
        const delay = calculateBackoffDelay(attempt);
        Logger.log("aiedit-debug: Retrying after " + delay + "ms");
        sleep(delay);
      }
    }
  }
  
  Logger.log("aiedit-debug: All comment update attempts failed");
  return {
    success: false,
    error: lastError
  };
}

/**
 * Creates the AI Editor menu when the document is opened
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('AI Editor')
    .addItem('Open Editor', 'showSidebar')
    .addToUi();
}

/**
 * Displays the AI Editor sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
    .setTitle('AI Editor')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Sanitize text for safe document insertion
 * 
 * @param {String} text - Text to sanitize
 * @return {String} Sanitized text
 */
function sanitizeText(text) {
  return text
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, '') // Remove control characters
    .replace(/\u2028|\u2029/g, '\n')              // Normalize line separators
    .trim();
}

/**
 * Get the current state of a comment from its replies
 * 
 * @param {Object} comment - Comment object from Drive API
 * @return {String|null} Current state or null if unprocessed
 */
function getCommentState(comment) {
  if (!comment.replies || comment.replies.length === 0) {
    return null;
  }
  
  // Check replies in reverse order to get most recent state
  for (let i = comment.replies.length - 1; i >= 0; i--) {
    const reply = comment.replies[i];
    for (const state in COMMENT_STATE) {
      if (reply.content.includes(COMMENT_STATE[state])) {
        return COMMENT_STATE[state];
      }
    }
  }
  
  return null;
}

/**
 * Retrieves all comments in the document that contain "AI:" prefix
 * and returns their text selections and instructions
 * 
 * @return {Array} Array of objects containing comment info
 */
function getAIComments() {
  const doc = DocumentApp.getActiveDocument();
  const fileId = doc.getId();
  
  try {
    // Get all comments from the document using Drive API
    const response = Drive.Comments.list(fileId, {
      fields: "comments(id,content,quotedFileContent,replies,anchor,resolved)",
      maxResults: 100
    });
    
    const comments = response.comments || [];
    Logger.log("aiedit: Retrieved " + comments.length + " total comments");
    
    // Filter for AI comments that are either unprocessed or rejected
    const aiComments = comments.filter(comment => {
      const content = comment.content.trim();
      const isAIComment = content.startsWith('AI:');
      const isActive = !comment.resolved;
      const hasQuotedText = comment.quotedFileContent && comment.quotedFileContent.value;
      const hasValidAnchor = comment.anchor && comment.anchor.length > 0;
      
      // Get current state from replies
      const state = getCommentState(comment);
      
      // Comment is eligible if:
      // 1. Never processed (no state)
      // 2. Was rejected (has REJECTED state)
      // 3. Not currently being processed
      const isEligible = !state || state === COMMENT_STATE.REJECTED;
      
      return isAIComment && isActive && hasQuotedText && hasValidAnchor && isEligible;
    });
    
    Logger.log("aiedit: Found " + aiComments.length + " eligible AI comments");
    
    // Map to our internal format with full context
    return aiComments.map(comment => ({
      id: comment.id,
      instruction: comment.content.trim().substring(3).trim(), // Remove 'AI:' prefix
      text: comment.quotedFileContent.value,
      anchor: comment.anchor,
      state: getCommentState(comment) || 'unprocessed'
    }));
  } catch (e) {
    Logger.log("aiedit: Error retrieving comments: " + e.message);
    throw new Error("Failed to retrieve comments: " + e.message);
  }
}

/**
 * Validate comment ID format and existence
 * 
 * @param {String} fileId - Document ID
 * @param {String} commentId - Comment ID to validate
 * @return {Boolean} Whether the comment ID is valid
 */
function validateCommentId(fileId, commentId) {
  if (!commentId || typeof commentId !== 'string' || commentId.length === 0) {
    return false;
  }
  
  try {
    const comment = Drive.Comments.get(fileId, commentId, {
      fields: 'id'
    });
    return comment && comment.id === commentId;
  } catch (e) {
    Logger.log("aiedit: Error validating comment ID: " + e.message);
    return false;
  }
}

/**
 * Verify text location and return success status
 * 
 * @param {String} fileId - Document ID
 * @param {String} commentId - Comment ID
 * @param {String} originalText - Text to verify
 * @return {Object} Success status and location info
 */
function verifyTextLocation(fileId, commentId, originalText) {
  try {
    // Get document
    const doc = DocumentApp.getActiveDocument();
    
    // Validate comment ID first
    if (!validateCommentId(fileId, commentId)) {
      return { success: false, error: "Invalid comment ID" };
    }
    
    // Get fresh comment data
    const comment = Drive.Comments.get(fileId, commentId, {
      fields: 'id,anchor,quotedFileContent,content'
    });
    
    if (!comment || !comment.anchor || !comment.quotedFileContent) {
      return { 
        success: false, 
        error: "Comment is missing required data" 
      };
    }
    
    // Get the document body
    const body = doc.getBody();
    
    // Normalize the text for searching
    const normalizedOriginal = originalText.replace(/[\u0000-\u001F\u007F-\u009F]/g, '')
      .replace(/\u2028|\u2029/g, '\n')
      .replace(/\s+/g, ' ')
      .trim();
    
    // Find all occurrences of the text
    let foundElements = [];
    let searchResult = body.findText(normalizedOriginal);
    
    while (searchResult) {
      const element = searchResult.getElement();
      const startOffset = searchResult.getStartOffset();
      const endOffset = searchResult.getEndOffsetInclusive();
      
      foundElements.push({
        element: element,
        startOffset: startOffset,
        endOffset: endOffset
      });
      
      searchResult = body.findText(normalizedOriginal, searchResult);
    }
    
    if (foundElements.length === 0) {
      return {
        success: false,
        error: "Could not find the text in the document"
      };
    }
    
    // If we have exactly one match, use it
    if (foundElements.length === 1) {
      const match = foundElements[0];
      return {
        success: true,
        location: {
          element: match.element,
          startOffset: match.startOffset,
          endOffset: match.endOffset,
          anchorStart: match.startOffset,
          anchorEnd: match.endOffset
        }
      };
    }
    
    // Multiple matches - try to find the best one using context
    const quotedContext = comment.quotedFileContent.value;
    let bestMatch = null;
    let bestScore = -1;
    
    for (const match of foundElements) {
      const elementText = match.element.getText();
      const contextStart = Math.max(0, match.startOffset - 50);
      const contextEnd = Math.min(elementText.length, match.endOffset + 51);
      const context = elementText.substring(contextStart, contextEnd);
      
      const contextSimilarity = calculateSimilarity(context, quotedContext);
      
      if (contextSimilarity > bestScore) {
        bestScore = contextSimilarity;
        bestMatch = match;
      }
    }
    
    if (bestMatch && bestScore > 0.8) {  // 80% confidence threshold
      return {
        success: true,
        location: {
          element: bestMatch.element,
          startOffset: bestMatch.startOffset,
          endOffset: bestMatch.endOffset,
          anchorStart: bestMatch.startOffset,
          anchorEnd: bestMatch.endOffset
        }
      };
    }
    
    return {
      success: false,
      error: "Could not reliably determine text location"
    };
    
  } catch (e) {
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Calculate similarity between two strings
 * Uses Levenshtein distance normalized by length
 * 
 * @param {String} str1 - First string
 * @param {String} str2 - Second string
 * @return {Number} Similarity score between 0 and 1
 */
function calculateSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  
  // Normalize strings
  str1 = str1.replace(/\s+/g, ' ').trim().toLowerCase();
  str2 = str2.replace(/\s+/g, ' ').trim().toLowerCase();
  
  if (str1 === str2) return 1;
  
  const len1 = str1.length;
  const len2 = str2.length;
  const matrix = Array(len2 + 1).fill().map(() => Array(len1 + 1).fill(0));
  
  for (let i = 0; i <= len1; i++) matrix[0][i] = i;
  for (let j = 0; j <= len2; j++) matrix[j][0] = j;
  
  for (let j = 1; j <= len2; j++) {
    for (let i = 1; i <= len1; i++) {
      const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
      matrix[j][i] = Math.min(
        matrix[j - 1][i] + 1,      // deletion
        matrix[j][i - 1] + 1,      // insertion
        matrix[j - 1][i - 1] + cost // substitution
      );
    }
  }
  
  const distance = matrix[len2][len1];
  const maxLen = Math.max(len1, len2);
  return maxLen === 0 ? 1 : 1 - (distance / maxLen);
}

/**
 * Highlight text in the document to indicate a conflict or issue
 * 
 * @param {Object} location - Location information for the text
 * @param {String} color - Background color for highlighting
 */
function highlightConflictArea(location) {
  if (!location || !location.element) return;
  
  try {
    const originalStyle = location.element.getBackgroundColor();
    
    // Highlight in yellow to indicate conflict
    location.element.setBackgroundColor(
      location.startOffset,
      location.endOffset,
      '#FFEB3B'  // Material Design Yellow 500
    );
    
    // Reset highlight after 5 seconds
    Utilities.sleep(5000);
    location.element.setBackgroundColor(
      location.startOffset,
      location.endOffset,
      originalStyle
    );
  } catch (e) {
    Logger.log("aiedit: Warning - Could not highlight conflict area: " + e.message);
  }
}

/**
 * Check for document changes and verify text integrity
 * 
 * @param {String} initialVersion - Initial document version
 * @param {String} currentText - Current text to verify
 * @param {Object} location - Location information
 * @return {Object} Status object with change details
 */
function checkDocumentChanges(initialVersion, currentText, location) {
  const doc = DocumentApp.getActiveDocument();
  const currentVersion = doc.getBody().getText();
  
  if (currentVersion === initialVersion) {
    return { changed: false };
  }
  
  // Document changed, analyze the changes
  const originalTextIntact = currentText.substring(
    location.anchorStart,
    location.anchorEnd + 1
  ) === location.element.getText().substring(
    location.startOffset,
    location.endOffset + 1
  );
  
  return {
    changed: true,
    targetAreaChanged: !originalTextIntact,
    changeLocation: location
  };
}

/**
 * Apply AI-generated text by resolving the comment and updating the document
 * 
 * @param {String} fileId - ID of the document
 * @param {String} commentId - ID of the comment
 * @param {String} suggestedText - AI-generated replacement text
 * @param {Boolean} accepted - Whether the suggestion was accepted
 * @return {Boolean} Success status
 */
function applyAIEdit(fileId, commentId, suggestedText, accepted) {
  const doc = DocumentApp.getActiveDocument();
  let originalText = null;
  let commentResolved = false;
  let location = null;
  
  try {
    // Validate inputs
    if (!fileId || !commentId || !suggestedText) {
      throw new Error("Missing required parameters for applying AI edit");
    }
    
    if (!validateCommentId(fileId, commentId)) {
      throw new Error("The comment no longer exists or is inaccessible");
    }
    
    // Get the comment and verify it exists
    const comment = Drive.Comments.get(fileId, commentId, {
      fields: 'id,quotedFileContent,anchor,resolved'
    });
    
    if (!comment) {
      throw new Error("Could not find the specified comment");
    }
    
    if (comment.resolved) {
      throw new Error("This comment has already been resolved");
    }
    
    if (!comment.quotedFileContent || !comment.quotedFileContent.value) {
      throw new Error("The comment is missing the original text selection");
    }
    
    // Store original text
    originalText = comment.quotedFileContent.value;
    
    if (accepted) {
      // First update the comment to mark it as being processed
      const processingUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.PROCESSING + '\n\nApplying suggested changes...',
        { resolved: false }
      );
      
      if (!processingUpdate.success) {
        throw new Error("Failed to mark comment as processing");
      }
      
      // Verify text location before making changes
      const verifyResult = verifyTextLocation(fileId, commentId, originalText);
      if (!verifyResult.success) {
        throw new Error(verifyResult.error || "Could not verify text location");
      }
      
      location = verifyResult.location;
      
      // Prepare sanitized text
      const sanitizedText = sanitizeText(suggestedText);
      if (!sanitizedText) {
        throw new Error("The suggested text is empty after sanitization");
      }
      
      // Replace the text
      location.element.asText().deleteText(location.startOffset, location.endOffset);
      location.element.asText().insertText(location.startOffset, sanitizedText);
      
      // Verify the replacement
      const verifyText = location.element.asText().getText()
        .substring(location.startOffset, location.startOffset + sanitizedText.length);
      
      if (verifyText !== sanitizedText) {
        throw new Error("Failed to verify text replacement");
      }
      
      // Update comment to mark as accepted
      const acceptUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.ACCEPTED + '\n\nChanges applied successfully:\n\n' +
        'Original text:\n' + originalText + '\n\n' +
        'New text:\n' + sanitizedText,
        { resolved: true }
      );
      
      if (!acceptUpdate.success) {
        // If comment update fails, restore original text
        location.element.asText().deleteText(location.startOffset, location.endOffset);
        location.element.asText().insertText(location.startOffset, originalText);
        throw new Error("Failed to mark comment as accepted");
      }
      
      return true;
      
    } else {
      // For rejections, simply update the comment
      const rejectUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.REJECTED + '\n\nChanges rejected.\n\n' +
        'You can edit the comment and try again.',
        { resolved: false }
      );
      
      if (!rejectUpdate.success) {
        throw new Error("Failed to mark comment as rejected");
      }
      
      return true;
    }
    
  } catch (e) {
    // Only restore text if we made changes and have location info
    if (location && accepted) {
      try {
        location.element.asText().deleteText(location.startOffset, location.endOffset);
        location.element.asText().insertText(location.startOffset, originalText);
      } catch (restoreError) {
        e.message += "\nAdditionally, failed to restore original text: " + restoreError.message;
      }
    }
    
    throw new Error(
      "Failed to apply changes: " + e.message + "\n\n" +
      "Please try again. If the problem persists, " +
      "verify the text hasn't been modified and the comment still exists."
    );
  }
}

/**
 * Get the current document's ID
 * 
 * @return {String} The document ID
 */
function getDocumentId() {
  return DocumentApp.getActiveDocument().getId();
}

/**
 * Fetch available models from Ollama API
 * 
 * @return {Array} List of available model names
 */
function getOllamaModels() {
  try {
    // This function will be called from the client side since
    // Google Apps Script can't directly call external APIs
    return [];
  } catch (e) {
    console.error('Error fetching Ollama models:', e);
    return [];
  }
}

/**
 * Process AI edit with Ollama
 * This is a placeholder since actual API calls will be made client-side
 * 
 * @param {String} text - Text to process
 * @param {String} instruction - AI instruction
 * @param {String} model - Ollama model to use
 * @return {Object} Status object
 */
function processWithOllama(text, instruction, model) {
  // This function exists as a placeholder
  // Actual API calls will be made from the client side
  return {
    success: false,
    message: 'This function should be called from client side'
  };
}

/**
 * Debug test function to verify text location and replacement
 * 
 * @return {Object} Debug test results
 */
function debugTestFirstComment() {
  // Use CacheService for log persistence
  const cache = CacheService.getScriptCache();
  cache.remove('debugLogs'); // Clear previous logs

  const logs = [];
  function log(message, data = null) {
    const entry = {
      timestamp: new Date().toISOString(),
      message: message,
      data: data
    };
    logs.push(entry);
    Logger.log(JSON.stringify(entry)); // Keep Logger.log for Execution Logs
  }

  let location = null;
  let originalText = null;
  let commentState = null;
  
  try {
    log("Starting debug test");
    
    // Get document and file ID
    const doc = DocumentApp.getActiveDocument();
    const fileId = doc.getId();
    
    // Get comments
    log("Fetching comments");
    const comments = Drive.Comments.list(fileId, {
      fields: 'comments(id,content,quotedFileContent,anchor,resolved)',
      maxResults: 100
    }).comments || [];
    
    log("Retrieved comments", { count: comments.length });
    
    // Filter AI comments and log each one
    log("Examining all comments");
    const aiComments = comments.filter(comment => {
      const isAI = comment.content.trim().toLowerCase().startsWith('ai:');
      const isResolved = comment.resolved;
      const hasQuoted = !!comment.quotedFileContent;
      log("Comment found", {
        id: comment.id,
        isAI: isAI,
        isResolved: isResolved,
        hasQuoted: hasQuoted,
        content: comment.content.trim(),
        quotedText: comment.quotedFileContent ? comment.quotedFileContent.value : null,
        anchor: comment.anchor,
        resolved: comment.resolved
      });
      return isAI && !isResolved && hasQuoted;
    });
    
    log("Filtered AI comments", { count: aiComments.length });
    
    if (aiComments.length === 0) {
      throw new Error("No unprocessed AI comments found");
    }
    
    // Get first comment
    const comment = aiComments[0];
    log("Selected first comment for testing", {
      id: comment.id,
      content: comment.content,
      hasQuotedText: !!comment.quotedFileContent,
      quotedText: comment.quotedFileContent.value,
      anchor: comment.anchor,
      resolved: comment.resolved
    });
    
    // Store original state
    commentState = comment.resolved;
    originalText = comment.quotedFileContent.value;
    
    // Validate comment ID
    log("Validating comment ID");
    if (!validateCommentId(fileId, comment.id)) {
      throw new Error("Invalid comment ID");
    }
    
    // Get initial text location
    log("Getting initial text location");
    location = verifyTextLocation(fileId, comment.id, originalText);
    
    log("Initial text location result", location);
    
    if (!location) {
      throw new Error("Could not verify initial text location");
    }
    
    // Replace text with test message
    log("Attempting text replacement");
    const testText = "It works!!";
    
    // Replace the text
    log("Replacing text", {
      originalText: originalText,
      testText: testText,
      startOffset: location.startOffset,
      endOffset: location.endOffset
    });
    
    location.element.asText().deleteText(location.startOffset, location.endOffset);
    location.element.asText().insertText(location.startOffset, testText);

    // Re-find the location of the test text
    log("Finding new location after replacement");
    const newLocation = verifyTextLocation(fileId, comment.id, testText);
    
    if (!newLocation) {
      throw new Error("Could not verify location after replacement");
    }
    
    log("New location after replacement", newLocation);

    // Pause briefly to make the change visible
    Utilities.sleep(2000);

    try {
      // Restore original text using new location
      log("Restoring original text", {
        startOffset: newLocation.startOffset,
        endOffset: newLocation.endOffset,
        originalText: originalText
      });
      
      newLocation.element.asText().deleteText(newLocation.startOffset, newLocation.endOffset);
      newLocation.element.asText().insertText(newLocation.startOffset, originalText);

      // Verify restoration
      const finalLocation = verifyTextLocation(fileId, comment.id, originalText);
      if (!finalLocation) {
        throw new Error("Could not verify final text restoration");
      }

      const restoredText = finalLocation.element.asText().getText()
        .substring(finalLocation.startOffset, finalLocation.startOffset + originalText.length);
      
      log("Restoration verification", {
        expected: originalText,
        actual: restoredText,
        matches: restoredText === originalText
      });

      if (restoredText !== originalText) {
        throw new Error("Failed to restore original text exactly");
      }
    } catch (restoreError) {
      log("Error during text restoration", {
        error: restoreError.message,
        stack: restoreError.stack
      });
      throw new Error("Failed to restore original text: " + restoreError.message);
    }

    log("Debug test completed successfully");
    cache.put('debugLogs', JSON.stringify(logs), 600); // Store logs, expire in 10 minutes
    return { success: true }; // Return simple success, logs are in cache
    
  } catch (error) {
    log("Debug test failed", {
      error: error.message,
      stack: error.stack
    });
    cache.put('debugLogs', JSON.stringify(logs), 600);
    throw new Error(JSON.stringify({
      success: false,
      error: error.message // Only include message in top-level error
    }));
  }
}

/**
 * Get the debug logs
 * 
 * @return {String} The debug logs
 */
function getDebugLogs() {
  const cache = CacheService.getScriptCache();
  const logs = cache.get('debugLogs');
  cache.remove('debugLogs'); // Clear the cache after retrieval
  return logs || '[]'; // Return empty array string if no logs
}