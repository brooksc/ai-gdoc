/**
 * Mon Feb 24 14:36:52 PST 2025
 */
/**
 * AI Editor for Google Docs
 * This script allows users to highlight text, add "AI: [instruction]" comments,
 * and process them using local Ollama models.
 */

// Global logging configuration
const LOG_CONFIG = {
  CATEGORIES: {
    COMMENT: 'comment',
    TEXT: 'text',
    ERROR: 'error',
    DEBUG: 'debug',
    STATE: 'state',
    NUX: 'nux'
  },
  MAX_LOG_SIZE: 50000 // Maximum number of characters to keep in log
};

/**
 * Enhanced logging function that ensures consistent log format
 * @param {string} category - The category of the log
 * @param {string} message - The main log message
 * @param {Object} [data] - Optional data to include in the log
 */
function logDebug(category, message, data = null) {
  const logEntry = {
    timestamp: new Date().toISOString(),
    category: category,
    message: message,
    data: data,
    documentId: DocumentApp.getActiveDocument().getId()
  };

  // Use JSON.stringify with a replacer function to handle circular references
  const safeStringify = (obj) => {
    const seen = new Set();
    return JSON.stringify(obj, (key, value) => {
      if (typeof value === 'object' && value !== null) {
        if (seen.has(value)) {
          return '[Circular Reference]';
        }
        seen.add(value);
      }
      return value;
    }, 2);
  };

  // Create the log string
  let logString;
  try {
    logString = safeStringify(logEntry);
  } catch (e) {
    // If JSON stringification fails, create a simpler log
    logString = JSON.stringify({
      timestamp: logEntry.timestamp,
      category: logEntry.category,
      message: logEntry.message,
      error: "Failed to stringify data: " + e.message
    });
  }

  // Log to Logger service (for sidebar display)
  Logger.log(logString);
  
  // Also log to console for debugging in script editor
  console.log(logString);
  
  // Trim logs if they get too long
  try {
    const logs = Logger.getLog();
    if (logs && logs.length > LOG_CONFIG.MAX_LOG_SIZE) {
      // Clear the log and add a truncation message
      Logger.clear();
      Logger.log(JSON.stringify({
        timestamp: new Date().toISOString(),
        category: LOG_CONFIG.CATEGORIES.DEBUG,
        message: "Logs were truncated due to size limit"
      }));
      // Re-add the current log entry
      Logger.log(logString);
    }
  } catch (e) {
    console.error("Error managing log size:", e);
  }
}



// Comment state markers
const COMMENT_STATE = {
  ACCEPTED: "[STATE:ACCEPTED]",
  PROCESSING: "[STATE:PROCESSING]",
  REJECTED: "[STATE:REJECTED]"
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
 */
function sleep(ms) {
  Utilities.sleep(ms);
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
  
  // Validate input parameters
  if (!fileId || !commentId || !content) {
    Logger.log("aiedit-debug: Invalid parameters for comment update", {
      hasFileId: !!fileId,
      hasCommentId: !!commentId,
      hasContent: !!content
    });
    return { 
      success: false, 
      error: new Error("Invalid parameters"),
      details: { fileId, commentId, contentLength: content ? content.length : 0 }
    };
  }

  Logger.log("aiedit-debug: Starting comment update with full details", {
    fileId: fileId,
    commentId: commentId,
    contentLength: content.length,
    content: content.substring(0, Math.min(100, content.length)) + (content.length > 100 ? "..." : ""),
    options: JSON.stringify(options),
    maxAttempts: RETRY_CONFIG.MAX_ATTEMPTS,
    initialDelay: RETRY_CONFIG.INITIAL_DELAY_MS,
    maxDelay: RETRY_CONFIG.MAX_DELAY_MS
  });
  
  var attempt;
  for (attempt = 0; attempt < RETRY_CONFIG.MAX_ATTEMPTS; attempt += 1) {
    try {
      Logger.log("aiedit-debug: Starting update attempt", {
        attemptNumber: attempt + 1,
        totalAttempts: RETRY_CONFIG.MAX_ATTEMPTS
      });
      
      // Get the current comment first to check for existence and state
      let comment;
      try {
        comment = Drive.Comments.get(fileId, commentId, { 
          fields: "id,content,resolved,anchor,quotedFileContent"
        });
        
        // Safe substring operation on content
        let contentPreview = '';
        if (comment && comment.content) {
          contentPreview = comment.content.substring(0, Math.min(100, comment.content.length)) + 
                          (comment.content.length > 100 ? '...' : '');
        }
        
        Logger.log("aiedit-debug: Retrieved current comment state", {
          commentExists: !!comment,
          commentId: comment ? comment.id : null,
          resolved: comment ? comment.resolved : null,
          status: comment ? comment.status : null,
          hasContent: comment ? !!comment.content : false,
          contentLength: comment && comment.content ? comment.content.length : 0,
          hasQuotedContent: comment ? !!comment.quotedFileContent : false,
          hasAnchor: comment ? !!comment.anchor : false,
          hasReplies: comment ? (comment.replies || []).length : 0,
          contentPreview: contentPreview
        });
        
        if (!comment) {
          throw new Error('Comment not found');
        }
        
        if (!comment.content && options.resolved) {
          throw new Error('Cannot resolve comment with no content');
        }
        
      } catch (getError) {
        Logger.log("aiedit-debug: Failed to get comment", {
          error: getError.toString(),
          errorName: getError.name,
          errorStack: getError.stack,
          errorMessage: getError.message,
          commentId: commentId
        });
        throw new Error("Failed to get comment: " + getError.message);
      }
      
      if (!comment) {
        Logger.log("aiedit-debug: Comment not found condition triggered", {
          commentId: commentId,
          fileId: fileId
        });
        throw new Error("Comment not found");
      }
      
      // Validate the update content
      if (!content || typeof content !== "string") {
        throw new Error("Invalid content for comment update");
      }

      // Create update object with validated fields
      const updateObject = {
        content: content
      };
      
      // Store pre-update state for verification
      const preUpdateState = {
        content: comment.content,
        resolved: comment.resolved,
        status: comment.status
      };
      
      Logger.log("aiedit-debug: Preparing comment update", {
        updateObject: JSON.stringify(updateObject),
        preUpdateState: JSON.stringify(preUpdateState),
        attempt: attempt + 1,
        contentLength: content.length,
        willResolve: options.resolved === true,
        commentId: commentId
      });
      
      // Perform the update with proper fields parameter
      let updatedComment;
      try {
        Logger.log("aiedit-debug: Attempting Drive.Comments.update call");
        updatedComment = Drive.Comments.update(
          updateObject,
          fileId,
          commentId,
          { fields: 'id,content,resolved' }
        );
        Logger.log("aiedit-debug: Drive.Comments.update call completed", {
          success: !!updatedComment,
          updatedCommentObject: JSON.stringify(updatedComment)
        });
      } catch (updateError) {
        Logger.log("aiedit-debug: Drive.Comments.update call failed", {
          error: updateError.toString(),
          errorName: updateError.name,
          errorStack: updateError.stack,
          errorMessage: updateError.message,
          updateObject: JSON.stringify(updateObject)
        });
        throw new Error("Update API call failed: " + updateError.message);
      }
      
      Logger.log("aiedit-debug: Verifying update result", {
        hasUpdatedComment: !!updatedComment,
        updatedCommentId: updatedComment ? updatedComment.id : null,
        updatedResolved: updatedComment ? updatedComment.resolved : null,
        expectedResolved: options.resolved === true,
        contentMatch: updatedComment ? (updatedComment.content === content) : false,
        actualContent: updatedComment ? updatedComment.content.substring(0, 100) + "..." : null,
        expectedContent: content.substring(0, 100) + "..."
      });
      
      // Verify the update was successful with detailed checks
      if (!updatedComment) {
        Logger.log("aiedit-debug: No response from update call", {
          attempt: attempt + 1,
          commentId: commentId
        });
        throw new Error("Update returned no response");
      }
      
      // Verify content update
      const contentMatch = updatedComment.content === content;
      const resolvedMatch = updatedComment.resolved === options.resolved;
      
      Logger.log("aiedit-debug: Update verification", {
        contentMatches: contentMatch,
        resolvedMatches: resolvedMatch,
        expectedContent: content.substring(0, Math.min(100, content.length)) + (content.length > 100 ? '...' : ''),
        actualContent: updatedComment.content.substring(0, Math.min(100, updatedComment.content.length)) + 
                      (updatedComment.content.length > 100 ? '...' : ''),
        expectedResolved: options.resolved,
        actualResolved: updatedComment.resolved,
        commentId: commentId
      });
      
      if (!contentMatch || !resolvedMatch) {
        throw new Error(
          `Update verification failed: ` +
          `${!contentMatch ? 'Content mismatch ' : ''}` +
          `${!resolvedMatch ? 'Resolved state mismatch' : ''}`
        );
      }
      
      if (updatedComment.resolved !== options.resolved) {
        Logger.log("aiedit-debug: Resolved state mismatch", {
          expectedResolved: options.resolved,
          actualResolved: updatedComment.resolved
        });
        throw new Error("Resolved state mismatch after update");
      }
      
      Logger.log("aiedit-debug: Comment update successful", {
        commentId: updatedComment.id,
        resolved: updatedComment.resolved,
        contentLength: updatedComment.content.length
      });
      return { success: true };
      
    } catch (error) {
      lastError = error;
      Logger.log("aiedit-debug: Comment update attempt failed", {
        attempt: attempt + 1,
        error: error.toString(),
        errorName: error.name,
        errorStack: error.stack,
        errorMessage: error.message,
        isLastAttempt: attempt === RETRY_CONFIG.MAX_ATTEMPTS - 1
      });
      
      if (attempt < RETRY_CONFIG.MAX_ATTEMPTS - 1) {
        const delay = calculateBackoffDelay(attempt);
        Logger.log("aiedit-debug: Will retry after delay", {
          delayMs: delay,
          nextAttemptNumber: attempt + 2,
          remainingAttempts: RETRY_CONFIG.MAX_ATTEMPTS - (attempt + 1)
        });
        sleep(delay);
      }
    }
  }
  
  Logger.log("aiedit-debug: All comment update attempts failed", {
    totalAttempts: RETRY_CONFIG.MAX_ATTEMPTS,
    finalError: lastError ? lastError.toString() : null,
    errorDetails: lastError ? {
      name: lastError.name,
      message: lastError.message,
      stack: lastError.stack
    } : null
  });
  
  return {
    success: false,
    error: lastError,
    details: {
      attempts: RETRY_CONFIG.MAX_ATTEMPTS,
      fileId: fileId,
      commentId: commentId,
      contentLength: content.length,
      wantedResolved: options.resolved === true
    }
  };
}

/**
 * Creates the AI Editor menu and opens the sidebar when the document is opened
 */
function onOpen() {
  // Create the menu
  DocumentApp.getUi()
    .createMenu("AI Editor")
    .addItem("Open Editor", "showSidebar")
    .addToUi();
    
  // Automatically open the sidebar
  showSidebar();
}

/**
 * Displays the AI Editor sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar.html")
    .setTitle("AI Editor")
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
  // Create RegExp objects instead of literal regex with control characters
  const controlCharsRegex = new RegExp('[\\u0000-\\u001F\\u007F-\\u009F]', 'g');
  const lineSeparatorsRegex = new RegExp('\\u2028|\\u2029', 'g');
  
  return text
    .replace(controlCharsRegex, "") // Remove control characters
    .replace(lineSeparatorsRegex, "\n") // Normalize line separators
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
      pageSize: 100,
      includeDeleted: false
    });
    
    const comments = response.comments || [];
    Logger.log("aiedit: Retrieved " + comments.length + " total comments");
    
    // Filter for AI comments that are either unprocessed or rejected
    const aiComments = comments.filter(comment => {
      const content = comment.content.trim();
      const isAIComment = content.startsWith("AI:");
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
      state: getCommentState(comment) || "unprocessed"
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
    Logger.log("aiedit-debug: Invalid comment ID format", {
      commentId: commentId,
      type: typeof commentId
    });
    return false;
  }
  
  try {
    const comment = Drive.Comments.get(fileId, commentId, {
      fields: 'id'
    });
    Logger.log("aiedit-debug: Comment validation result", {
      commentId: commentId,
      exists: !!comment,
      matches: comment ? (comment.id === commentId) : false
    });
    return comment && comment.id === commentId;
  } catch (e) {
    Logger.log("aiedit-debug: Error validating comment ID", {
      commentId: commentId,
      error: e.toString(),
      errorName: e.name,
      errorStack: e.stack
    });
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
      fields: "id,anchor,quotedFileContent,content"
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
    const controlCharsRegex = new RegExp('[\\u0000-\\u001F\\u007F-\\u009F]', 'g');
    const lineSeparatorsRegex = new RegExp('\\u2028|\\u2029', 'g');
    const whitespaceRegex = new RegExp('\\s+', 'g');
    
    const normalizedOriginal = originalText
      .replace(controlCharsRegex, "")
      .replace(lineSeparatorsRegex, "\n")
      .replace(whitespaceRegex, " ")
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
    let bestScore = 0;
    
    for (const match of foundElements) {
      const elementText = match.element.getText();
      const contextStart = Math.max(0, match.startOffset - 50);
      const contextEnd = Math.min(elementText.length, match.endOffset + 51);
      const context = elementText.substring(contextStart, contextEnd);
      
      const contextSimilarity = calculateSimilarity(context, quotedContext);
      
      if (contextSimilarity > bestScore && contextSimilarity > 0.7) {
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
    sleep(5000);
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
  // Get active document - needed for text operations later
  const doc = DocumentApp.getActiveDocument();
  let originalText = null;
  let location = null;
  
  Logger.log("aiedit-debug: Starting applyAIEdit", {
    fileId: fileId,
    commentId: commentId,
    suggestedTextLength: suggestedText ? suggestedText.length : 0,
    accepted: accepted,
    documentId: doc.getId() // Use doc to log the document ID
  });
  
  try {
    // Validate inputs
    if (!fileId || !commentId || !suggestedText) {
      Logger.log("aiedit-debug: Missing required parameters", {
        hasFileId: !!fileId,
        hasCommentId: !!commentId,
        hasSuggestedText: !!suggestedText
      });
      throw new Error("Missing required parameters for applying AI edit");
    }
    
    if (!validateCommentId(fileId, commentId)) {
      Logger.log("aiedit-debug: Comment validation failed");
      throw new Error("The comment no longer exists or is inaccessible");
    }
    
    // Get the comment and verify it exists with full details
    let comment;
    try {
      // First log the API request we're about to make
      Logger.log('aiedit-debug: Requesting comment from Drive API', {
        fileId: fileId,
        commentId: commentId,
        requestedFields: 'id,quotedFileContent,anchor,resolved,content,createdTime,modifiedTime,replies'
      });

      comment = Drive.Comments.get(fileId, commentId, {
        fields: "id,content,resolved,anchor,quotedFileContent,replies",
        includeDeleted: false
      });
      
      // Log the raw API response
      Logger.log("aiedit-debug: Raw Drive API response", {
        response: JSON.stringify(comment, null, 2)
      });

      // Log parsed details
      Logger.log("aiedit-debug: Parsed comment details", {
        commentId: commentId,
        exists: !!comment,
        hasQuotedContent: comment ? !!comment.quotedFileContent : false,
        quotedContentLength: comment && comment.quotedFileContent ? comment.quotedFileContent.value.length : 0,
        resolved: comment ? comment.resolved : null,
        createdTime: comment ? comment.createdTime : null,
        modifiedTime: comment ? comment.modifiedTime : null,
        hasAnchor: comment ? !!comment.anchor : false,
        hasReplies: comment ? (comment.replies && comment.replies.length > 0) : false,
        commentKeys: comment ? Object.keys(comment) : [],
        content: comment ? comment.content : null
      });
    } catch (getError) {
      Logger.log("aiedit-debug: Failed to get comment details", {
        error: getError.toString(),
        stack: getError.stack,
        commentId: commentId
      });
      throw new Error("Failed to get comment details: " + getError.message);
    }
    
    Logger.log("aiedit-debug: Retrieved comment", {
      exists: !!comment,
      hasQuotedContent: comment ? !!comment.quotedFileContent : false,
      resolved: comment ? comment.resolved : null
    });
    
    if (!comment) {
      Logger.log("aiedit-debug: Comment not found", { commentId: commentId });
      throw new Error("Could not find the specified comment");
    }
    
    if (comment.resolved) {
      Logger.log("aiedit-debug: Comment already resolved", {
        commentId: commentId,
        status: comment.status,
        resolved: comment.resolved
      });
      throw new Error("This comment has already been resolved");
    }
    
    // Validate comment structure
    if (!comment.content) {
      Logger.log("aiedit-debug: Comment missing content", { commentId: commentId });
      throw new Error("Comment is missing content");
    }
    
    if (!comment.quotedFileContent || !comment.quotedFileContent.value) {
      throw new Error("The comment is missing the original text selection");
    }
    
    // Store original text
    originalText = comment.quotedFileContent.value;
    
    // Verify text location before proceeding
    Logger.log("aiedit-debug: Verifying text location", {
      originalTextLength: originalText.length
    });
    
    const verifyResult = verifyTextLocation(fileId, commentId, originalText);
    Logger.log("aiedit-debug: Text location verification result", verifyResult);
    
    if (!verifyResult.success) {
      throw new Error(verifyResult.error || "Could not verify text location");
    }
    location = verifyResult.location;
    
    // Prepare sanitized text
    const sanitizedText = sanitizeText(suggestedText);
    Logger.log("aiedit-debug: Sanitized suggested text", {
      originalLength: suggestedText.length,
      sanitizedLength: sanitizedText.length,
      changed: sanitizedText !== suggestedText
    });
    
    if (!sanitizedText) {
      throw new Error("The suggested text is empty after sanitization");
    }
    
    if (accepted) {
      try {
        // Get the text element and verify indices
        const textElement = location.element.asText();
        const elementLength = textElement.getText().length;
        
        Logger.log("aiedit-debug: Text element verification", {
          elementLength: elementLength,
          startOffset: location.startOffset,
          endOffset: location.endOffset,
          textToReplaceLength: location.endOffset - location.startOffset + 1
        });
        
        // Validate indices are within bounds
        if (location.startOffset < 0 || location.endOffset >= elementLength) {
          throw new Error(`Invalid text range: start=${location.startOffset}, end=${location.endOffset}, length=${elementLength}`);
        }
        
        Logger.log("aiedit-debug: Attempting text replacement");
        
        // Replace text in a single operation to maintain consistency
        Logger.log('aiedit-debug: Text replacement details', {
          startOffset: location.startOffset,
          endOffset: location.endOffset,
          originalText: textElement.getText().substring(location.startOffset, location.endOffset + 1),
          sanitizedText: sanitizedText
        });
        // Use deleteText + insertText instead of replaceText since the API is finicky
        Logger.log('aiedit-debug: Using deleteText + insertText', {
          startOffset: location.startOffset,
          endOffset: location.endOffset,
          replacement: sanitizedText
        });
        textElement.deleteText(location.startOffset, location.endOffset);
        textElement.insertText(location.startOffset, sanitizedText);
        
        // Verify the replacement
        const currentText = textElement.getText();
        const verifyText = currentText.substring(location.startOffset, location.startOffset + sanitizedText.length);
        Logger.log('aiedit-debug: Replacement verification details', {
          currentTextLength: currentText.length,
          verifyStartOffset: location.startOffset,
          verifyEndOffset: location.startOffset + sanitizedText.length,
          verifyTextLength: verifyText.length,
          expectedLength: sanitizedText.length
        });
        
        Logger.log("aiedit-debug: Text replacement verification", {
          expectedLength: sanitizedText.length,
          actualLength: verifyText.length,
          matches: verifyText === sanitizedText,
          verifyText: verifyText,
          sanitizedText: sanitizedText
        });
        
        // Compare exact lengths first, then content
        if (verifyText.length !== sanitizedText.length || verifyText !== sanitizedText) {
          Logger.log("aiedit-debug: Text replacement verification failed", {
            expectedLength: sanitizedText.length,
            actualLength: verifyText.length,
            expectedText: sanitizedText,
            actualText: verifyText,
            startOffset: location.startOffset,
            endOffset: location.endOffset
          });
          
          // If verification fails, attempt to restore original text
          try {
            const currentText = textElement.getText();
            Logger.log("aiedit-debug: Current text state before restoration", {
              fullTextLength: currentText.length,
              modifiedSection: currentText.substring(
                Math.max(0, location.startOffset - 10),
                Math.min(currentText.length, location.startOffset + sanitizedText.length + 10)
              )
            });
            
            // Check if text was actually modified
            const modifiedText = currentText.substring(location.startOffset, location.startOffset + sanitizedText.length);
            if (modifiedText === sanitizedText) {
              Logger.log("aiedit-debug: Restoring original text");
              // Use deleteText + insertText instead of replaceText
              textElement.deleteText(location.startOffset, location.startOffset + sanitizedText.length - 1);
              textElement.insertText(location.startOffset, originalText);
              
              // Verify restoration
              const restoredText = textElement.getText().substring(location.startOffset, location.startOffset + originalText.length);
              if (restoredText !== originalText) {
                throw new Error("Failed to restore original text");
              }
            } else {
              Logger.log("aiedit-debug: Text was not modified as expected", {
                expectedModification: sanitizedText,
                actualText: modifiedText
              });
            }

             } catch (restoreError) {
                Logger.log("aiedit-debug: Error during text restoration attempt", {
                    error: restoreError.message,
                    stack: restoreError.stack
                });
             }
          throw new Error("Failed to verify text replacement");
        }
        


        // Update comment to mark as accepted
        Logger.log("aiedit-debug: Attempting to mark comment as accepted", {
          fileId: fileId,
          commentId: commentId,
          textLength: sanitizedText.length,
          resolved: true,
          originalCommentContent: comment ? comment.content : null
        });
        
        // First update the comment content
        const acceptUpdate = retryCommentUpdate(
          fileId,
          commentId,
          COMMENT_STATE.ACCEPTED + '\n\nChanges applied successfully:\n\n' +
          'Original text:\n' + originalText + '\n\n' +
          'New text:\n' + sanitizedText
        );

        // Then create a resolving reply
        if (acceptUpdate.success) {
          try {
            Logger.log("aiedit-debug: Creating resolving reply");
            // First check if comment is already resolved
            const currentComment = Drive.Comments.get(fileId, commentId, {
              fields: 'id,resolved'
            });
            
            if (currentComment.resolved) {
              Logger.log("aiedit-debug: Comment already resolved", {
                commentId: commentId
              });
              return; // Already resolved, no need to do it again
            }
            
            const reply = Drive.Replies.create(
              {
                action: 'resolve',
                content: 'Accepted AI suggestion'
              },
              fileId,
              commentId,
              { fields: 'id,action,resolved' }
            );
            
            // Verify the comment was actually resolved
            const verifyComment = Drive.Comments.get(fileId, commentId, {
              fields: 'id,resolved'
            });
            
            Logger.log("aiedit-debug: Created resolving reply", {
              replyId: reply.id,
              action: reply.action,
              commentResolved: verifyComment.resolved
            });
            
            if (!verifyComment.resolved) {
              throw new Error('Failed to resolve comment despite successful reply creation');
            }
          } catch (error) {
            Logger.log("aiedit-debug: Failed to create resolving reply", {
              error: error.toString()
            });
            // Don't throw here - the text was updated successfully
            // Just log that we couldn't resolve the comment
          }
        }
        
        if (!acceptUpdate.success) {
          Logger.log("aiedit-debug: Failed to mark comment as accepted", {
            error: acceptUpdate.error,
            details: acceptUpdate.details,
            fileId: fileId,
            commentId: commentId
          });
          
          // If comment update fails, restore original text
          const currentText = textElement.getText();
            //check and see if it's equal first.
            if (currentText.substring(location.startOffset, location.startOffset + sanitizedText.length) === sanitizedText)            {
              textElement.deleteText(location.startOffset, location.startOffset + sanitizedText.length - 1);
              textElement.insertText(location.startOffset, originalText);
          }
          throw new Error("Failed to mark comment as accepted: " + 
            (acceptUpdate.error ? acceptUpdate.error.message : "Unknown error") +
            "\nDetails: " + JSON.stringify(acceptUpdate.details));
        }
        
        Logger.log("aiedit-debug: Successfully marked comment as accepted");
        return true;
      } catch (error) {
        Logger.log("aiedit-debug: Error during text replacement", {
          error: error.toString(),
          stack: error.stack
        });
        
        // If anything fails during text replacement, restore original text *only if* it was changed.
        try {
          const textElement = location.element.asText();
          const currentText = textElement.getText();

          // Check if text was actually modified using exact length comparison
          const modifiedText = currentText.substring(location.startOffset, location.startOffset + sanitizedText.length);
          
          Logger.log("aiedit-debug: Checking for text restoration after error", {
            modifiedTextLength: modifiedText.length,
            expectedLength: sanitizedText.length,
            matches: modifiedText === sanitizedText
          });
          
          // Log the current state for debugging
          Logger.log("aiedit-debug: Text state before restoration", {
            currentText: currentText.substring(Math.max(0, location.startOffset - 10), 
                                              Math.min(currentText.length, location.startOffset + sanitizedText.length + 10)),
            modifiedTextActual: modifiedText,
            sanitizedTextExpected: sanitizedText,
            startOffset: location.startOffset,
            originalTextLength: originalText.length
          });

          if (modifiedText.length === sanitizedText.length && modifiedText === sanitizedText) {
            Logger.log("aiedit-debug: Restoring text after error");
            // Text *was* modified, so restore the original in a single operation
            textElement.replaceText(location.startOffset, 
                                  location.startOffset + sanitizedText.length - 1,
                                  originalText);
            
            // Verify restoration
            const restoredText = textElement.getText().substring(location.startOffset,
                                                               location.startOffset + originalText.length);
            Logger.log("aiedit-debug: Text restoration verification", {
              restoredText: restoredText,
              originalText: originalText,
              matches: restoredText === originalText
            });
            
            if (restoredText !== originalText) {
              throw new Error("Failed to restore original text after error");
            }
          } else {
            Logger.log("aiedit-debug: No restoration needed - text was not modified");
          }

        } catch (restoreError) {
          Logger.log("aiedit-debug: Error during text restoration", {
            error: restoreError.toString(),
            stack: restoreError.stack
          });
          error.message += "\nAdditionally, failed to restore original text: " + restoreError.message;
        }
        throw error;
      }
    } else {
      // For rejections, simply update the comment
      Logger.log("aiedit-debug: Attempting to mark comment as rejected");
      const rejectUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.REJECTED + '\n\nChanges rejected.\n\nYou can edit the comment and try again.',
        { resolved: false }
      );
      
      if (!rejectUpdate.success) {
        Logger.log("aiedit-debug: Failed to mark comment as rejected", {
          error: rejectUpdate.error
        });
        throw new Error("Failed to mark comment as rejected");
      }
      
      Logger.log("aiedit-debug: Successfully marked comment as rejected");
      return true;
    }
    
  } catch (e) {
    Logger.log("aiedit-debug: Final error in applyAIEdit", {
      error: e.toString(),
      stack: e.stack
    });
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
  // This function is a placeholder for client-side usage
  // In real usage, it would make API calls from JavaScript in the HTML
  return [];
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
  // These parameters are intentionally unused in this server-side version
  // as the actual processing happens client-side
  Logger.log("aiedit-debug: processWithOllama called (server-side placeholder)", {
    textLength: text ? text.length : 0,
    hasInstruction: !!instruction,
    model: model || "unknown"
  });
  
  return {
    success: false,
    message: 'This function should be called from client side'
  };
}

/**
 * Process text with Gemini API
 * 
 * @param {String} prompt - Full prompt to send to Gemini
 * @return {Object} Response object with Gemini API response
 */
function processWithGemini(prompt) {
  try {
    // Get the API key
    const apiKey = getGeminiApiKey();
    
    if (!apiKey) {
      throw new Error("No Gemini API key found. Please add your API key in the settings.");
    }
    
    Logger.log("aiedit-debug: Processing with Gemini API", {
      promptLength: prompt.length
    });
    
    // Set up the API request
    const apiUrl = 'https://api.gemini.google.com/v1beta/models/gemini-2.0-flash:generate';
    const requestOptions = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${apiKey}`
      },
      payload: JSON.stringify({
        contents: [
          {
            parts: [
              { text: prompt }
            ]
          }
        ],
        generationConfig: {
          maxOutputTokens: 8192,
          temperature: 0.7
        }
      }),
      muteHttpExceptions: true
    };
    
    // Send the request
    const response = UrlFetchApp.fetch(apiUrl, requestOptions);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log("aiedit-debug: Received Gemini API response", {
      responseCode: responseCode,
      responseLength: responseText.length
    });
    
    if (responseCode !== 200) {
      throw new Error(`Gemini API request failed with status ${responseCode}: ${responseText}`);
    }
    
    // Parse the response
    const responseJson = JSON.parse(responseText);
    
    // Extracting the response based on actual Gemini API response format
    // Format: { candidates: [{ content: { parts: [{ text: "..." }] } }] }
    if (!responseJson.candidates || 
        !responseJson.candidates[0] || 
        !responseJson.candidates[0].content || 
        !responseJson.candidates[0].content.parts || 
        !responseJson.candidates[0].content.parts[0] || 
        !responseJson.candidates[0].content.parts[0].text) {
      
      // Log the actual response structure for debugging
      Logger.log("aiedit-debug: Unexpected Gemini API response format", {
        responsePreview: JSON.stringify(responseJson).substring(0, 500) + "..."
      });
      
      throw new Error("Unexpected Gemini API response format");
    }
    
    const aiResponse = responseJson.candidates[0].content.parts[0].text || "";
    
    return {
      success: true,
      response: aiResponse
    };
  } catch (e) {
    Logger.log("aiedit-debug: Error processing with Gemini API", {
      error: e.toString(),
      stack: e.stack
    });
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Apply a suggested change to the document
 * 
 * @param {Object} suggestion - Suggestion object with location and revised text
 * @return {Object} Result with success status
 */
function applySuggestedChange(suggestion) {
  try {
    if (!suggestion || !suggestion.location || !suggestion.revisedText) {
      throw new Error("Invalid suggestion data");
    }
    
    Logger.log("aiedit-debug: Applying suggested change", {
      originalTextLength: suggestion.originalText.length,
      revisedTextLength: suggestion.revisedText.length,
      startOffset: suggestion.location.startOffset,
      endOffset: suggestion.location.endOffset
    });
    
    const textElement = suggestion.location.element.asText();
    const elementLength = textElement.getText().length;
    
    // Validate indices are within bounds
    if (suggestion.location.startOffset < 0 || suggestion.location.endOffset >= elementLength) {
      throw new Error(`Invalid text range: start=${suggestion.location.startOffset}, end=${suggestion.location.endOffset}, length=${elementLength}`);
    }
    
    // Replace text in the document
    textElement.deleteText(suggestion.location.startOffset, suggestion.location.endOffset);
    textElement.insertText(suggestion.location.startOffset, suggestion.revisedText);
    
    // Verify the replacement
    const verifyText = textElement.getText().substring(
      suggestion.location.startOffset, 
      suggestion.location.startOffset + suggestion.revisedText.length
    );
    
    if (verifyText !== suggestion.revisedText) {
      // Try to restore original text if verification fails
      textElement.deleteText(
        suggestion.location.startOffset, 
        suggestion.location.startOffset + suggestion.revisedText.length - 1
      );
      textElement.insertText(suggestion.location.startOffset, suggestion.originalText);
      
      throw new Error("Failed to verify text replacement");
    }
    
    Logger.log("aiedit-debug: Successfully applied suggestion");
    return { success: true };
    
  } catch (e) {
    Logger.log("aiedit-debug: Error applying suggested change", {
      error: e.toString(),
      stack: e.stack
    });
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Save the Gemini API key to user properties
 * 
 * @param {String} apiKey - Gemini API key
 * @return {Boolean} Success status
 */
function saveGeminiApiKey(apiKey) {
  try {
    if (!apiKey) {
      Logger.log("aiedit-debug: No API key provided");
      return false;
    }
    
    Logger.log("aiedit-debug: Saving Gemini API key");
    
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('geminiApiKey', apiKey);
    
    return true;
  } catch (e) {
    Logger.log("aiedit-debug: Error saving Gemini API key", {
      error: e.toString(),
      stack: e.stack
    });
    return false;
  }
}

/**
 * Get the Gemini API key from user properties
 * 
 * @return {String} Gemini API key or empty string if not found
 */
function getGeminiApiKey() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const apiKey = userProperties.getProperty('geminiApiKey');
    
    Logger.log("aiedit-debug: Retrieving Gemini API key", {
      exists: !!apiKey
    });
    
    return apiKey || '';
  } catch (e) {
    Logger.log("aiedit-debug: Error retrieving Gemini API key", {
      error: e.toString(),
      stack: e.stack
    });
    return '';
  }
}

/**
 * Parses AI-generated text to extract suggested changes using the specified format
 * <suggestion>Original text<changeto/>Revised text</suggestion>
 * 
 * @param {string} aiResponse - The text response from the AI model
 * @return {Array} An array of objects with original and revised text
 */
function parseSuggestedChanges(aiResponse) {
  logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Parsing suggested changes from AI response", 
           { responseLength: aiResponse ? aiResponse.length : 0 });
  
  if (!aiResponse) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "No AI response to parse");
    return [];
  }
  
  const suggestionRegex = /<suggestion>([\s\S]*?)<changeto\/>([\s\S]*?)<\/suggestion>/g;
  const suggestions = [];
  let match;
  
  try {
    while ((match = suggestionRegex.exec(aiResponse)) !== null) {
      if (match.length === 3) {
        suggestions.push({
          original: match[1].trim(),
          revised: match[2].trim()
        });
      }
    }
    
    logDebug(LOG_CONFIG.CATEGORIES.DEBUG, `Found ${suggestions.length} suggestions in AI response`);
    return suggestions;
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error parsing suggested changes", { error: error.toString() });
    return [];
  }
}

/**
 * Locates each suggestion's position in the document text
 * 
 * @param {string} documentText - The full document text
 * @param {Array} suggestions - Array of suggestion objects with original and revised text
 * @return {Array} The suggestions with their locations in the document
 */
function findSuggestionLocations(documentText, suggestions) {
  logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Finding suggestion locations in document", 
           { docLength: documentText.length, suggestionCount: suggestions.length });
  
  if (!documentText || !suggestions || suggestions.length === 0) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Missing document text or suggestions for location finding");
    return [];
  }
  
  const locatedSuggestions = [];
  const unlocatedSuggestions = [];
  
  suggestions.forEach((suggestion, index) => {
    try {
      const exactIndex = documentText.indexOf(suggestion.original);
      
      if (exactIndex !== -1) {
        // Found exact match
        locatedSuggestions.push({
          ...suggestion,
          index: exactIndex,
          length: suggestion.original.length
        });
        return;
      }
      
      // No exact match found, try finding similar matches
      let bestMatch = null;
      let bestScore = 0;
      
      // Split document into chunks of similar size to the original text
      const chunkSize = Math.max(100, suggestion.original.length * 2);
      for (let i = 0; i < documentText.length - suggestion.original.length; i += chunkSize / 2) {
        const chunk = documentText.substr(i, chunkSize);
        const similarity = calculateSimilarity(chunk, suggestion.original);
        
        if (similarity > bestScore && similarity > 0.7) { // Only consider matches with high similarity
          bestScore = similarity;
          bestMatch = {
            index: i,
            length: chunk.length
          };
        }
      }
      
      if (bestMatch) {
        locatedSuggestions.push({
          ...suggestion,
          index: bestMatch.index,
          length: suggestion.original.length,
          fuzzyMatch: true,
          similarity: bestScore
        });
      } else {
        unlocatedSuggestions.push(suggestion);
      }
    } catch (error) {
      logDebug(LOG_CONFIG.CATEGORIES.ERROR, `Error locating suggestion ${index}`, { error: error.toString() });
      unlocatedSuggestions.push(suggestion);
    }
  });
  
  logDebug(LOG_CONFIG.CATEGORIES.DEBUG, `Located ${locatedSuggestions.length} suggestions, failed to locate ${unlocatedSuggestions.length}`);
  
  return {
    located: locatedSuggestions,
    unlocated: unlocatedSuggestions
  };
}

/**
 * Converts the active document to markdown format
 * 
 * @return {string} Document content in markdown format
 */
function getDocumentAsMarkdown() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const numElements = body.getNumChildren();
    let markdown = '';
    
    logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Converting document to markdown", { numElements });
    
    // Process each element in the document
    for (let i = 0; i < numElements; i++) {
      const element = body.getChild(i);
      const type = element.getType();
      
      switch (type) {
        case DocumentApp.ElementType.PARAGRAPH:
          markdown += processParagraphToMarkdown(element.asParagraph());
          break;
        case DocumentApp.ElementType.TABLE:
          markdown += processTableToMarkdown(element.asTable());
          break;
        case DocumentApp.ElementType.LIST_ITEM:
          markdown += processListItemToMarkdown(element.asListItem());
          break;
        case DocumentApp.ElementType.HORIZONTAL_RULE:
          markdown += '---\n\n';
          break;
        // Add other element types as needed
      }
    }
    
    logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Completed document to markdown conversion", 
             { markdownLength: markdown.length });
    
    return markdown;
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error converting document to markdown", 
             { error: error.toString() });
    return '';
  }
}

/**
 * Converts a paragraph element to markdown
 * 
 * @param {Paragraph} paragraph - The paragraph element
 * @return {string} Markdown representation of the paragraph
 */
function processParagraphToMarkdown(paragraph) {
  const text = paragraph.getText();
  const headingType = paragraph.getHeading();
  
  // Skip empty paragraphs
  if (!text || text.trim().length === 0) {
    return '\n';
  }
  
  // Handle headings
  switch (headingType) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return `# ${text}\n\n`;
    case DocumentApp.ParagraphHeading.HEADING2:
      return `## ${text}\n\n`;
    case DocumentApp.ParagraphHeading.HEADING3:
      return `### ${text}\n\n`;
    case DocumentApp.ParagraphHeading.HEADING4:
      return `#### ${text}\n\n`;
    case DocumentApp.ParagraphHeading.HEADING5:
      return `##### ${text}\n\n`;
    case DocumentApp.ParagraphHeading.HEADING6:
      return `###### ${text}\n\n`;
    default:
      // Process inline text formatting
      return `${processTextWithFormatting(paragraph)}\n\n`;
  }
}

/**
 * Converts a list item to markdown
 * 
 * @param {ListItem} listItem - The list item element
 * @return {string} Markdown representation of the list item
 */
function processListItemToMarkdown(listItem) {
  const text = listItem.getText();
  const glyphType = listItem.getGlyphType();
  const indentLevel = listItem.getNestingLevel();
  const indent = '  '.repeat(indentLevel);
  
  // Handle different list types
  if (glyphType === DocumentApp.GlyphType.NUMBER) {
    return `${indent}1. ${text}\n`;
  } else {
    return `${indent}* ${text}\n`;
  }
}

/**
 * Converts a table to markdown
 * 
 * @param {Table} table - The table element
 * @return {string} Markdown representation of the table
 */
function processTableToMarkdown(table) {
  let markdown = '';
  const numRows = table.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const row = table.getRow(i);
    const numCells = row.getNumCells();
    const cells = [];
    
    for (let j = 0; j < numCells; j++) {
      cells.push(row.getCell(j).getText().trim());
    }
    
    markdown += `| ${cells.join(' | ')} |\n`;
    
    // Add header separator row after first row
    if (i === 0) {
      markdown += `| ${cells.map(() => '---').join(' | ')} |\n`;
    }
  }
  
  return markdown + '\n';
}

/**
 * Processes text with inline formatting
 * 
 * @param {Paragraph} paragraph - The paragraph containing text
 * @return {string} Text with markdown formatting
 */
function processTextWithFormatting(paragraph) {
  const text = paragraph.getText();
  let markdown = text;
  
  // This is a simplified version - a full implementation would need to handle
  // overlapping formatting and other complexities
  const textObj = paragraph.editAsText();
  
  // Check for bold text
  for (let i = 0; i < text.length; i++) {
    const isBold = textObj.isBold(i);
    const isItalic = textObj.isItalic(i);
    
    // This is just a placeholder - real implementation would be more complex
    // to handle proper markdown conversion of formatting
  }
  
  // For simplicity, we're just returning the plain text
  // A complete implementation would handle all text formatting
  return markdown;
}

/**
 * Processes the document and generates inline suggestions based on user prompt
 * 
 * @param {string} prompt - The user prompt for processing
 * @param {string} modelName - The model to use for processing
 * @return {Object} Object with success status and suggestions
 */
function processDocumentForInlineSuggestions(prompt, modelName) {
  try {
    Logger.log("aiedit-debug: Processing document for inline suggestions", {
      promptLength: prompt.length,
      modelName: modelName
    });
    
    // Get document content as markdown
    const documentMarkdown = getDocumentAsMarkdown();
    
    // Prepare the API request with suggestion format instructions
    const formattingInstructions = `
You are an AI editor tasked with improving the clarity, grammar, and overall quality of the following document. 
For each identified improvement, output an inline suggestion using the following format:

<suggestion>Original text<changeto/>Revised text</suggestion>

Please adhere to these instructions:
1. Only output the suggestions where changes are needed.
2. Do not reproduce the entire document.
3. Make only minimal, targeted edits—do not alter parts of the document that don't need changes.
4. Each suggestion should contain the exact original text that needs to be replaced.
5. The revised text should maintain the same general meaning but improve clarity, grammar, or style.
`;

    // Create the full prompt
    const fullPrompt = `${formattingInstructions}\n\n${prompt}\n\nHere is the document content:\n\n${documentMarkdown}`;
    
    let aiResponse;
    
    // Check if using Gemini API (model name starts with "gemini")
    if (modelName.toLowerCase().startsWith("gemini")) {
      // Process with Gemini API
      const geminiResult = processWithGemini(fullPrompt);
      
      if (!geminiResult.success) {
        throw new Error("Gemini API processing failed: " + (geminiResult.error || "Unknown error"));
      }
      
      aiResponse = geminiResult.response;
    } else {
      // Process with Ollama API
      const apiUrl = 'http://localhost:11434/api/generate';
      const requestOptions = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          model: modelName,
          prompt: fullPrompt,
          stream: false
        }),
        muteHttpExceptions: true
      };
      
      Logger.log("aiedit-debug: Sending Ollama API request", {
        modelName: modelName,
        promptPreview: prompt.substring(0, Math.min(100, prompt.length)) + 
                      (prompt.length > 100 ? "..." : ""),
        documentLength: documentMarkdown.length
      });
      
      // Send the request
      const response = UrlFetchApp.fetch(apiUrl, requestOptions);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log("aiedit-debug: Received Ollama API response", {
        responseCode: responseCode,
        responseLength: responseText.length
      });
      
      if (responseCode !== 200) {
        throw new Error(`Ollama API request failed with status ${responseCode}: ${responseText}`);
      }
      
      // Parse the response
      const responseJson = JSON.parse(responseText);
      aiResponse = responseJson.response;
    }
    
    // Parse suggested changes
    const suggestions = parseSuggestedChanges(aiResponse);
    
    // Find text locations for suggestions
    const locatedSuggestions = findSuggestionLocations(documentMarkdown, suggestions);
    
    Logger.log("aiedit-debug: Completed processing document for inline suggestions", {
      totalSuggestions: suggestions.length,
      locatedSuggestions: locatedSuggestions.length
    });
    
    return {
      success: true,
      suggestions: locatedSuggestions,
      unlocatedCount: suggestions.length - locatedSuggestions.length,
      totalSuggestions: suggestions.length
    };
  } catch (e) {
    Logger.log("aiedit-debug: Error processing document for inline suggestions", {
      error: e.toString(),
      stack: e.stack
    });
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Save user settings to user properties
 * @param {string} settingsJson - JSON string of user settings
 * @returns {boolean} - Success status
 */
function saveUserSettings(settingsJson) {
  try {
    // Log the request
    logDebug(LOG_CONFIG.CATEGORIES.STATE, "Saving user settings", { settingsLength: settingsJson.length });
    
    // Save to user properties
    PropertiesService.getUserProperties().setProperty('userSettings', settingsJson);
    
    return true;
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error saving user settings", { error: error.toString() });
    throw new Error("Failed to save settings: " + error.message);
  }
}

/**
 * Get user settings from user properties
 * @returns {string} - JSON string of user settings
 */
function getUserSettings() {
  try {
    // Get from user properties
    const settings = PropertiesService.getUserProperties().getProperty('userSettings');
    
    logDebug(LOG_CONFIG.CATEGORIES.STATE, "Retrieved user settings", { 
      exists: !!settings,
      settingsLength: settings ? settings.length : 0 
    });
    
    return settings || '{}';
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error retrieving user settings", { error: error.toString() });
    throw new Error("Failed to retrieve settings: " + error.message);
  }
}

/**
 * Save the selected model preference
 * @param {string} modelName - The name of the selected model
 * @returns {boolean} - Success status
 */
function saveModelPreference(modelName) {
  try {
    if (!modelName) {
      logDebug(LOG_CONFIG.CATEGORIES.ERROR, "No model name provided to save");
      return false;
    }
    
    logDebug(LOG_CONFIG.CATEGORIES.STATE, "Saving model preference", { modelName: modelName });
    
    // Get current settings
    const userProperties = PropertiesService.getUserProperties();
    const settingsJson = userProperties.getProperty('userSettings') || '{}';
    
    try {
      // Parse current settings
      const settings = JSON.parse(settingsJson);
      
      // Update model preference
      settings.selectedModel = modelName;
      
      // Save updated settings
      userProperties.setProperty('userSettings', JSON.stringify(settings));
      
      logDebug(LOG_CONFIG.CATEGORIES.STATE, "Model preference saved successfully", { modelName: modelName });
      return true;
    } catch (parseError) {
      // If settings JSON is invalid, create new settings object
      const settings = { selectedModel: modelName };
      userProperties.setProperty('userSettings', JSON.stringify(settings));
      
      logDebug(LOG_CONFIG.CATEGORIES.STATE, "Created new settings with model preference", { modelName: modelName });
      return true;
    }
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error saving model preference", { error: error.toString() });
    return false;
  }
}

/**
 * Gets the debug logs for display in the sidebar
 * @return {string} The debug logs as a string
 */
function getDebugLogs() {
  try {
    // Get logs from Logger
    const logs = Logger.getLog();
    
    // Log that we're retrieving logs (meta-logging)
    console.log("Retrieving debug logs, size: " + (logs ? logs.length : 0));
    
    // If no logs, return a message
    if (!logs || logs.trim() === "") {
      return "No logs available";
    }
    
    return logs;
  } catch (error) {
    console.error("Error retrieving debug logs:", error);
    return "Error retrieving logs: " + error.message;
  }
}

/**
 * Test function to process the first AI comment
 * Used for debugging purposes
 * @returns {Object} Result object with status and details
 */
function testProcessFirstComment() {
  try {
    logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Starting test of first AI comment");
    
    // Get all AI comments
    const comments = getAIComments();
    
    // If no comments found, return error
    if (!comments || comments.length === 0) {
      logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "No AI comments found for testing");
      return {
        success: false,
        error: "No AI comments found to test"
      };
    }
    
    // Get the first comment
    const firstComment = comments[0];
    
    // Log details about the comment
    logDebug(LOG_CONFIG.CATEGORIES.DEBUG, "Found first AI comment for testing", {
      commentId: firstComment.id,
      instruction: firstComment.instruction,
      textLength: firstComment.text.length,
      state: firstComment.state
    });
    
    // Return the comment details but don't actually process it
    // This is safer for testing purposes
    return {
      success: true,
      comment: {
        id: firstComment.id,
        instruction: firstComment.instruction,
        textPreview: firstComment.text.substring(0, Math.min(100, firstComment.text.length)) + 
                    (firstComment.text.length > 100 ? "..." : ""),
        state: firstComment.state
      }
    };
  } catch (error) {
    // Log the error
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, "Error testing first AI comment", {
      error: error.toString(),
      stack: error.stack
    });
    
    // Return error details
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Checks if the current user is a first-time user
 * @return {boolean} True if the user is a first-time user, false otherwise
 */
function checkFirstTimeUser() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const hasSeenIntro = userProperties.getProperty('hasSeenIntro');
    
    logDebug(LOG_CONFIG.CATEGORIES.STATE, 'Checking first time user status', {
      hasSeenIntro: hasSeenIntro ? 'true' : 'false'
    });
    
    if (!hasSeenIntro) {
      // Set the flag so they won't see the intro again
      userProperties.setProperty('hasSeenIntro', 'true');
      logDebug(LOG_CONFIG.CATEGORIES.NUX, 'First time user detected - showing intro');
      return true;
    }
    
    logDebug(LOG_CONFIG.CATEGORIES.NUX, 'Returning user detected');
    return false;
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, 'Error checking first time user status', {error: error.toString()});
    return false;
  }
}

/**
 * Resets the user experience by clearing all user properties and settings
 * For debugging purposes only
 * @return {boolean} True if reset was successful, false otherwise
 */
function resetUserExperience() {
  try {
    logDebug(LOG_CONFIG.CATEGORIES.NUX, 'Starting NUX reset');
    
    const userProperties = PropertiesService.getUserProperties();
    const oldSettings = userProperties.getProperties();
    
    // Log current settings before deletion
    logDebug(LOG_CONFIG.CATEGORIES.STATE, 'Current settings before reset', oldSettings);
    
    // Delete all properties
    userProperties.deleteAllProperties();
    
    // Set reset timestamp
    const resetTime = new Date().toISOString();
    userProperties.setProperty('lastNuxReset', resetTime);
    
    logDebug(LOG_CONFIG.CATEGORIES.NUX, 'NUX reset completed successfully', {
      resetTime: resetTime
    });
    
    return true;
  } catch (error) {
    logDebug(LOG_CONFIG.CATEGORIES.ERROR, 'Error resetting user experience', {
      error: error.toString(),
      stack: error.stack
    });
    return false;
  }
}
