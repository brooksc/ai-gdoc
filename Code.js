/**
 * Wed Feb 19 20:55:39 PST 2025
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
    STATE: 'state'
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
          return '[Circular]';
        }
        seen.add(value);
      }
      // Truncate long strings
      if (typeof value === 'string' && value.length > 500) {
        return value.substring(0, 500) + '...[truncated]';
      }
      return value;
    }, 2);
  };

  Logger.log(safeStringify(logEntry));
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
  return text
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, "") // Remove control characters
    .replace(/\u2028|\u2029/g, "\n")              // Normalize line separators
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
    const normalizedOriginal = originalText.replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
      .replace(/\u2028|\u2029/g, "\n")
      .replace(/\s+/g, " ")
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
  const doc = DocumentApp.getActiveDocument();
  let originalText = null;
  let location = null;
  
  Logger.log("aiedit-debug: Starting applyAIEdit", {
    fileId: fileId,
    commentId: commentId,
    suggestedTextLength: suggestedText ? suggestedText.length : 0,
    accepted: accepted
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
            const reply = Drive.Replies.create(
              {
                action: 'resolve',
                content: 'Accepted AI suggestion'
              },
              fileId,
              commentId,
              { fields: 'id,action' }
            );
            Logger.log("aiedit-debug: Created resolving reply", {
              replyId: reply.id,
              action: reply.action
            });
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
      fields: "comments(id,content,quotedFileContent,anchor,resolved)",
      pageSize: 100,
      includeDeleted: false
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

    // Pause briefly to make the change visible
    sleep(2000);

    try {
      // Restore original text. Get a *fresh* location using the *original* text.
      log("Restoring original text");
      const finalLocation = verifyTextLocation(fileId, comment.id, originalText); // Use originalText!

      if (!finalLocation) {
        throw new Error("Could not verify location for restoration");
      }

      finalLocation.element.asText().deleteText(finalLocation.startOffset, finalLocation.endOffset);
      finalLocation.element.asText().insertText(finalLocation.startOffset, originalText);
    } catch (restoreError) {
        log("Error during text restoration", { error: restoreError.message, stack: restoreError.stack});
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

/**
 * Save the selected model to user properties
 * 
 * @param {String} modelName - Name of the selected model
 */
function saveSelectedModel(modelName) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('selectedModel', modelName);
}

/**
 * Get the previously selected model from user properties
 * 
 * @return {String} Previously selected model name or empty string
 */
function getSelectedModel() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('selectedModel') || '';
}

/**
 * Get server logs for debugging
 * 
 * @return {String} JSON string of logs
 */
function getServerLogs() {
  try {
    const logs = [];
    const now = new Date();
    
    // Get standard Logger logs
    const standardLogs = Logger.getLog();
    if (standardLogs) {
      const logLines = standardLogs.split('\n').filter(line => line.trim());
      logLines.forEach(line => {
        let type = 'info';
        let message = line;
        
        // Try to parse debug/error indicators
        if (line.includes('aiedit-debug:')) {
          type = 'debug';
          message = line.split('aiedit-debug:')[1].trim();
          try {
            // Try to parse JSON data if present
            const jsonStart = message.indexOf('{');
            if (jsonStart > -1) {
              const jsonPart = message.substring(jsonStart);
              const data = JSON.parse(jsonPart);
              message = message.substring(0, jsonStart).trim();
              logs.push({
                timestamp: now.toISOString(),
                type: type,
                message: message,
                data: data
              });
              return;
            }
          } catch (e) {
            // If JSON parsing fails, continue with regular message
          }
        } else if (line.toLowerCase().includes('error')) {
          type = 'error';
        }
        
        logs.push({
          timestamp: now.toISOString(),
          type: type,
          message: message.trim()
        });
      });
    }
    
    // Get debug session logs from cache
    const cache = CacheService.getScriptCache();
    const debugLogs = cache.get('debugLogs');
    if (debugLogs) {
      try {
        const parsedDebugLogs = JSON.parse(debugLogs);
        logs.push(...parsedDebugLogs);
      } catch (e) {
        logs.push({
          timestamp: now.toISOString(),
          type: 'error',
          message: 'Failed to parse debug logs: ' + e.message
        });
      }
    }
    
    // If no logs found at all
    if (logs.length === 0) {
      return JSON.stringify({
        timestamp: now.toISOString(),
        type: 'info',
        message: "No logs found in the last 24 hours"
      });
    }
    
    // Sort logs by timestamp
    logs.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
    
    return JSON.stringify(logs, null, 2);
    
  } catch (e) {
    return JSON.stringify({
      timestamp: new Date().toISOString(),
      type: 'error',
      message: 'Failed to retrieve logs: ' + e.message,
      stack: e.stack
    });
  }
}
