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
 * @param {Object} options - Additional options (status, etc)
 * @return {Object} Result object with success status and error if any
 */
function retryCommentUpdate(fileId, commentId, content, options = {}) {
  let lastError = null;
  
  for (let attempt = 0; attempt < RETRY_CONFIG.MAX_ATTEMPTS; attempt++) {
    try {
      Drive.Comments.insert({
        content: content,
        resolved: options.status === 'resolved'  // Convert status to resolved boolean for v3
      }, fileId, commentId);
      
      return { success: true };
    } catch (error) {
      lastError = error;
      Logger.log(`aiedit: Comment update failed (attempt ${attempt + 1}/${RETRY_CONFIG.MAX_ATTEMPTS}): ${error.message}`);
      
      if (attempt < RETRY_CONFIG.MAX_ATTEMPTS - 1) {
        const delay = calculateBackoffDelay(attempt);
        sleep(delay);
      }
    }
  }
  
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
 * Parse a Drive API comment anchor to get text location
 * 
 * @param {String} anchor - Comment anchor string from Drive API
 * @return {Object} Location information or null if invalid
 */
function parseCommentAnchor(anchor) {
  try {
    // Drive API comment anchors are in the format:
    // kix.{random}:{startOffset}-{endOffset}
    const match = anchor.match(/kix\.[^:]+:(\d+)-(\d+)/);
    if (!match) {
      return null;
    }
    
    return {
      startOffset: parseInt(match[1], 10),
      endOffset: parseInt(match[2], 10)
    };
  } catch (e) {
    Logger.log("aiedit: Error parsing comment anchor: " + e.message);
    return null;
  }
}

/**
 * Verify text location is still valid and unchanged using comment anchor
 * 
 * @param {String} fileId - Document ID
 * @param {String} commentId - Comment ID
 * @param {String} originalText - Original text to verify
 * @return {Object} Location information or null if invalid
 */
function verifyTextLocation(fileId, commentId, originalText) {
  const doc = DocumentApp.getActiveDocument();
  
  try {
    // Validate comment ID first
    if (!validateCommentId(fileId, commentId)) {
      return null;
    }
    
    // Get fresh comment data with anchor
    const comment = Drive.Comments.get(fileId, commentId);
    if (!comment || comment.status === 'resolved' || !comment.anchor) {
      return null;
    }
    
    // Parse the anchor to get precise location
    const anchorLocation = parseCommentAnchor(comment.anchor);
    if (!anchorLocation) {
      return null;
    }
    
    // Get the document text
    const body = doc.getBody();
    const fullText = body.getText();
    
    // Verify text at anchor location matches original
    const textAtLocation = fullText.substring(anchorLocation.startOffset, anchorLocation.endOffset + 1);
    if (textAtLocation !== originalText) {
      return null;
    }
    
    // Find the element and offset within that element
    let currentOffset = 0;
    const elements = body.getNumChildren();
    
    for (let i = 0; i < elements; i++) {
      const element = body.getChild(i);
      if (element.getType() !== DocumentApp.ElementType.TEXT) {
        continue;
      }
      
      const elementText = element.asText().getText();
      const elementLength = elementText.length;
      
      // Check if our target range overlaps with this element
      if (currentOffset + elementLength > anchorLocation.startOffset) {
        const elementStartOffset = Math.max(0, anchorLocation.startOffset - currentOffset);
        const elementEndOffset = Math.min(elementLength - 1, anchorLocation.endOffset - currentOffset);
        
        // Verify the text portion in this element
        const elementPortion = elementText.substring(elementStartOffset, elementEndOffset + 1);
        const expectedPortion = originalText.substring(
          Math.max(0, currentOffset - anchorLocation.startOffset),
          Math.min(originalText.length, currentOffset + elementLength - anchorLocation.startOffset)
        );
        
        if (elementPortion !== expectedPortion) {
          return null;
        }
        
        if (anchorLocation.startOffset >= currentOffset && 
            anchorLocation.endOffset < currentOffset + elementLength) {
          // Found the complete text within this element
          return {
            element: element.asText(),
            startOffset: elementStartOffset,
            endOffset: elementEndOffset,
            anchorStart: anchorLocation.startOffset,
            anchorEnd: anchorLocation.endOffset
          };
        }
      }
      
      currentOffset += elementLength;
    }
    
    return null;
    
  } catch (e) {
    Logger.log("aiedit: Error verifying text location: " + e.message);
    return null;
  }
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
  let initialDocumentVersion = null;
  let textReplaced = false;
  let commentUpdateError = null;
  let conflictLocation = null;
  
  try {
    // Get initial document state for concurrent edit detection
    initialDocumentVersion = doc.getBody().getText();
    
    // Validate inputs
    if (!fileId || !commentId || !suggestedText) {
      throw new Error("Missing required parameters");
    }
    
    if (!validateCommentId(fileId, commentId)) {
      throw new Error("Invalid comment ID");
    }
    
    // Get the comment and verify it exists
    const comments = Drive.Comments.list(fileId, {
      fields: 'comments(id,quotedFileContent,anchor,resolved)',
      maxResults: 100
    }).comments || [];
    
    const comment = comments.find(c => c.id === commentId);
    if (!comment || !comment.quotedFileContent || comment.resolved) {
      throw new Error("Could not find comment or it has been resolved");
    }
    
    // Store original text for potential rollback
    originalText = comment.quotedFileContent.value;
    
    if (accepted) {
      // Verify text location and get element before making any changes
      const location = verifyTextLocation(fileId, commentId, originalText);
      if (!location) {
        throw new Error("Could not verify the exact location of the original text");
      }
      
      // Check for concurrent edits with detailed analysis
      const changes = checkDocumentChanges(initialDocumentVersion, originalText, location);
      if (changes.changed) {
        conflictLocation = changes.changeLocation;
        
        // Highlight the conflict area
        if (conflictLocation) {
          highlightConflictArea(conflictLocation);
        }
        
        const errorMessage = changes.targetAreaChanged ?
          "The section you're trying to edit has been modified by another user. " +
          "The AI suggestion was NOT applied. The conflicting area has been highlighted in yellow. " +
          "Please review the changes and try processing the AI comment again if needed." :
          "The document was modified by another user while processing. " +
          "The AI suggestion was NOT applied, but the changes were in a different section. " +
          "The relevant section has been highlighted in yellow for reference. " +
          "Please review and try processing the AI comment again if needed.";
        
        throw new Error(errorMessage);
      }
      
      // Check for concurrent edits
      const currentDocumentVersion = doc.getBody().getText();
      if (currentDocumentVersion !== initialDocumentVersion) {
        throw new Error("Document has been modified by another user - please try again");
      }
      
      // Sanitize text before making any changes
      const sanitizedText = sanitizeText(suggestedText);
      if (!sanitizedText) {
        throw new Error("Invalid suggested text after sanitization");
      }
      
      let commentUpdateSuccess = false;
      
      // First attempt to update comment
      const initialUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.ACCEPTED + '\n\nSuggestion accepted and applied:\n\nOriginal text:\n' + originalText + '\n\nNew text:\n' + sanitizedText,
        { status: 'resolved' }
      );
      
      if (initialUpdate.success) {
        commentResolved = true;
        commentUpdateSuccess = true;
      } else {
        commentUpdateError = initialUpdate.error;
        Logger.log("aiedit: Warning - Failed to update comment state after retries: " + initialUpdate.error.message);
      }
      
      try {
        // Now attempt the text replacement
        // Note: setSelection is used only for visual feedback, not for undo grouping
        doc.setSelection(doc.newRange()
          .addElement(location.element, location.startOffset, location.endOffset)
          .build());
        
        // Replace the text
        location.element.asText().deleteText(location.startOffset, location.endOffset);
        location.element.asText().insertText(location.startOffset, sanitizedText);
        
        // Verify the replacement using anchor offsets
        const fullText = doc.getBody().getText();
        const textAtAnchor = fullText.substring(location.anchorStart, location.anchorStart + sanitizedText.length);
        
        if (textAtAnchor !== sanitizedText) {
          throw new Error("Text replacement verification failed");
        }
        
        // Final concurrent edit check
        const finalDocumentVersion = doc.getBody().getText();
        if (finalDocumentVersion !== fullText) {
          throw new Error("Document was modified during the operation - changes have been reverted");
        }
        
        // If initial comment update failed, try one final time after text replacement
        if (!commentUpdateSuccess) {
          const finalUpdate = retryCommentUpdate(
            fileId,
            commentId,
            COMMENT_STATE.ACCEPTED + '\n\nSuggestion accepted and applied:\n\nOriginal text:\n' + originalText + '\n\nNew text:\n' + sanitizedText,
            { status: 'resolved' }
          );
          
          if (finalUpdate.success) {
            commentResolved = true;
            commentUpdateError = null;
          } else {
            Logger.log("aiedit: CRITICAL - Failed to update comment state after all retries");
            throw new Error(
              "Text was updated successfully, but the comment could not be marked as resolved after multiple attempts. " +
              "The document may be in an inconsistent state. Original error: " + finalUpdate.error.message
            );
          }
        }
        
      } catch (operationError) {
        // If text replacement fails but comment was resolved, try to unresolve it
        if (commentResolved) {
          try {
            Drive.Comments.insert({
              content: COMMENT_STATE.REJECTED + '\n\nError applying suggestion - operation reversed:\n\n' + operationError.message,
              status: 'open'
            }, fileId, commentId);
          } catch (commentError) {
            Logger.log("aiedit: Error unresolving comment: " + commentError.message);
          }
        }
        
        // Try to restore the original text if needed
        if (location) {
          try {
            location.element.asText().deleteText(location.startOffset, location.endOffset);
            location.element.asText().insertText(location.startOffset, originalText);
            
            // Re-verify the text was restored correctly
            const restoredText = location.element.asText().getText()
              .substring(location.startOffset, location.startOffset + originalText.length);
            if (restoredText !== originalText) {
              Logger.log("aiedit: Warning - Text restoration may not have been complete");
            }
          } catch (restoreError) {
            Logger.log("aiedit: Error restoring original text: " + restoreError.message);
          }
        }
        
        // Try to update comment to reflect the error
        if (commentResolved) {
          const errorUpdate = retryCommentUpdate(
            fileId,
            commentId,
            COMMENT_STATE.REJECTED + '\n\nError applying suggestion - operation reversed:\n\n' + operationError.message,
            { status: 'open' }
          );
          
          if (!errorUpdate.success) {
            Logger.log("aiedit: Error updating comment after failure: " + errorUpdate.error.message);
            operationError.message += "\nAdditionally, failed to update comment: " + errorUpdate.error.message;
          }
        }
        
        throw operationError;
      }
      
    } else {
      // For rejections, try to add the rejection state
      const rejectUpdate = retryCommentUpdate(
        fileId,
        commentId,
        COMMENT_STATE.REJECTED + '\n\nSuggestion rejected:\n\n' + sanitizeText(suggestedText) + '\n\nYou can edit the comment and try again.',
        { resolved: false }  // Keep comment open in v3
      );
      
      if (!rejectUpdate.success) {
        Logger.log("aiedit: Error updating comment for rejection: " + rejectUpdate.error.message);
        throw new Error("Failed to mark comment as rejected after multiple attempts: " + rejectUpdate.error.message);
      }
    }
    
    // If we had a comment update error but text replacement succeeded,
    // warn the user but don't fail the operation
    if (commentUpdateError) {
      Logger.log("aiedit: Warning - Text was updated but comment state may be incorrect: " + commentUpdateError.message);
    }
    
    Logger.log("aiedit: Successfully processed comment " + commentId + " (accepted: " + accepted + ")");
    return true;
    
  } catch (e) {
    // Enhanced error logging with conflict information
    if (conflictLocation) {
      Logger.log("aiedit: Concurrent edit detected - Conflict area: " +
        `offset ${conflictLocation.anchorStart}-${conflictLocation.anchorEnd}`);
    }
    Logger.log("aiedit: Error applying AI edit: " + e.message);
    throw new Error(e.message);
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