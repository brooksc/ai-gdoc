# Google Drive API Reference for Comments and Replies

This document summarizes the key details of the Google Drive API endpoints we use in our application, focusing on comments and replies management.

## Comments Resource

### Key Fields
- `id` (string, output-only): The ID of the comment
- `content` (string): The plain text content of the comment. Used for setting content.
- `htmlContent` (string, output-only): The content with HTML formatting for display
- `resolved` (boolean, output-only): Whether the comment has been resolved by a reply
- `anchor` (string): Region of the document the comment refers to
- `quotedFileContent` (object): The file content the comment refers to
  - `mimeType` (string): MIME type of quoted content
  - `value` (string): The quoted content itself (plain text)

### Important Notes
1. The `resolved` field is output-only and cannot be set directly
2. Comments can only be resolved by creating a reply with `action: 'resolve'`
3. When setting content, use the `content` field, not `htmlContent`
4. When displaying content, use `htmlContent` for proper formatting

### Methods

#### Get Comment
```javascript
Drive.Comments.get(fileId, commentId, {
  fields: 'id,content,resolved,anchor,quotedFileContent'
});
```

#### List Comments
```javascript
Drive.Comments.list(fileId, {
  fields: 'comments(id,content,quotedFileContent,replies,anchor,resolved)',
  pageSize: 100,  // Not maxResults
  includeDeleted: false
});
```

#### Update Comment
```javascript
Drive.Comments.update(
  {
    content: 'New content'  // Don't set resolved here
  },
  fileId,
  commentId,
  { fields: 'id,content,resolved' }
);
```

## Replies Resource

### Key Fields
- `id` (string, output-only): The ID of the reply
- `content` (string): Plain text content of the reply
- `action` (string): The action performed on parent comment. Valid values:
  - `'resolve'`: Resolves the parent comment
  - `'reopen'`: Reopens the parent comment

### Important Notes
1. To resolve a comment, create a reply with `action: 'resolve'`
2. The `content` field is required when creating a reply if no action is specified
3. When displaying content, use `htmlContent` for proper formatting

### Methods

#### Create Resolving Reply
```javascript
Drive.Replies.create(
  {
    action: 'resolve',
    content: 'Accepting changes'  // Optional when action is specified
  },
  fileId,
  commentId,
  { fields: 'id,action' }
);
```

## Common Patterns

### Accepting and Resolving a Comment
1. First update the comment content with acceptance message
2. Then create a resolving reply to mark it as resolved
```javascript
// Step 1: Update comment
Drive.Comments.update(
  { content: 'Accepted: <details>' },
  fileId,
  commentId,
  { fields: 'id,content,resolved' }
);

// Step 2: Create resolving reply
Drive.Replies.create(
  {
    action: 'resolve',
    content: 'Changes accepted'
  },
  fileId,
  commentId,
  { fields: 'id,action' }
);
```

### Error Handling Best Practices
1. Always check for comment existence before operations
2. Verify text replacements after making changes
3. Handle reply creation errors gracefully - text updates may succeed even if resolution fails
4. Use appropriate fields parameter to minimize data transfer
5. Include `includeDeleted: false` when listing or getting comments unless deleted comments are needed

## Google Apps Script DocumentApp API

### Document Access
```javascript
const doc = DocumentApp.getActiveDocument();
const documentId = doc.getId();
const documentName = doc.getName();
```

### Text Operations
When working with text elements:

1. Get text location from comment anchor:
```javascript
const text = doc.getBody().findText(anchorString);
const textElement = text.getElement();
const startOffset = text.getStartOffset();
const endOffset = text.getEndOffset();
```

2. Modify text:
```javascript
// Delete then insert (more reliable than replaceText)
textElement.deleteText(startOffset, endOffset);
textElement.insertText(startOffset, newText);

// Alternative: Use replaceText with exact string
textElement.replaceText(originalText, newText);
```

### Important Notes
1. Always verify text modifications by comparing the result with expected content
2. Text operations are atomic - either all succeed or all fail
3. Handle cases where text may have been modified by other users
4. Use deleteText + insertText instead of replaceText for more reliable updates
5. Maintain proper cursor and selection state when making text changes

## Google Apps Script UI Services

### HtmlService
Used to create and display HTML interfaces:
```javascript
// Create HTML output from file
const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
  .setTitle('AI Editor')  // Optional: Set the title
  .setWidth(300);        // Optional: Set the width

// Create HTML output from string
const html = HtmlService.createHtmlOutput('<p>Hello</p>');
```

### UI Service
Manages document UI elements:
```javascript
// Show sidebar
DocumentApp.getUi().showSidebar(html);

// Show dialog
DocumentApp.getUi().showModalDialog(html, 'Title');

// Show alert
DocumentApp.getUi().alert('Message');
```

### PropertiesService
Persistent key-value storage:
```javascript
// Get user-specific properties
const userProperties = PropertiesService.getUserProperties();

// Store a value
userProperties.setProperty('key', 'value');

// Retrieve a value
const value = userProperties.getProperty('key');

// Delete a value
userProperties.deleteProperty('key');

// Get all properties
const allProps = userProperties.getProperties();
```

### Important Notes
1. HtmlService:
   - HTML files must be included in the project
   - JavaScript in HTML must use `google.script.run` for server calls
   - Content Security Policy (CSP) restrictions apply
   - Set width and title for better UX:
     ```javascript
     HtmlService.createHtmlOutputFromFile('file.html')
       .setTitle('Title')
       .setWidth(300);
     ```

2. UI Service:
   - Limited to Google Workspace UI elements
   - Modal dialogs block until closed
   - Sidebars persist between document reloads

3. PropertiesService:
   - Values are always stored as strings
   - Limited storage capacity (check quotas)
   - User properties are specific to each user
   - Document properties are shared between users

## Google Apps Script Runtime Environment

### Execution Limitations
1. Time Quotas:
   - 6 minute execution time limit per execution
   - 30 minute total daily runtime per user
   - Use time-based triggers for longer operations

2. Memory Limits:
   - 50MB memory per execution
   - 50MB script properties storage
   - 9MB maximum response size from server to client

3. API Quotas:
   - Drive API: 20,000 queries/day/user
   - Document: 10 operations/second
   - Properties: 50,000 read/write operations/day
   - Check [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas) for latest limits

### Client-Server Communication
1. Server to Client:
```javascript
// Return value to client
return { success: true, data: result };

// Throw error to client
throw new Error('Error message');
```

2. Client to Server:
```javascript
// In HTML/JavaScript
google.script.run
  .withSuccessHandler(function(result) {
    // Handle success
  })
  .withFailureHandler(function(error) {
    // Handle error
  })
  .yourServerFunction(args);
```

3. Sidebar Communication:
```javascript
// From sidebar to parent document
google.script.host.close();  // Close sidebar
google.script.host.editor.focus();  // Focus editor

// From server to sidebar
return HtmlService.createHtmlOutput(html)
  .setTitle('Title')
  .setWidth(300);
```

### Error Handling Best Practices
1. Always use try-catch for API calls:
```javascript
try {
  // Get comment first to verify state
  const comment = Drive.Comments.get(fileId, commentId, {
    fields: 'id,content,resolved,anchor,quotedFileContent'
  });
  
  // Log useful debug info
  Logger.log('Retrieved comment:', {
    id: comment.id,
    content: comment.content.substring(0, 100),  // Truncate long content
    resolved: comment.resolved
  });
  
  // Perform update
  const updatedComment = Drive.Comments.update(
    updateObject,
    fileId,
    commentId,
    { fields: 'id,content,resolved' }
  );
} catch (error) {
  // Log error details
  Logger.log('Operation failed:', error);
  console.error('API call failed:', {
    error: error.toString(),
    stack: error.stack
  });
  
  // Return meaningful error to client
  return {
    success: false,
    error: 'Failed to update comment: ' + error.message
  };
}
2. Check preconditions:
```javascript
function updateComment(fileId, commentId, newContent) {
  // Validate input
  if (!fileId || !commentId || !newContent) {
    throw new Error('Missing required parameters');
  }
  
  // Check document access
  if (!DocumentApp.getActiveDocument()) {
    throw new Error('No active document');
  }
  
  // Verify comment exists and state
  const comment = Drive.Comments.get(fileId, commentId, {
    fields: 'id,resolved'
  });
  
  if (comment.resolved) {
    throw new Error('Cannot update resolved comment');
  }
}
```

3. Common response patterns:
```javascript
// Success response
return {
  success: true,
  data: result
};

// Error response
return {
  success: false,
  error: 'Operation failed: ' + error.message
};
```
```

### Debugging Tools
1. Logger:
```javascript
// Log debug information
Logger.log('Debug info: %s', JSON.stringify(data));

// View logs: View > Execution log
```

2. Stack Traces:
```javascript
try {
  // Code
} catch (e) {
  Logger.log('Error: ' + e.toString());
  Logger.log('Stack: ' + e.stack);
}
```

### Security Considerations
1. Authorization:
   - Use OAuth2 scopes in manifest (appsscript.json)
   - Request only necessary permissions
   - Handle authorization gracefully

2. Data Validation:
   - Validate all user inputs
   - Sanitize content before display
   - Use content security policies

3. Properties Security:
   - Use user properties for sensitive data
   - Don't store secrets in script properties
   - Encrypt sensitive data if needed
