# AI Editor for Google Docs Specification

## Overview
The AI Editor is a Google Docs add-on that allows users to request AI-powered edits to their document using comments. Users can highlight text and add comments starting with "AI:" followed by instructions. The add-on processes these comments using local Ollama models and provides suggested edits with accept/reject functionality.

## Core Functionality

### Comment Processing
1. Users highlight text in their document
2. Users add a comment starting with "AI:" followed by their editing instructions
3. The add-on identifies these AI comments and processes them sequentially
4. AI suggestions are displayed in the sidebar for review, showing:
   - Original comment and instruction
   - Original highlighted text
   - Suggested revision
5. Users can accept or reject each suggestion

### Comment States
- Unprocessed: AI comment with no suggestion yet
- Processing: Comment being sent to AI for suggestions (one at a time)
- Pending Review: Has AI suggestion awaiting user decision
- Accepted: Suggestion was applied and comment resolved
- Rejected: Suggestion was rejected, can be reprocessed
- Invalid: Comment without proper "AI:" prefix or required components

### Comment Structure
- Prefix: "AI:" (required)
- Instruction: The editing instruction that follows the prefix
- Selected Text: The highlighted text the instruction applies to (required)
- Suggestion: The AI-generated revision when processed

### Processing Rules
1. Only process comments that:
   - Start with "AI:"
   - Have highlighted text
   - Have not been processed or were rejected
2. Skip comments that:
   - Don't start with "AI:"
   - Have no highlighted text
   - Have already been accepted
3. Processing Constraints:
   - Only one comment can be processed at a time
   - Clear indication of which comment is being processed
   - Show full context (instruction, original text, suggestion)

## Technical Implementation

### Comment Retrieval
- Use Drive API to get all comments from document
- Filter for comments starting with "AI:"
- Verify presence of text selection
- Track processing state of each comment

### AI Processing
1. Sequential processing of AI comments:
   - Process one comment at a time
   - Show clear processing status in UI
   - Display full context for current comment
   - Wait for Ollama response before processing next comment

### Revision Management
1. Accept Action:
   - Replace original text with AI suggestion
   - Mark comment as resolved/remove it
   - Update document state

2. Reject Action:
   - Keep original text unchanged
   - Allow comment editing and reprocessing
   - Update comment state

### User Interface
1. Sidebar shows:
   - Count of unprocessed AI comments
   - Model selection dropdown
   - Process button
   - Currently processing comment with:
     * Original instruction
     * Original text
     * Processing status
   - List of pending revisions with:
     * Original text
     * Suggested revision
     * Accept/Reject buttons
   - Clear progress indicators

### Error Handling
1. Invalid comments are skipped
2. Network errors are reported in UI
3. Processing errors are logged
4. Failed suggestions are marked in UI
5. Clear error messaging for Ollama connection issues

## Security and Limitations
1. All AI processing happens locally via Ollama
2. No external API calls except to local Ollama server
3. Document content stays within user's system
4. Changes require explicit user approval
5. Sequential processing to respect Ollama limitations

## Future Enhancements
1. Batch processing options
2. Custom instruction templates
3. Style preservation
4. Formatting suggestions
5. Direct edit mode (with user permission)
6. Revision history tracking
7. Comment state persistence