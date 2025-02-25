# AI Editor for Google Docs Specification

## Overview

The AI Editor is a Google Docs add-on that allows users to request AI-powered edits to their document using comments. Users can highlight text and add comments starting with "AI:" followed by instructions. The add-on processes these comments using local Ollama models or cloud-based Gemini API and provides suggested edits with accept/reject functionality. It also includes a feature for generating inline document suggestions based on user-provided or pre-defined prompts, with full support for markdown formatting.

## Core Functionality

### Comment-Based Editing

1.  Users highlight text in their document.
2.  Users add a comment starting with "AI:" followed by their editing instructions.
3.  The add-on identifies these AI comments and processes them sequentially.
4.  AI suggestions are displayed in the sidebar for review, showing:
    *   Original comment and instruction
    *   Original highlighted text
    *   Suggested revision
5.  Users can accept or reject each suggestion.

### Comment States

*   Unprocessed: AI comment with no suggestion yet.
*   Processing: Comment being sent to AI for suggestions (one at a time).
*   Pending Review: Has AI suggestion awaiting user decision.
*   Accepted: Suggestion was applied and comment resolved.
*   Rejected: Suggestion was rejected, can be reprocessed.
*   Invalid: Comment without proper "AI:" prefix or required components.

### Comment Structure

*   Prefix: "AI:" (required).
*   Instruction: The editing instruction that follows the prefix.
*   Selected Text: The highlighted text the instruction applies to (required).
*   Suggestion: The AI-generated revision when processed.

### Processing Rules

1.  Only process comments that:
    *   Start with "AI:"
    *   Have highlighted text
    *   Have not been processed or were rejected.
2.  Skip comments that:
    *   Don't start with "AI:"
    *   Have no highlighted text
    *   Have already been accepted.
3.  Processing Constraints:
    *   Only one comment can be processed at a time.
    *   Clear indication of which comment is being processed.
    *   Show full context (instruction, original text, suggestion).

### Markdown Conversion

1. **Document to Markdown Conversion:**
   * The add-on can convert Google Docs document content to markdown format.
   * Supports various document elements:
     * Headings (H1-H6)
     * Lists (ordered and unordered)
     * Horizontal rules
     * Paragraphs
   * Used when sending document content to AI models to preserve structure.

2. **Markdown to Document Conversion:**
   * Converts markdown-formatted text back to Google Docs format.
   * Preserves and renders markdown formatting:
     * Headings using proper Google Docs heading styles
     * Ordered and unordered lists
     * Code blocks with monospace formatting
     * Blockquotes with indentation and styling
     * Horizontal rules
   * Handles inline formatting:
     * Bold text (`**text**` or `__text__`)
     * Italic text (`*text*` or `_text_`)
     * Code spans (`` `text` ``)
     * Links (`[text](url)`)

3. **Formatting Fallback:**
   * If AI responses lack proper markdown formatting, the system detects patterns and applies formatting:
     * All-caps lines ending with colons become headings
     * Lines ending with colons become subheadings
     * Lines starting with bullet characters are formatted as lists

## Document-Level Editing Features

### Whole-Document Suggestion Mode

1.  **Prompt Input:**
    *   The user can enter a custom prompt in a designated input field in the sidebar.
    *   The user can select a pre-defined prompt from a dropdown menu.
    *   The user has an option to "Save" a custom prompt, adding it to the list of pre-defined prompts for future use. The prompts are stored using `PropertiesService.getUserProperties()`.
    *   **Prompt Management:** Users can edit or delete saved prompts via a "Manage Prompts" option (either in the dropdown menu or a separate modal).
    *   **Prompt Preview:** When a user selects a pre-defined prompt, a preview of the full prompt text will be displayed (e.g., in a tooltip or a dedicated area in the sidebar).

2. **Inline Suggested Changes:**
   * Instead of appending text to the document, the system now provides inline suggested changes.
   * The AI analyzes the document and identifies specific passages that need improvement.
   * Each suggestion includes:
     * Original text
     * Revised text
     * Visual indication of what changed
   * Suggestions use a specific format: `<suggestion>Original text<changeto/>Revised text</suggestion>`
   * Changes display with visual differentiation:
     * Red strikethrough formatting for deletions
     * Blue highlighting for additions
   * Users can review each suggestion individually in the sidebar.

3. **Review Interface:**
   * Users navigate through suggestions one at a time
   * Each suggestion shows before/after text clearly
   * Accept/reject controls for each suggestion
   * Navigation buttons to move between suggestions
   * Progress indicator showing current position in suggestion list

## Technical Implementation

### Comment Retrieval

*   Use Drive API to get all comments from document.
*   Filter for comments starting with "AI:".
*   Verify presence of text selection.
*   Track processing state of each comment.

### AI Processing

1.  **Sequential processing of AI comments:**
    *   Process one comment at a time.
    *   Show clear processing status in UI.
    *   Display full context for current comment.
    *   Wait for AI response before processing next comment.
2.  **Processing Cancellation:** The user will have the ability to cancel the ongoing AI processing. This will stop the current request to the AI model and prevent further comments from being processed until the "Process AI Comments" button is clicked again.
3.  **AI API Timeout:** To prevent the add-on from becoming unresponsive due to long-running or failed requests, a timeout (configurable in user settings, default 300 seconds) will be implemented for each API call. If the timeout is reached, the request will be aborted, and an error will be displayed to the user.
4. **API Rate Limiting:** To prevent excessive API calls (both to Google's APIs and the AI model servers), debouncing will be implemented for fetching comments and models. This will ensure that API requests are not triggered more frequently than necessary (e.g., a 1-second debounce).
5.  **Whole-Document Processing:** For the "Process Document" feature, the entire document content is converted to markdown format and sent to the selected AI model in a single request, along with the user's chosen prompt and formatting instructions.

### Model Selection and Configuration

1. **Multiple Model Support:**
   * **Ollama (Local)**: Uses local Ollama server for AI processing
   * **Gemini 2.0 Flash (Cloud)**: Uses Google's Gemini API for AI processing
   
2. **Model Configuration:**
   * For Ollama: Select from available local models
   * For Gemini: Enter and save API key securely using UserProperties
   * Model selection persists across sessions
   * API keys remain securely stored between invocations

3. **Model-Specific Settings:**
   * Appropriate temperature and token limit settings for each model
   * Default prompt templates optimized for each model's capabilities
   * Automatic adaptation of requests to match each model's API format

### Revision Management

1.  **Accept Action:**
    *   Replace original text with AI suggestion.
    *   Mark comment as resolved/remove it.
    *   Update document state.
2.  **Reject Action:**
    *   Keep original text unchanged.
    *   Allow comment editing and reprocessing.
    *   Update comment state.
3.  **Text Verification:** Before applying any suggested edit (from comments), the add-on will verify that the original text *still* exists in the document at the expected location. This prevents accidental edits if the document has been modified since the suggestion was generated. If the verification fails, an error will be displayed, and the edit will not be applied.
4.  **Retry Mechanism:** When applying changes to comments, if the update fails, the system implements an exponential backoff retry strategy to improve reliability.

### User Interface Organization

1. **Collapsible Sections:** The UI is organized into three main collapsible sections, with only one section expanded at a time:

   a. **Section 1: Select Content**
      * Two mutually exclusive subsections:
        * **AI-Edit Comments:** 
          * Button to "Search AI Comments"
          * Button to "Process AI Comments"
        * **AI-Edit Doc:**
          * Custom/pre-defined prompt selection
          * Prompt management functions
          * "Process Document" button
   
   b. **Section 2: Review Proposed Changes**
      * Shows one suggestion at a time with:
        * Original text (with visual formatting)
        * Revised text (with visual formatting)
        * Accept/Reject buttons
        * Previous/Next navigation buttons
        * Progress indicator (e.g., "3 of 7 suggestions")
   
   c. **Section 3: Settings**
      * Model selection dropdown (Ollama/Gemini)
      * Model-specific configuration:
        * For Gemini: API key entry field
      * Debug toggle option
      * API timeout configuration

2. **State Management:** The add-on maintains its state across sidebar visibility changes and reloads using both session-based and persistent storage:
    * **Session Storage:** Retains temporary state (processing status, fetched comment list) within the current user session using `sessionStorage`.
    * **User Properties:** Maintains persistent data across sessions, including:
      * Selected model preference and API keys
      * Saved custom prompts
      * User settings (debug mode, timeout values)
      * This persistent storage uses `PropertiesService.getUserProperties()`.
    * If an error occurs during state restoration, the add-on gracefully resets to a default state.

3.  **Help/Documentation:**  A "Help" link or icon will be included in the sidebar, linking to a separate document (e.g., a Google Doc) that provides instructions and best practices for using the add-on.

### Debug & Logging System

1.  **Enhanced Logging:** The add-on includes a comprehensive logging system for debugging:
    * Categorized logs (error, debug, comment, text, state)
    * Log rotation with maximum size limit
    * Timestamped entries with document ID
    * Standardized format with optional data payloads
    * Safe JSON serialization that handles circular references
    * Truncation of large log entries

2.  **Debug Mode:** When enabled via settings:
    * Displays a debug panel in the sidebar
    * Includes "Test First Comment" functionality
    * Shows streaming logs for real-time debugging
    * Provides buttons to copy, clear, and refresh logs
    * Logs can be retrieved from both client and server

3.  **Error Diagnostics:** Error logs include:
    * Detailed error messages and stack traces
    * Context information (document ID, operation type)
    * Related data for troubleshooting
    * Cache-based persistence for debug session logs

### Error Handling

1.  Invalid comments are skipped.
2.  Network errors are reported in UI.
3.  Processing errors are logged.
4.  Failed suggestions are marked in UI.
5.  Clear error messaging for AI model connection issues.
6.  **Unload Handling:** If the user attempts to close the document or navigate away while AI processing is active, a warning will be displayed to prevent accidental interruption and potential data loss.
7. **Detailed Logging:** For debugging purposes, the add-on will include detailed logging of API requests, responses, and processing steps. These logs are primarily for development and troubleshooting.
8. **Specific Error Messages:** The add-on will provide more specific error messages to the user, where possible, to aid in troubleshooting. Examples include:
    * "Ollama server not found. Please ensure Ollama is running."
    * "Invalid prompt. Please check your prompt for errors."
    * "Document too large. Please try a smaller document or use comment-based editing."
    * "Network error connecting to AI model"
    * "Timeout error. AI model took too long to respond."
    * "Invalid Gemini API key. Please check your settings."

## Security and Limitations

1.  **Data Processing Location:**
    * For Ollama: All AI processing happens locally via Ollama.
    * For Gemini: Processing occurs on Google's servers with API key authentication.
2.  Document content stays within user's system with Ollama or Google's secure environment with Gemini.
3.  Changes require explicit user approval (either via accepting comment suggestions or inline document edits).
4.  Sequential processing to respect model limitations (for comment-based editing). Whole-document suggestions are processed in a single request.
5.  **HTML Escaping:** User-generated content is escaped to prevent Cross-Site Scripting (XSS) attacks.
6.  **Text Sanitization:** All text is sanitized before processing to remove control characters and normalize line separators.
7.  **API Key Protection:** Gemini API keys are stored securely using UserProperties and never exposed in client-side code.

