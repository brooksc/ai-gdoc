# AI Editor for Google Docs Specification

## Overview

The AI Editor is a Google Docs add-on that allows users to request AI-powered edits to their document using comments. Users can highlight text and add comments starting with "AI:" followed by instructions. The add-on processes these comments using local Ollama models and provides suggested edits with accept/reject functionality. It also includes a feature for generating whole-document suggestions based on a user-provided or pre-defined prompt.

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

## Whole-Document AI Suggest Feature

1.  **Prompt Input:**
    *   The user can enter a custom prompt in a designated input field in the sidebar.
    *   The user can select a pre-defined prompt from a dropdown menu.
    *   The user has an option to "Save" a custom prompt, adding it to the list of pre-defined prompts for future use. The prompts are stored using `google.script.Properties`.
    *   **Prompt Management:** Users can edit or delete saved prompts via a "Manage Prompts" option (either in the dropdown menu or a separate modal).
    *   **Prompt Preview:** When a user selects a pre-defined prompt, a preview of the full prompt text will be displayed (e.g., in a tooltip or a dedicated area in the sidebar).

2.  **AI Suggest Action:**
    *   When the user clicks an "AI Suggest" button, the add-on:
        *   Retrieves the entire content of the current document.
        *   Sends the document content and the selected prompt to the chosen Ollama model.
        *   Receives the AI-generated text from Ollama.
        *   Appends the AI-generated text to the *end* of the document.
        *   Inserts a horizontal rule (`<hr>`) above the appended text to visually separate it from the original content.
        *   Adds an H1 heading above the appended text, styled in red, with the text "AI Enhanced Text".

3.  **User Interaction:**
    *   The user can then manually review, copy, paste, and integrate parts of the AI-generated text into their original document as desired. This is a manual process; the add-on does not automatically merge the content.

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
    *   Wait for Ollama response before processing next comment.
2.  **Processing Cancellation:** The user will have the ability to cancel the ongoing AI processing. This will stop the current request to Ollama (if any) and prevent further comments from being processed until the "Generate Suggestions" button is clicked again.
3.  **Ollama API Timeout:** To prevent the add-on from becoming unresponsive due to long-running or failed Ollama requests, a timeout (e.g., 30 seconds) will be implemented for each API call. If the timeout is reached, the request will be aborted, and an error will be displayed to the user.
4. **API Rate Limiting:** To prevent excessive API calls (both to Google's APIs and the local Ollama server), debouncing will be implemented for fetching comments and models. This will ensure that API requests are not triggered more frequently than necessary (e.g., a 1-second debounce).
5.  **Whole-Document Processing:** For the "AI Suggest" feature, the entire document content will be sent to Ollama in a single request, along with the user's chosen prompt.

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

### User Interface

1.  Sidebar shows:
    *   Count of unprocessed AI comments.
    *   Model selection dropdown.
    *   Process button (for comments).
    *   Currently processing comment (when applicable) with:
        *   Original instruction
        *   Original text
        *   Processing status
    *   List of pending revisions (for comments) with:
        *   Original text
        *   Suggested revision
        *   Accept/Reject buttons
    *   Clear progress indicators.
    *   **Tooltips:** UI elements (buttons, input fields, etc.) will have tooltips providing brief explanations of their functionality.
    * **AI Suggest Section:**
        *   Input field for custom prompts.
        *   Dropdown menu for selecting pre-defined prompts.
        *    "Save Prompt" button to add the current custom prompt to the pre-defined list.
        *   "AI Suggest" button to trigger the whole-document AI generation.
        *    "Manage Prompts" option (in dropdown or separate modal) for editing/deleting saved prompts.
        *    Prompt preview area (tooltip or dedicated section).

2. **Refresh Comments Button:** If no unprocessed AI comments are located, the add-on will display a `Refresh Comments` button.
3.  **State Management:** The add-on will maintain its state (selected model, processing status, fetched comment list, and custom prompts) across sidebar visibility changes and reloads within the same user session. This will be achieved using `sessionStorage` and `google.script.Properties`. If an error occurs during state restoration, the add-on should gracefully reset to a default state.
4.  **Help/Documentation:**  A "Help" link or icon will be included in the sidebar, linking to a separate document (e.g., a Google Doc) that provides instructions and best practices for using the add-on.

### Error Handling

1.  Invalid comments are skipped.
2.  Network errors are reported in UI.
3.  Processing errors are logged.
4.  Failed suggestions are marked in UI.
5.  Clear error messaging for Ollama connection issues.
6.  **Unload Handling:** If the user attempts to close the document or navigate away while AI processing is active, a warning will be displayed to prevent accidental interruption and potential data loss.
7. **Detailed Logging:** For debugging purposes, the add-on will include detailed logging of API requests, responses, and processing steps. These logs are primarily for development and troubleshooting.
8. **Specific Error Messages:** The add-on will provide more specific error messages to the user, where possible, to aid in troubleshooting. Examples include:
 * "Ollama server not found. Please ensure Ollama is running."
    *   "Invalid prompt. Please check your prompt for errors."
    *   "Document too large. Please try a smaller document or use comment-based editing."
    * "Network error connecting to Ollama"
    * "Timeout error. Ollama took too long to respond."

## Security and Limitations

1.  All AI processing happens locally via Ollama.
2.  No external API calls except to local Ollama server.
3.  Document content stays within user's system.
4.  Changes require explicit user approval (either via accepting comment suggestions or manually copying/pasting from the AI-generated text).
5.  Sequential processing to respect Ollama limitations (for comment-based editing).  Whole-document suggestions are processed in a single request.
6.  **HTML Escaping:** User-generated content is escaped in order to prevent Cross-Site Scripting (XSS) attacks.

