# AI Editor for Google Docs

A Google Docs add-on that enables AI-powered document editing using local Ollama models. Seamlessly integrate AI assistance into your document workflow by adding comments with instructions and receiving suggested edits.

## Features

- 🤖 Local AI Processing: Uses Ollama for secure, local AI text generation
- 💬 Comment-Based Interface: Add "AI:" comments to request edits
- 👀 Review System: Accept or reject AI suggestions
- 🔒 Privacy-First: All processing happens locally on your machine
- 🎯 Precise Editing: Target specific text with custom instructions
- ⚡ Real-Time Updates: See suggestions as they're generated
- 🔄 Flexible Workflow: Reject and reprocess suggestions as needed

## Requirements

- Google Docs (with editor access to the document)
- [Ollama](https://ollama.ai/) installed and running locally
- At least one Ollama model installed (e.g., llama2, mistral, etc.)
- Modern web browser with JavaScript enabled

## Installation

1. Install Ollama:
   ```bash
   # macOS/Linux
   curl -fsSL https://ollama.ai/install.sh | sh
   
   # Windows
   # Download from https://ollama.ai/download
   ```

2. Configure Ollama CORS:
   ```bash
   # Allow all origins (development/testing only)
   export OLLAMA_ORIGINS=*

   # For production, use specific Google Docs origins
   # export OLLAMA_ORIGINS=https://*.google.com
   ```

3. Pull an Ollama model:
   ```bash
   ollama pull llama2  # or your preferred model
   ```

4. Start Ollama:
   ```bash
   ollama serve
   ```

5. Install the Google Docs Add-on:
   - [Add-on installation instructions will be added when published]

## Usage

1. **Open the Editor**
   - In Google Docs, click `Extensions > AI Editor > Open Editor`
   - The editor sidebar will appear on the right

2. **Select Your Model**
   - Choose your preferred Ollama model from the dropdown
   - Click "Refresh" if your model isn't listed

3. **Add AI Instructions**
   - Highlight the text you want to edit
   - Add a comment starting with "AI:" followed by your instruction
   - Example: "AI: Make this paragraph more concise"

4. **Generate Suggestions**
   - Click "Generate Suggestions" in the sidebar
   - The add-on will process each AI comment sequentially
   - Review suggestions as they appear

5. **Review and Apply**
   - For each suggestion:
     - View the original text and proposed changes
     - Click "Accept" to apply the changes
     - Click "Reject" to keep the original text
   - Rejected suggestions can be reprocessed with modified instructions

## Security & Privacy

- All AI processing happens locally through Ollama
- No document content is sent to external servers
- Changes require explicit user approval
- Document integrity is protected with concurrent edit detection

## Troubleshooting

### Common Issues

1. **"Cannot connect to Ollama"**
   - Ensure Ollama is running (`ollama serve`)
   - Check if Ollama is accessible at `http://localhost:11434`

2. **"Invalid comment ID"**
   - The comment may have been deleted or modified
   - Try adding a new AI comment

3. **"Document has been modified"**
   - Another user edited the document during processing
   - Review changes and try again

4. **Model not appearing in dropdown**
   - Click the "Refresh" button
   - Verify the model is installed (`ollama list`)

### Error Messages

- `ScriptError`: Usually indicates an issue with the Google Docs API
- `SyntaxError`: Check your comment format (must start with "AI:")
- Network errors: Verify Ollama connection and model availability

## Development

The project uses:
- Google Apps Script for document integration
- Drive API v3 for comment management
- Local Ollama API for AI processing
- Modern JavaScript for the sidebar interface

## License

[License information to be added]

## Contributing

[Contribution guidelines to be added]

## Support

[Support information to be added] 