# AI-Commenter

AI-Commenter is a Google Docs Add-on that automatically generates and inserts intelligent writing feedback throughout your document using Claude AI. The ultimate goal of this project is to train Claude on various managers' commenting style on a variety of work products, then replicate that style in future documents. The remainder of this ReadMe was written by Claude as a descriptive overview on February 2nd, 2025. I'll update the ReadMe substantially when the project is complete.

## Features

- Automatically analyzes your document's writing style
- Generates contextual comments and suggestions using Claude AI
- Creates clickable bookmarks linked to specific text passages
- Adds comments with direct links to the relevant text
- Includes ability to remove all bookmarks when needed

## How It Works

The add-on performs these main functions:

1. Reads the content of your active Google Document
2. Sends the content to Claude AI for analysis
3. Receives writing suggestions in a structured format
4. Creates bookmarks at specific locations in your text
5. Adds AI-generated comments with links to those bookmarks

## Main Functions

### `createBookmarksAndComments()`
The primary function that:
- Analyzes document content
- Creates bookmarks at specific text locations
- Adds AI-generated comments with links to those bookmarks

### `queryClaude()`
- Sends document content to Claude AI
- Requests writing style analysis
- Returns formatted feedback

### `deleteAllBookmarks()`
- Utility function to remove all bookmarks from the document

## Format of AI Feedback

The AI returns feedback in the following format:
```
Text passage|||Comment about the passage
Another passage|||Another comment
```

## Requirements

- Google Docs access
- Claude API key (must be configured in the code)
- Appropriate Google Apps Script permissions

## Setup

1. Add your Claude API key in the `queryClaude()` function
2. Deploy as a Google Docs Add-on
3. Grant necessary permissions

## Note

This tool requires a valid Claude API key to function. The current version uses Claude 3.5 Sonnet model for analysis.