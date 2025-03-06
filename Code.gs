function createBookmarksAndComments() {
  Logger.log("Starting createBookmarksAndComments");
  
  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId();
  Logger.log("Document ID: " + docId);
  
  var body = doc.getBody();
  var documentTab = doc.getActiveTab().asDocumentTab();
  
  // Get comments based on document contents and custom instructions.
  var testString = cleanClaudeResponse(queryClaude());
  
  // Fallback to a test string if queryClaude returns null or empty.
  if (!testString) {
    Logger.log("queryClaude returned null or empty string. Exiting the function.");
    return;
  }
  
  var locationsAndComments = parseLocationsAndComments(testString);
  
  // Loop through each tuple.
  locationsAndComments.forEach(function(tuple) {
    var searchText = tuple[0];
    var insertionText = tuple[1];
    Logger.log("Processing tuple: searchText='" + searchText + "', insertionText='" + insertionText + "'");
    
    // Find all occurrences of searchText.
    var searchResult = body.findText(searchText);
  
    while (searchResult !== null) {
      try {
        var element = searchResult.getElement();
        var startOffset = searchResult.getStartOffset();
        Logger.log("Found '" + searchText + "' at offset " + startOffset);
        
        // Create a position at the found occurrence.
        var position = doc.newPosition(element, startOffset);
        Logger.log("Created position for bookmark.");
        
        // Insert a bookmark at that position.
        var bookmark = documentTab.addBookmark(position);
        Logger.log("Bookmark created with ID: " + bookmark.getId());
        
        // Generate the bookmark link.
        var bookmarkLink = getBookmarkLink(bookmark);
        Logger.log("Bookmark link: " + bookmarkLink);
        
        // Build the comment text.
        var commentText = bookmarkLink + "\n" + insertionText;
        Logger.log("Comment text: " + commentText);
        
        // Add an unanchored comment with the generated text.
        addComment(docId, commentText);
      } catch (e) {
        Logger.log("Error processing occurrence of '" + searchText + "': " + e);
      }
      
      // Search for the next occurrence after the current one.
      searchResult = body.findText(searchText, searchResult);
    }
  });
  Logger.log("Finished createBookmarksAndComments");
}

function getBookmarkLink(bookmark) {
  var docId = DocumentApp.getActiveDocument().getId();
  var link = "https://docs.google.com/document/d/" + docId + "/edit?tab=t.0#bookmark=" + bookmark.getId();
  Logger.log("getBookmarkLink returning: " + link);
  return link;
}

function addComment(docId, commentText) {
  Logger.log("Adding comment to docId: " + docId);
  var payload = {
    content: commentText
  };
  Logger.log("Comment payload: " + JSON.stringify(payload));
  
  // URL for adding a comment via Drive API.
  var url = "https://www.googleapis.com/drive/v2/files/" + docId + "/comments";
  Logger.log("Posting comment to URL: " + url);
  
  var params = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, params);
    Logger.log("Drive API response: " + response.getContentText());
  } catch (e) {
    Logger.log("Error in addComment: " + e);
  }
}

function deleteAllBookmarks() {
  var doc = DocumentApp.getActiveDocument();
  var bookmarks = doc.getBookmarks();

  for (var i = 0; i < bookmarks.length; i++) {
    bookmarks[i].remove();
  }

  Logger.log('Removed ' + bookmarks.length + ' bookmarks');
}

function parseLocationsAndComments(inputString) {
  if (!inputString) {
    Logger.log("Input string is null or empty. Returning empty array.");
    return [];
  }
  return inputString
    .trim() // Remove any leading/trailing spaces or newlines
    .split("\n") // Split into lines
    .map(function(line) {
      return line.split("|||");
    }); // Split each line into a tuple
}

function queryClaude() {
  // Define instructions for the output format.
  const writingInstructions = `Instructions: Please review the preceding document and follow these instructions:
  
  The ‘In a nutshell’ section should discuss elements of the grant that are covered in the following sections. For each section of the grant page after the ‘In a nutshell’ section, excepting ‘Internal forecasts’, ‘Our process’. “Plans for Follow-up’, ‘GiveWell context’, and ‘Relationship disclosures’ find the accompanying part of the ‘In a nutshell’ and comment there on whether the corresponding section later in the document:


(1) Accurately summarizes the most important aspects of the grant page discussed in the section below
(2) Does not omit any important considerations
(3) Does not discuss minor or ancillary aspects of the grant

You should also comment on whether all points discussed in the ‘In a nutshell’ section are covered in more depth and that claims are appropriately cited in the later sections of the page. Comment on any claims in the ‘In a nutshell’ section that aren’t discussed further in the page.

The person to whom you're providing commentary welcomes critical feedback and is interested in making their document as good as it can be.

FORMAT REQUIREMENTS (CRITICAL):
Your output MUST strictly follow this exact format with no exceptions:
- Each text/comment pair must be formatted as: TextToBeCommentedOn|||Comment
- Each pair MUST be separated by EXACTLY ONE newline character (\n)
- NEVER use consecutive newline characters (\n\n) anywhere in your output
- Text to be commented on must NOT contain newline characters
- Do not include any additional text, headers, explanations, or formatting
- No spaces before or after the ||| delimiter

TEXT SELECTION GUIDELINES:
- Select the SMALLEST POSSIBLE TEXT SNIPPET that identifies the location for your comment
- Aim for 3-7 words that uniquely identify the beginning of a sentence or section
- Never include more than 10 words in your text selection
- For long sentences, select only the first few words (enough to be unique)
- Prefer to comment on section titles, sentence beginnings, or key phrases
- Trust that the reader will understand your comment applies to the full context

EXAMPLES OF GOOD TEXT SELECTIONS:
✅ In February 2025, GiveWell recommended|||This grant summary is clear and concise.
✅ Important reservations about this grant|||These reservations are well-articulated.
✅ The GDG is responsible for|||This accurately describes the organization's role.

EXAMPLES OF BAD TEXT SELECTIONS:
❌ TOO LONG: In February 2025, GiveWell recommended a $416,292 grant to the World Health Organization's Guidelines Development Group (GDG) to fund evidence reviews and updates of guidelines for two malaria treatments
❌ TOO SHORT AND NOT UNIQUE ENOUGH: This grant
❌ TOO SHORT/NOT UNIQUE ENOUGH: The

VALIDATION STEP:
Before submitting your final response, check your entire output to ensure:
1. No instance of "\n\n" exists anywhere
2. Every text/comment pair is properly separated by exactly one "\n"
3. All text selections are brief (3-10 words) but unique enough to locate
4. No text selection exceeds 10 words

EXAMPLE OF CORRECT FORMAT:
TextToBeCommentedOn1|||Comment1
TextToBeCommentedOn2|||Comment2
TextToBeCommentedOn3|||Comment3

EXAMPLES OF INCORRECT FORMAT:
❌ TextToBeCommentedOn1|||Comment1

TextToBeCommentedOn2|||Comment2  [ERROR: Has double newline]

❌ TextToBeCommentedOn1 ||| Comment1  [ERROR: Has spaces around delimiter]

❌ TextToBeCommentedOn1|Comment1  [ERROR: Incorrect delimiter]

❌ TextToBeCommentedOn
continues on next line|||Comment  [ERROR: Newline in text portion]
`;

  // Get the document contents.
  var doc = DocumentApp.getActiveDocument();
  var documentText = doc.getBody().getText();

  // Build the prompt by combining the document text and the writing instructions.
  var prompt = "Document Content:\n" + documentText + "\n\n" + writingInstructions;

  // Define the API endpoint and your API key for Claude.
  var apiUrl = "https://api.anthropic.com/v1/messages";
  var apiKey = "YOUR_API_KEY_HERE";

  // Create the payload for the request.
  var payload = {
    model: "claude-3-7-sonnet-20250219",
    max_tokens: 2048,
    messages: [{
      role: "user",
      content: prompt
    }]
  };

  // Set up the request options including the required "anthropic-version" header.
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    muteHttpExceptions: true
  };

  // Query the Claude API.
  try {
    Logger.log("Sending request to Claude API.");
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseCode = response.getResponseCode();
    Logger.log("Response code: " + responseCode);

    var responseText = response.getContentText();
    Logger.log("Response text: " + responseText);

    var result = JSON.parse(responseText);

    // Check for the expected content field in the response
    if (result && result.content && result.content[0] && result.content[0].text) {
      Logger.log("QueryClaude completed successfully.");
      return result.content[0].text;
    } else {
      Logger.log("Unexpected response format: " + JSON.stringify(result));
      return null;
    }
  } catch (error) {
    Logger.log("Error querying Claude API: " + error);
    return null;
  }
}

function cleanClaudeResponse(response) {
  if (!response) return null;
  // Replace all instances of double newlines with single newlines
  return response.replace(/\n\n/g, '\n');
}