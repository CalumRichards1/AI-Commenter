function createBookmarksAndComments() {
  Logger.log("Starting createBookmarksAndComments");
  
  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId();
  Logger.log("Document ID: " + docId);
  
  var body = doc.getBody();
  var documentTab = doc.getActiveTab().asDocumentTab();
  
  // Get comments based on document contents and custom instructions.
  var testString = queryClaude();
  
  // Fallback to a test string if queryClaude returns null or empty.
  if (!testString) {
    Logger.log("queryClaude returned null or empty string. Using fallback test string.");
    testString = `In a quiet corner of the city|||Consider varying the sentence structure to avoid a repetitive rhythm. This phrase is effective but similar patterns appear frequently.
As the day wore on, the city slowly began to wind down.|||You might explore a more dynamic transition here to maintain reader engagement.
The city that never sleeps had finally succumbed to slumber|||This phrase is evocative but slightly contradictory. Consider clarifying whether the city is always active or if it does indeed rest.
As the first rays of dawn peeked over the horizon|||This phrase is poetic but repeated later in the piece. Consider rewording to maintain freshness.
In a hidden speakeasy tucked away behind an unassuming door|||This description is engaging but might benefit from a more unique or sensory detail to differentiate it from the rest of the nightlife descriptions.
As the night wore on, the city's secrets continued to unfold.|||This sentence builds intrigue, but the paragraph that follows shifts quickly between different settings. Consider linking them more smoothly.`;
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
  const writingInstructions = `Instructions: Please review the preceding document and provide comments on the writing style. Your output should provide comments with respect to sections of the document strictly using the following format, and including nothing else:
	•	Each entry is on a new line.
	•	The text to be commented on and the associated comment are separated by |||.
	•	No extra spaces around the delimiter.
	•	No blank lines.

Example:

TextToBeCommentedOn 1|||Comment 1  
TextToBeCommentedOn 2|||Comment 2  
TextToBeCommentedOn 3|||Comment 3
`;

  // Get the document contents.
  var doc = DocumentApp.getActiveDocument();
  var documentText = doc.getBody().getText();

  // Build the prompt by combining the document text and the writing instructions.
  var prompt = "Document Content:\n" + documentText + "\n\n" + writingInstructions;

  // Define the API endpoint and your API key for Claude.
  var apiUrl = "https://api.anthropic.com/v1/messages";
  var apiKey = "CLAUDE-API-KEY";

  // Create the payload for the request.
  var payload = {
    model: "claude-3-5-sonnet-20241022",
    max_tokens: 1024,
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