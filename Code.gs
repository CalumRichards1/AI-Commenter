function createBookmarksAndComments() {
  Logger.log("Starting createBookmarksAndComments");
  
  var doc = DocumentApp.getActiveDocument();
  var docId = doc.getId();
  Logger.log("Document ID: " + docId);
  
  var body = doc.getBody();
  var documentTab = doc.getActiveTab().asDocumentTab();
  
  // Note that the comments in testString below are from sample data in my test doc. Planning to replace it with an API call.

  var testString = `In a quiet corner of the city|||Consider varying the sentence structure to avoid a repetitive rhythm. This phrase is effective but similar patterns appear frequently.  
As the day wore on, the city slowly began to wind down.|||You might explore a more dynamic transition here to maintain reader engagement.  
The city that never sleeps had finally succumbed to slumber|||This phrase is evocative but slightly contradictory. Consider clarifying whether the city is always active or if it does indeed rest.  
As the first rays of dawn peeked over the horizon|||This phrase is poetic but repeated later in the piece. Consider rewording to maintain freshness.  
In a hidden speakeasy tucked away behind an unassuming door|||This description is engaging but might benefit from a more unique or sensory detail to differentiate it from the rest of the nightlife descriptions.  
As the night wore on, the city's secrets continued to unfold.|||This sentence builds intrigue, but the paragraph that follows shifts quickly between different settings. Consider linking them more smoothly.`
  
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
  const doc = DocumentApp.getActiveDocument();
  const bookmarks = doc.getBookmarks();

  for (let i = 0; i < bookmarks.length; i++) {
    bookmarks[i].remove();
  }

  Logger.log('Removed ' + bookmarks.length + ' bookmarks');
}

function parseLocationsAndComments(inputString) {
  return inputString
    .trim() // Remove any leading/trailing spaces or newlines
    .split("\n") // Split into lines
    .map(line => line.split("|||")); // Split each line into a tuple
}