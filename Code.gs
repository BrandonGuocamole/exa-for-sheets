function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Exa AI')
    .addItem('Set API Key', 'showApiKeySidebar')
    .addToUi();
}

function showApiKeySidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Exa API Key');
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveApiKey(key) {
  PropertiesService.getUserProperties().setProperty('EXA_API_KEY', key);
}

function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('EXA_API_KEY');
}


/**
 * Queries the Exa /answer endpoint to provide an AI-generated answer based on search results.
 * Allows adding prefix/suffix text and optionally includes source citations.
 * By default, extracts and returns only the core answer text before any inline citations like " ([Source](URL)...)".
 *
 * @param {string} prompt The main question or prompt to send to Exa. Can be a cell reference.
 * @param {string} [prefix=""] Optional. Text to add before the main prompt.
 * @param {string} [suffix=""] Optional. Text to add after the main prompt.
 * @param {boolean} [includeCitations=FALSE] Optional. If TRUE, returns the full answer string as received from the API (including inline citations) AND appends any additional citations from the API's 'citations' array. Defaults to FALSE (extracts core answer only).
 * @return {string} The core answer, the full answer with citations, or an error message.
 * @customfunction
 */
function EXA_ANSWER(prompt, prefix, suffix, includeCitations) {
  const apiKey = getApiKey();
  if (!apiKey) return "❌ No API key set. Please use 'Set API Key' in the menu.";

  // --- Parameter Validation and Processing ---
  if (!prompt || typeof prompt !== 'string' || prompt.trim() === "") {
    return "❌ Please provide a valid prompt/question.";
  }

  const finalPrompt = `${prefix || ''} ${prompt} ${suffix || ''}`.trim();
  // Explicitly check if includeCitations is exactly TRUE.
  const shouldShowFullAnswerWithCitations = includeCitations === true;

  // --- API Call ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/answer", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ query: finalPrompt }),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    // --- Response Handling ---
    if (responseCode === 200) {
      const result = JSON.parse(responseBody);

      if (result && typeof result.answer === 'string') {
        const fullAnswerFromApi = result.answer;
        let finalOutput = fullAnswerFromApi; // Default to the full answer

        if (!shouldShowFullAnswerWithCitations) {
          // --- Default Behavior: Extract Core Answer ---
          // Find the start of the citation block pattern " (["
          const citationStartIndex = fullAnswerFromApi.indexOf(" ([");

          if (citationStartIndex !== -1) {
            // If the pattern is found, extract the text before it
            finalOutput = fullAnswerFromApi.substring(0, citationStartIndex).trim();
          } else {
            // If the pattern isn't found, assume no inline citations; use the full answer
            finalOutput = fullAnswerFromApi;
          }
          // Now finalOutput contains only the core answer part (hopefully)

        } else {
          // --- Include Citations Behavior ---
          // Use the full answer from API, and append from the citations array if present
           finalOutput = fullAnswerFromApi; // Start with the full API answer

           if (result.citations && Array.isArray(result.citations) && result.citations.length > 0) {
              const formattedCitations = result.citations.map(citation => {
                  const title = citation.title || 'Source';
                  const url = citation.url;
                  return url ? `([${title}](${url}))` : null;
              })
              .filter(c => c !== null)
              .join(', ');

              if (formattedCitations.length > 0) {
                  const separator = (finalOutput.slice(-1) !== ' ' && finalOutput.slice(-1) !== '\n') ? ' ' : '';
                  finalOutput += separator + formattedCitations; // Append the extra citations
              }
           }
           // If includeCitations is true but no citations array, finalOutput remains the fullAnswerFromApi
        }

        return finalOutput; // Return the processed output

      } else {
        return "❌ API returned a valid response, but no 'answer' field (string) was found.";
      }
    } else if (responseCode === 401) {
      return "❌ API Error: Invalid API Key.";
    } else { // Handle other errors
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return errorMessage;
    }
  } catch (e) {
    Logger.log(`EXA_ANSWER Error: ${e} for prompt: ${finalPrompt}`);
    return `❌ Script Error: ${e.message}`;
  }
}

/**
 * Retrieves the text content of a given URL using the Exa /contents endpoint.
 *
 * @param {string} url The full URL (including http/https) to fetch content from.
 * @return {string} The main text content of the URL or an error message.
 * @customfunction
 */
function EXA_CONTENTS(url) {
  const apiKey = getApiKey();
  if (!apiKey) return "❌ No API key set. Please use 'Set API Key' in the menu.";

  // Basic URL validation
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
      return "❌ Please provide a valid URL starting with http or https.";
  }

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/contents", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ urls: [url] }), // Exa's /contents endpoint expects an array of URLs
      headers: { "x-api-key": apiKey }, // Corrected header based on Exa Docs
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
        const result = JSON.parse(responseBody);
        // The response structure for /contents typically wraps results in a 'results' array
        const contentData = result.results && result.results[0];
        if (contentData) {
            // Prioritize text content based on Exa's common response structure
            return contentData.text || contentData.highlights || "No relevant content found in response.";
        } else {
            return "❌ API returned successfully, but no content data found for this URL.";
        }
    } else if (responseCode === 401) {
        return "❌ API Error: Invalid API Key. Please check your key in the menu.";
    } else {
        // Try to parse error message from API if possible, otherwise return generic error
        let errorMessage = `❌ API Error: Received status code ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return errorMessage;
    }
  } catch (e) {
    return `❌ Script Error: ${e.message}`;
  }
}

/**
 * Finds URLs similar to the input URL using the Exa /findSimilar endpoint, with optional filters.
 * Returns a vertical list of similar URLs.
 *
 * @param {string} url The URL to find similar links for (must include http/https).
 * @param {number} [numResults=5] Optional. The maximum number of similar URLs to return (1-10). Defaults to 5.
 * @param {string} [includeDomainsStr=""] Optional. Comma-separated list of domains to restrict results to (e.g., "example.com,anotherexample.org").
 * @param {string} [excludeDomainsStr=""] Optional. Comma-separated list of domains to exclude from results (e.g., "exclude.net,badsite.co").
 * @param {string} [includeTextStr=""] Optional. A phrase that MUST be present in the content of result pages.
 * @param {string} [excludeTextStr=""] Optional. A phrase that MUST NOT be present in the content of result pages.
 * @return {string[][]} A vertical array of similar URLs, or a single cell error message.
 * @customfunction
 */
function EXA_FINDSIMILAR(url, numResults, includeDomainsStr, excludeDomainsStr, includeTextStr, excludeTextStr) {
  const apiKey = getApiKey();
  if (!apiKey) return [["❌ No API key set. Please use 'Set API Key' in the menu."]];

  // --- Parameter Validation and Processing ---
  if (!url || typeof url !== 'string' || !url.startsWith('http')) {
    return [["❌ Please provide a valid URL starting with http or https."]];
  }

  // Validate and set numResults (sensible default and limits)
  const count = (typeof numResults === 'number' && numResults >= 1 && numResults <= 10)
                ? Math.floor(numResults)
                : 5; // Default to 5 if invalid, NaN, or outside 1-10 range

  // Process domain lists (comma-separated string to array)
  const processDomains = (domainStr) => {
    if (typeof domainStr === 'string' && domainStr.trim() !== '') {
      return domainStr.split(',').map(d => d.trim()).filter(d => d.length > 0);
    }
    return null; // Return null if empty or not a string
  };

  const includeDomains = processDomains(includeDomainsStr);
  const excludeDomains = processDomains(excludeDomainsStr);

  // Process text filters (use the string directly if provided)
  const includeText = (typeof includeTextStr === 'string' && includeTextStr.trim() !== '') ? [includeTextStr.trim()] : null;
  const excludeText = (typeof excludeTextStr === 'string' && excludeTextStr.trim() !== '') ? [excludeTextStr.trim()] : null;
  // Note: Exa API docs mention limit of 1 string, 5 words for text filters. We send as array[1].

  // --- Build API Payload ---
  const payload = {
    url: url,
    numResults: count,
    excludeSourceDomain: true // Good default to avoid getting the input URL back
  };

  if (includeDomains && includeDomains.length > 0) {
    payload.includeDomains = includeDomains;
  }
  if (excludeDomains && excludeDomains.length > 0) {
    payload.excludeDomains = excludeDomains;
  }
  if (includeText && includeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.includeText = includeText.slice(0, 1);
  }
  if (excludeText && excludeText.length > 0) {
      // Ensure only one item is sent if API has that restriction
     payload.excludeText = excludeText.slice(0, 1);
  }

  // --- API Call and Response Handling ---
  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/findSimilar", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        return result.results.map(item => [item.url || "N/A"]); // Map URLs, provide fallback
      } else {
        return [["✅ No similar URLs found matching the criteria."]];
      }
    } else if (responseCode === 401) {
      return [["❌ API Error: Invalid API Key."]];
    } else if (responseCode === 400) {
        // Handle potential bad request errors (e.g., invalid filters)
        let errorMessage = `❌ API Error (Bad Request): Status ${responseCode}.`;
        try {
            const errorResult = JSON.parse(responseBody);
            errorMessage += ` Message: ${errorResult.error || responseBody}`;
        } catch (e) {
            errorMessage += ` Response: ${responseBody}`;
        }
        return [[errorMessage]];
    }
    else { // Handle other errors
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    // Catch script execution errors (e.g., network issues)
    Logger.log(`EXA_FINDSIMILAR Error: ${e} for payload: ${JSON.stringify(payload)}`); // Log for debugging
    return [[`❌ Script Error: ${e.message}`]];
  }
}

/**
 * Searches the web using the Exa /search endpoint based on a query.
 * Returns a vertical list of result URLs.
 *
 * @param {string} query The search query.
 * @param {number} [numResults=5] Optional. The maximum number of result URLs to return. Defaults to 5.
 * @param {string} [searchType="auto"] Optional. The type of search ('auto', 'neural', 'keyword'). Defaults to 'auto'.
 * @return {string[]} An array of result URLs or a single cell error message.
 * @customfunction
 */
function EXA_SEARCH(query, numResults, searchType) {
  const apiKey = getApiKey();
  if (!apiKey) return [["❌ No API key set. Please use 'Set API Key' in the menu."]];

  if (!query || typeof query !== 'string') {
    return [["❌ Please provide a valid search query."]];
  }

  const count = (typeof numResults === 'number' && numResults > 0 && numResults <= 10) ? Math.floor(numResults) : 5; // Default to 5, max 10 for free tier
  const type = (searchType && ['auto', 'neural', 'keyword'].includes(searchType)) ? searchType : 'auto'; // Default to 'auto'

  try {
    const response = UrlFetchApp.fetch("https://api.exa.ai/search", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        query: query,
        numResults: count,
        type: type,
        useAutoprompt: (type !== 'keyword') // Enable autoprompt for neural/auto by default
      }),
      headers: { "x-api-key": apiKey },
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

     if (responseCode === 200) {
      const result = JSON.parse(responseBody);
      if (result && result.results && result.results.length > 0) {
        // Map results to a 2D array for vertical spill
        return result.results.map(item => [item.url]);
      } else {
        return [["✅ API returned successfully, but no search results found."]];
      }
    } else if (responseCode === 401) {
      return [["❌ API Error: Invalid API Key. Please check your key in the menu."]];
    } else {
      let errorMessage = `❌ API Error: Status ${responseCode}.`;
      try {
        const errorResult = JSON.parse(responseBody);
        errorMessage += ` Message: ${errorResult.error || responseBody}`;
      } catch (e) {
        errorMessage += ` Response: ${responseBody}`;
      }
      return [[errorMessage]];
    }
  } catch (e) {
    return [[`❌ Script Error: ${e.message}`]];
  }
}
