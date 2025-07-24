/**
 * Utility functions for analyzing user feedback and extracting preferences
 */

/**
 * Analyzes feedback data to extract user preferences
 * @param {Object} pointFeedback - Object containing feedback (pointId: boolean)
 * @param {Object} hiddenPoints - Object containing hidden points (pointId: boolean)
 * @param {Array} insights - Array of insight text strings
 * @returns {Object} Preference data to guide future insight generation
 */
export const analyzeFeedback = (pointFeedback, hiddenPoints, insights) => {
  // Initialize preference data
  const preferences = {
    likedTopics: [],
    dislikedTopics: [],
    hiddenTopics: [],
    preferredMetrics: [],
    avoidedMetrics: [],
    keywordPreferences: {},
  };

  // Extract insight text by pointId
  const insightMap = {};

  // Process insights to create a map of pointId to insight text
  if (insights && insights.length > 0) {
    insights.forEach((insight) => {
      // Extract point number from the insight (assuming format: "1. Point text")
      const match = insight.match(/^(\d+)\.\s+(.+)$/);
      if (match) {
        const pointNumber = match[1];
        const pointText = match[2].replace(/\[NEG\]/g, "").trim();
        const pointId = `summary-${pointNumber}`;
        insightMap[pointId] = pointText;
      }
    });
  }

  // Process positive feedback
  Object.entries(pointFeedback).forEach(([pointId, isPositive]) => {
    const insightText = insightMap[pointId] || "";

    if (isPositive) {
      // Analyze liked insights
      preferences.likedTopics.push(extractTopic(insightText));
      preferences.preferredMetrics.push(...extractMetrics(insightText));
      updateKeywordPreferences(preferences.keywordPreferences, insightText, 1);
    } else {
      // Analyze disliked insights
      preferences.dislikedTopics.push(extractTopic(insightText));
      preferences.avoidedMetrics.push(...extractMetrics(insightText));
      updateKeywordPreferences(preferences.keywordPreferences, insightText, -1);
    }
  });

  // Process hidden points
  Object.entries(hiddenPoints).forEach(([pointId, isHidden]) => {
    if (isHidden) {
      const insightText = insightMap[pointId] || "";
      preferences.hiddenTopics.push(extractTopic(insightText));
      updateKeywordPreferences(preferences.keywordPreferences, insightText, -2);
    }
  });

  // Clean up and deduplicate
  preferences.likedTopics = [
    ...new Set(preferences.likedTopics.filter(Boolean)),
  ];
  preferences.dislikedTopics = [
    ...new Set(preferences.dislikedTopics.filter(Boolean)),
  ];
  preferences.hiddenTopics = [
    ...new Set(preferences.hiddenTopics.filter(Boolean)),
  ];
  preferences.preferredMetrics = [
    ...new Set(preferences.preferredMetrics.filter(Boolean)),
  ];
  preferences.avoidedMetrics = [
    ...new Set(preferences.avoidedMetrics.filter(Boolean)),
  ];

  return preferences;
};

/**
 * Extract the main topic from an insight text using NLP-inspired techniques
 * @param {string} text - The insight text
 * @returns {string} The main topic
 */
const extractTopic = (text) => {
  // Remove any [NEG] prefix and trim
  const cleanText = text.replace(/\[NEG\]/g, "").trim();

  // Convert to lowercase for consistent processing
  const lowerText = cleanText.toLowerCase();

  // List of common stop words to filter out
  const stopWords = [
    "a",
    "an",
    "the",
    "and",
    "or",
    "but",
    "is",
    "are",
    "was",
    "were",
    "be",
    "been",
    "being",
    "in",
    "on",
    "at",
    "to",
    "for",
    "with",
    "by",
    "about",
    "against",
    "between",
    "into",
    "through",
    "during",
    "before",
    "after",
    "above",
    "below",
    "from",
    "up",
    "down",
    "of",
    "this",
    "that",
    "these",
    "those",
    "has",
    "have",
    "had",
  ];

  // Split into words and filter out stop words and short words
  const words = lowerText
    .split(/\W+/)
    .filter((word) => word.length > 3 && !stopWords.includes(word));

  // Count word frequency
  const wordFrequency = {};
  words.forEach((word) => {
    wordFrequency[word] = (wordFrequency[word] || 0) + 1;
  });

  // Sort words by frequency (most frequent first)
  const sortedWords = Object.entries(wordFrequency)
    .sort((a, b) => b[1] - a[1])
    .map((entry) => entry[0]);

  // If we have frequent words, use the most frequent one as the topic
  if (sortedWords.length > 0) {
    return sortedWords[0];
  }

  // Fallback: use the first 2-3 significant words
  return words.slice(0, Math.min(3, words.length)).join(" ") || "general";
};

/**
 * Extract metrics and data types mentioned in the insight
 * @param {string} text - The insight text
 * @returns {Array} List of metrics and data types mentioned
 */
const extractMetrics = (text) => {
  const metrics = [];

  // Look for numerical patterns
  if (text.match(/\d+(\.\d+)?/)) {
    metrics.push("numerical");
  }

  // Look for percentage patterns
  const percentageMatches = text.match(/\d+(\.\d+)?%/g);
  if (percentageMatches) {
    metrics.push("percentage");
  }

  // Look for currency patterns ($ or other currency symbols)
  const currencyMatches = text.match(/[$€£¥]\d+(\.\d+)?(k|M|B)?/g);
  if (currencyMatches) {
    metrics.push("currency");
  }

  // Look for time/date references
  if (
    text.match(
      /year|month|day|week|quarter|annual|monthly|daily|weekly|quarterly|period/i
    )
  ) {
    metrics.push("time-based");
  }

  // Look for growth/decline terms
  if (
    text.match(
      /increase|decrease|growth|decline|higher|lower|up|down|rise|fall|grew|dropped/i
    )
  ) {
    metrics.push("trend");
  }

  // Look for comparison terms
  if (
    text.match(
      /compare|versus|vs\.|against|relative|compared|than|difference|between|ratio/i
    )
  ) {
    metrics.push("comparison");
  }

  // Look for statistical terms
  if (
    text.match(
      /average|mean|median|mode|variance|deviation|correlation|regression|forecast|predict/i
    )
  ) {
    metrics.push("statistical");
  }

  return metrics;
};

/**
 * Update keyword preferences based on insight text
 * @param {Object} keywordPrefs - The keyword preferences object to update
 * @param {string} text - The insight text
 * @param {number} value - The preference value (positive for liked, negative for disliked)
 */
const updateKeywordPreferences = (keywordPrefs, text, value) => {
  // Remove any [NEG] prefix and trim
  const cleanText = text.replace(/\[NEG\]/g, "").trim();

  // List of common stop words to filter out
  const stopWords = [
    "a",
    "an",
    "the",
    "and",
    "or",
    "but",
    "is",
    "are",
    "was",
    "were",
    "be",
    "been",
    "being",
    "in",
    "on",
    "at",
    "to",
    "for",
    "with",
    "by",
    "about",
    "against",
    "between",
    "into",
    "through",
    "during",
    "before",
    "after",
    "above",
    "below",
    "from",
    "up",
    "down",
    "of",
    "this",
    "that",
    "these",
    "those",
    "has",
    "have",
    "had",
    "will",
    "would",
    "could",
    "should",
    "can",
    "may",
    "might",
    "must",
    "shall",
  ];

  // Split text into words and analyze
  const words = cleanText
    .toLowerCase()
    .split(/\W+/)
    .filter(
      (word) => word.length > 3 && !stopWords.includes(word) && isNaN(word)
    );

  // Extract meaningful phrases (2-3 word combinations)
  const phrases = [];
  for (let i = 0; i < words.length - 1; i++) {
    if (!stopWords.includes(words[i]) && !stopWords.includes(words[i + 1])) {
      phrases.push(`${words[i]} ${words[i + 1]}`);
    }
  }

  // Update word preferences
  words.forEach((word) => {
    if (!keywordPrefs[word]) {
      keywordPrefs[word] = 0;
    }
    keywordPrefs[word] += value;
  });

  // Update phrase preferences (with higher weight)
  phrases.forEach((phrase) => {
    if (!keywordPrefs[phrase]) {
      keywordPrefs[phrase] = 0;
    }
    // Phrases get slightly higher weight
    keywordPrefs[phrase] += value * 1.5;
  });
};

/**
 * Generate a prompt enhancement based on user preferences
 * @param {Object} preferences - The analyzed preferences
 * @returns {string} A prompt enhancement for the AI
 */
export const generatePromptEnhancement = (preferences) => {
  const promptParts = [];

  // Add liked topics
  if (preferences.likedTopics.length > 0) {
    promptParts.push(
      `Focus more on these topics: ${preferences.likedTopics.join(", ")}.`
    );
  }

  // Add disliked topics
  if (preferences.dislikedTopics.length > 0) {
    promptParts.push(
      `Reduce emphasis on these topics: ${preferences.dislikedTopics.join(
        ", "
      )}.`
    );
  }

  // Add hidden topics to avoid
  if (preferences.hiddenTopics.length > 0) {
    promptParts.push(
      `Avoid these topics entirely: ${preferences.hiddenTopics.join(", ")}.`
    );
  }

  // Add preferred metrics
  if (preferences.preferredMetrics.length > 0) {
    promptParts.push(
      `Include these types of metrics: ${preferences.preferredMetrics.join(
        ", "
      )}.`
    );
  }

  // Add keyword preferences
  const positiveKeywords = Object.entries(preferences.keywordPreferences)
    .filter(([_, value]) => value > 0)
    .map(([key, _]) => key)
    .slice(0, 5);

  const negativeKeywords = Object.entries(preferences.keywordPreferences)
    .filter(([_, value]) => value < 0)
    .map(([key, _]) => key)
    .slice(0, 5);

  if (positiveKeywords.length > 0) {
    promptParts.push(
      `Try to include these keywords: ${positiveKeywords.join(", ")}.`
    );
  }

  if (negativeKeywords.length > 0) {
    promptParts.push(
      `Avoid using these keywords: ${negativeKeywords.join(", ")}.`
    );
  }

  return promptParts.join(" ");
};
