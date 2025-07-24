/**
 * Utility functions for cleaning insight text
 */

/**
 * Removes [NEG] tags from insight text
 * @param {string} text - The insight text to clean
 * @returns {string} - The cleaned text without [NEG] tags
 */
export const removeNegTags = (text) => {
  if (!text) return '';
  
  // Replace [NEG] with empty string, regardless of position
  return text.replace(/\[NEG\]/g, '');
};

/**
 * Cleans an array of insight strings by removing [NEG] tags
 * @param {Array<string>} insights - Array of insight strings
 * @returns {Array<string>} - Cleaned array of insight strings
 */
export const cleanInsightArray = (insights) => {
  if (!insights || !Array.isArray(insights)) return [];
  
  return insights.map(insight => removeNegTags(insight));
};

/**
 * Processes insight text for PDF generation
 * @param {string} text - The insight text
 * @param {Object} pointFeedback - Feedback data
 * @param {string} pointId - The point ID
 * @returns {Object} - Processed text and feedback
 */
export const processInsightForPDF = (text, pointFeedback = {}, pointId = null) => {
  // Remove [NEG] tag
  const cleanedText = removeNegTags(text);
  
  // Extract point number if available
  const match = cleanedText.match(/^(\d+)\.\s+(.+)$/);
  
  let result = {
    text: cleanedText,
    number: null,
    content: cleanedText,
    feedback: null
  };
  
  if (match) {
    result.number = match[1];
    result.content = match[2].trim();
    
    // Format with original numbering
    result.text = `${result.number}. ${result.content}`;
    
    // Add feedback if available
    if (pointId && pointFeedback[pointId] === true) {
      result.feedback = '👍';
    } else if (pointId && pointFeedback[pointId] === false) {
      result.feedback = '👎';
    }
  }
  
  return result;
};
