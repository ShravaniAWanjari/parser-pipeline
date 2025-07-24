import jsPDF from "jspdf";
// Note: processInsightForPDF from cleanInsights.js is not used in this implementation

/**
 * Generate a PDF from the insights data with clean formatting
 * @param {string} title - The title for the PDF
 * @param {string} summary - The summary text
 * @param {Array} additionalInsights - Additional insights array
 * @param {Object} pointFeedback - Feedback data for insights
 * @param {Object} hiddenPoints - Hidden points data
 * @returns {Promise<Blob>} - The generated PDF as a Blob
 */
export const generateInsightsPDF = async (
  title,
  summary,
  additionalInsights = [],
  pointFeedback = {},
  hiddenPoints = {}
) => {
  // Create a new PDF document
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const margin = 20;
  const contentWidth = pageWidth - 2 * margin;

  // Add title with consistent styling
  pdf.setFontSize(16); // Reduced size for better proportion
  pdf.setTextColor(16, 163, 127); // Green color
  pdf.text(title, margin, 20);

  // No "Summary" heading as per requirements

  // Add date and time with consistent styling
  pdf.setFontSize(9); // Smaller size for date
  pdf.setTextColor(100, 100, 100); // Gray color

  // Format date as MM/DD/YYYY
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const year = now.getFullYear();

  // Format time as HH:MM AM/PM
  let hours = now.getHours();
  const minutes = String(now.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  hours = hours ? hours : 12; // Convert 0 to 12
  const timeString = `${hours}:${minutes} ${ampm}`;

  // Combine date and time
  const formattedDateTime = `${month}/${day}/${year} at ${timeString}`;

  pdf.text(`Generated on ${formattedDateTime}`, margin, 30);

  // Add divider
  pdf.setDrawColor(200, 200, 200);
  pdf.line(margin, 35, pageWidth - margin, 35);

  // Add content - ensure consistent font size throughout the document
  pdf.setFontSize(11); // Standardized font size for all content
  pdf.setTextColor(50, 50, 50);

  let yPosition = 45;

  // Process summary points
  if (summary) {
    // Log the summary for debugging
    console.log("Summary data for PDF:", summary);

    // Log hidden points for debugging
    console.log("Hidden points:", hiddenPoints);

    // Enhanced splitting logic to handle various formats and special cases
    console.log("Original summary:", summary);

    // Handle the case where the summary is a single string with multiple points
    let summaryLines = [];

    // First, try to split by common line breaks
    const lineBreakSplit = summary
      .split(/\n|\r\n|\r/)
      .filter((line) => line.trim());

    if (lineBreakSplit.length > 1) {
      // We have multiple lines already
      summaryLines = lineBreakSplit;
      console.log("Split by line breaks successful");
    } else {
      // Try to detect if this is a single string with multiple points

      // Check for numbered points (e.g., "1.", "2.", etc.)
      const numberedPointsMatch = summary.match(/\d+\./g);

      if (numberedPointsMatch && numberedPointsMatch.length > 1) {
        console.log("Detected numbered points format");

        // Split by numbered points
        const parts = summary.split(/(?=\d+\.)/);
        summaryLines = parts.filter((part) => part.trim());

        console.log("Split by numbered points:", summaryLines);
      } else {
        // Try to split by sentences
        const sentenceSplit = summary
          .split(/(?<=[.!?])\s+/)
          .filter((line) => line.trim());

        if (sentenceSplit.length > 1) {
          summaryLines = sentenceSplit;
          console.log("Split by sentences successful");
        } else {
          // Last resort: split by periods, but only if they're followed by a space
          summaryLines = summary.split(/\.\s+/).filter((line) => line.trim());

          // If we still don't have multiple lines, just use the whole string as one line
          if (summaryLines.length <= 1) {
            summaryLines = [summary];
          }

          console.log("Using fallback splitting method");
        }
      }
    }

    // Clean up each line - remove any trailing periods and ensure proper formatting
    summaryLines = summaryLines
      .map((line) => {
        // Remove trailing periods
        let cleaned = line.trim();
        if (cleaned.endsWith(".")) {
          cleaned = cleaned.slice(0, -1);
        }
        // Remove any numbered point prefixes
        cleaned = cleaned.replace(/^\d+\.\s*/, "");
        return cleaned;
      })
      .filter((line) => line.trim().length > 0);

    console.log("Final processed lines:", summaryLines);

    // Log the split lines for debugging
    console.log("Split summary lines:", summaryLines);

    for (const line of summaryLines) {
      // No need to check for hidden points here anymore
      // The filtering is now done before the data is passed to the PDF generator
      const match = line.match(/^(\d+)\./);

      // Just log the point for debugging
      if (match) {
        console.log("Processing point:", match[1], line);
      }

      // Check if the insight is negative by looking for the [NEG] prefix or negative keywords
      const isNegativeInsight =
        line.includes("[NEG]") ||
        /\b(decline|decrease|negative|loss|worsened|critical|issue|down)\b/i.test(
          line
        );

      // Clean the line - remove [NEG] tags and numbered points completely
      // Make sure to handle both [NEG] and [NEG] with spaces after it
      let cleanedLine = line.replace(/\[NEG\](\s*)/g, "");
      // Remove the numbered point format (e.g., "1. ", "2. ", etc.)
      cleanedLine = cleanedLine.replace(/^(\d+)\.?\s*/, "");

      // Analyze each point individually - ensure it's properly formatted
      // This ensures each point is treated as a separate entity with its own bullet point and formatting
      cleanedLine = cleanedLine.trim();

      // Log each analyzed point for debugging
      console.log(
        "Analyzed point:",
        cleanedLine,
        "Negative:",
        isNegativeInsight
      );

      // Add feedback indicator if available
      let feedbackText = "";
      if (match) {
        const pointId = `summary-${match[1]}`;
        if (pointFeedback[pointId] === true) {
          feedbackText = " 👍";
        } else if (pointFeedback[pointId] === false) {
          feedbackText = " 👎";
        }
      }

      // Format text properly to ensure consistent display
      // Clean up any special characters or formatting issues
      let formattedText = cleanedLine.replace(/[Ø=ÜM]/g, "").trim();

      // Check for truncated text (ends abruptly)
      if (/\w+$/.test(formattedText) && formattedText.length > 5) {
        // If the text ends with a word without punctuation, it might be truncated
        // Check if it's likely truncated by looking at the last few characters
        const lastWord = formattedText.split(/\s+/).pop();

        // If the last word is very short, it might be truncated
        if (lastWord && lastWord.length <= 3) {
          // Add ellipsis to indicate truncation
          formattedText += "...";
        }
      }

      // Ensure text ends with proper punctuation
      if (formattedText.length > 0) {
        // Capitalize first letter
        formattedText =
          formattedText.charAt(0).toUpperCase() + formattedText.slice(1);

        // Add period at the end if missing and not already ending with ellipsis
        if (!/[.!?]$/.test(formattedText) && !formattedText.endsWith("...")) {
          formattedText += ".";
        }
      }

      // Split text with proper width to prevent overflow
      const textLines = pdf.splitTextToSize(
        formattedText + feedbackText,
        contentWidth - 10 // Further reduced width to ensure no overflow
      );

      // Check if we need a new page - ensure consistent positioning
      // Use reduced line height for calculations
      if (
        yPosition + textLines.length * 5 >
        pdf.internal.pageSize.getHeight() - margin
      ) {
        pdf.addPage();
        yPosition = margin;
      }

      // Draw colored bullet point - ensure perfect alignment with text
      const bulletSize = 0.8; // Optimized bullet size for perfect proportion with text
      const bulletY = yPosition + 2.5; // Adjusted to align perfectly with text

      // Set bullet color based on insight type
      if (isNegativeInsight) {
        pdf.setFillColor(234, 67, 53); // Red color for negative insights
      } else {
        pdf.setFillColor(16, 163, 127); // Green color for positive insights
      }

      // Draw the bullet point - ensure perfect alignment with text
      pdf.setLineWidth(0.2); // Thinner line width for more subtle appearance
      pdf.circle(margin - 3.5, bulletY, bulletSize, "F"); // Fine-tuned position for perfect alignment

      // No border around bullet point to keep it clean and proportional

      // Reset draw color
      pdf.setDrawColor(200, 200, 200);

      // Add the text with precise alignment with bullet points
      pdf.text(textLines, margin + 1.5, yPosition + 2.5); // Fine-tuned indent for perfect spacing with bullet points

      // Ensure consistent spacing between all points
      // Reduced spacing between points as requested
      const pointSpacing = 3; // Minimal extra space between points
      yPosition += textLines.length * 5 + pointSpacing;
    }
  }

  // Add additional insights
  if (additionalInsights.length > 0) {
    // Log additional insights for debugging
    console.log("Additional insights for PDF:", additionalInsights);

    // Double check for any hidden points
    console.log("Checking hidden points again:", Object.keys(hiddenPoints));

    // Add divider
    pdf.setDrawColor(200, 200, 200);

    // Check if we need a new page
    if (yPosition + 30 > pdf.internal.pageSize.getHeight() - margin) {
      pdf.addPage();
      yPosition = margin;
    } else {
      pdf.line(margin, yPosition, pageWidth - margin, yPosition);
      yPosition += 10;
    }

    // Process additional insights - maintain consistent font size
    pdf.setFontSize(11); // Same font size as summary points
    pdf.setTextColor(50, 50, 50);

    for (const insightSet of additionalInsights) {
      // Add personalization indicator if applicable - with consistent styling
      if (insightSet.isPersonalized) {
        pdf.setTextColor(16, 163, 127); // Green color
        pdf.setFontSize(10); // Slightly smaller for the label
        pdf.text("Personalized Insights", margin, yPosition);
        yPosition += 5;
        pdf.setFontSize(11); // Back to standard size
        pdf.setTextColor(50, 50, 50);
      }

      const insightLines = insightSet.content
        .split(/\\n|\n/)
        .filter((line) => line.trim());

      for (const line of insightLines) {
        // No need to check for hidden points here anymore
        // The filtering is now done before the data is passed to the PDF generator
        const match = line.match(/^(\d+)\./);

        // Just log the point for debugging
        if (match) {
          console.log("Processing additional point:", match[1], line);
        }

        // Check if the insight is negative by looking for the [NEG] prefix or negative keywords
        const isNegativeInsight =
          line.includes("[NEG]") ||
          /\b(decline|decrease|negative|loss|worsened|critical|issue|down)\b/i.test(
            line
          );

        // Clean the line - remove [NEG] tags and numbered points completely
        // Make sure to handle both [NEG] and [NEG] with spaces after it
        let cleanedLine = line.replace(/\[NEG\](\s*)/g, "");
        // Remove the numbered point format (e.g., "1. ", "2. ", etc.)
        cleanedLine = cleanedLine.replace(/^(\d+)\.?\s*/, "");

        // Analyze each point individually - ensure it's properly formatted
        // This ensures each point is treated as a separate entity with its own bullet point and formatting
        cleanedLine = cleanedLine.trim();

        // Log each analyzed point for debugging
        console.log(
          "Analyzed additional point:",
          cleanedLine,
          "Negative:",
          isNegativeInsight
        );

        // Add feedback indicator if available
        let feedbackText = "";
        if (match) {
          const pointId = `additional-${match[1]}`;
          if (pointFeedback[pointId] === true) {
            feedbackText = " 👍";
          } else if (pointFeedback[pointId] === false) {
            feedbackText = " 👎";
          }
        }

        // Format text properly to ensure consistent display
        // Clean up any special characters or formatting issues
        let formattedText = cleanedLine.replace(/[Ø=ÜM]/g, "").trim();

        // Check for truncated text (ends abruptly)
        if (/\w+$/.test(formattedText) && formattedText.length > 5) {
          // If the text ends with a word without punctuation, it might be truncated
          // Check if it's likely truncated by looking at the last few characters
          const lastWord = formattedText.split(/\s+/).pop();

          // If the last word is very short, it might be truncated
          if (lastWord && lastWord.length <= 3) {
            // Add ellipsis to indicate truncation
            formattedText += "...";
          }
        }

        // Ensure text ends with proper punctuation
        if (formattedText.length > 0) {
          // Capitalize first letter
          formattedText =
            formattedText.charAt(0).toUpperCase() + formattedText.slice(1);

          // Add period at the end if missing and not already ending with ellipsis
          if (!/[.!?]$/.test(formattedText) && !formattedText.endsWith("...")) {
            formattedText += ".";
          }
        }

        // Split text with proper width to prevent overflow
        const textLines = pdf.splitTextToSize(
          formattedText + feedbackText,
          contentWidth - 10 // Further reduced width to ensure no overflow
        );

        // Check if we need a new page - ensure consistent positioning
        // Use reduced line height for calculations
        if (
          yPosition + textLines.length * 5 >
          pdf.internal.pageSize.getHeight() - margin
        ) {
          pdf.addPage();
          yPosition = margin;
        }

        // Draw colored bullet point - ensure perfect alignment with text
        const bulletSize = 0.8; // Optimized bullet size for perfect proportion with text
        const bulletY = yPosition + 2.5; // Adjusted to align perfectly with text

        // Set bullet color based on insight type
        if (isNegativeInsight) {
          pdf.setFillColor(234, 67, 53); // Red color for negative insights
        } else {
          pdf.setFillColor(16, 163, 127); // Green color for positive insights
        }

        // Draw the bullet point - ensure perfect alignment with text
        pdf.setLineWidth(0.2); // Thinner line width for more subtle appearance
        pdf.circle(margin - 3.5, bulletY, bulletSize, "F"); // Fine-tuned position for perfect alignment

        // No border around bullet point to keep it clean and proportional

        // Reset draw color
        pdf.setDrawColor(200, 200, 200);

        // Add the text with precise alignment with bullet points
        pdf.text(textLines, margin + 1.5, yPosition + 2.5); // Fine-tuned indent for perfect spacing with bullet points

        // Ensure consistent spacing between all points
        // Reduced spacing between points as requested
        const pointSpacing = 3; // Minimal extra space between points
        yPosition += textLines.length * 5 + pointSpacing;
      }

      // Add minimal space between insight sets
      yPosition += 3;
    }
  }

  // Add footer with consistent styling
  const totalPages = pdf.internal.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    pdf.setPage(i);
    pdf.setFontSize(9); // Slightly smaller for footer
    pdf.setTextColor(150, 150, 150); // Light gray for footer
    pdf.text(
      `Page ${i} of ${totalPages}`,
      pageWidth - margin - 25,
      pdf.internal.pageSize.getHeight() - 10
    );
  }

  // Return the PDF as a blob
  return pdf.output("blob");
};

/**
 * Download a blob as a file
 * @param {Blob} blob - The blob to download
 * @param {string} filename - The filename
 */
export const downloadBlob = (blob, filename) => {
  // Create object URL
  const url = URL.createObjectURL(blob);

  // Create link element
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;

  // Set link to be invisible
  link.style.display = "none";

  // Add to DOM, trigger click, and remove
  document.body.appendChild(link);

  // Use setTimeout to ensure the download starts immediately
  setTimeout(() => {
    link.click();

    // Clean up
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 100);
  }, 0);
};
