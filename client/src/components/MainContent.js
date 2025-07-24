import React, { useState, useEffect } from "react";
import {
  Box,
  Typography,
  Paper,
  CircularProgress,
  Chip,
  Switch,
  Tooltip,
  IconButton,
  Snackbar,
  Alert,
  Badge,
} from "@mui/material";
import UploadFileIcon from "@mui/icons-material/UploadFile";
import AutoFixHighIcon from "@mui/icons-material/AutoFixHigh";
import PictureAsPdfIcon from "@mui/icons-material/PictureAsPdf";
import EmailIcon from "@mui/icons-material/Email";
import RestoreIcon from "@mui/icons-material/Restore";
import TroubleshootIcon from "@mui/icons-material/Troubleshoot";
import EmailClientDialog from "./EmailClientDialog";
import HiddenPointsDialog from "./HiddenPointsDialog";
import DeepDiveDialog from "./DeepDiveDialog";
import { useDropzone } from "react-dropzone";
import { useAppContext } from "../context/AppContext";
import api from "../services/api";
import InsightPoint from "./InsightPoint";
import {
  analyzeFeedback,
  generatePromptEnhancement,
} from "../utils/feedbackAnalyzer";
import { generateInsightsPDF, downloadBlob } from "../utils/pdfGenerator";
import "./MainContent.css";

const MainContent = () => {
  const {
    file,
    setFile,
    fileId,
    setFileId,
    isUploading,
    setIsUploading,
    uploadError,
    setUploadError,
    isProcessing,
    setIsProcessing,
    processingError,
    setProcessingError,
    tables,
    setTables,
    setSummary,
    setMarkdown,
    summary,
    setDownloadUrl,
    additionalInsights,
    setAdditionalInsights,
    isGeneratingMore,
    setIsGeneratingMore,
    pointFeedback,
    setPointFeedback,
    hiddenPoints,
    setHiddenPoints,
    resetState,
  } = useAppContext();

  // State for personalized insights
  const [hasPersonalizationData, setHasPersonalizationData] = useState(false);
  const [isPersonalizationEnabled, setIsPersonalizationEnabled] = useState(
    () => {
      // Check if user has a preference stored in localStorage
      const savedPreference = localStorage.getItem("personalizationEnabled");
      return savedPreference ? JSON.parse(savedPreference) : true; // Default to enabled
    }
  );
  const [feedbackPrompt, setFeedbackPrompt] = useState("");

  // State for hidden points dialog
  const [isHiddenPointsDialogOpen, setIsHiddenPointsDialogOpen] =
    useState(false);

  // State for deep dive dialog
  const [isDeepDiveDialogOpen, setIsDeepDiveDialogOpen] = useState(false);

  // State for PDF and email features
  const [isPdfGenerating, setIsPdfGenerating] = useState(false);
  const [isEmailDialogOpen, setIsEmailDialogOpen] = useState(false);
  const [snackbarOpen, setSnackbarOpen] = useState(false);
  const [snackbarMessage, setSnackbarMessage] = useState("");
  const [snackbarSeverity, setSnackbarSeverity] = useState("success");

  // Computed property for whether insights should be personalized
  const isPersonalized = hasPersonalizationData && isPersonalizationEnabled;

  // Save personalization preference to localStorage
  useEffect(() => {
    localStorage.setItem(
      "personalizationEnabled",
      JSON.stringify(isPersonalizationEnabled)
    );
  }, [isPersonalizationEnabled]);

  // Effect to analyze feedback and generate prompt enhancement when feedback changes
  useEffect(() => {
    // Only analyze if we have some feedback and summary data
    if (
      (Object.keys(pointFeedback).length > 0 ||
        Object.keys(hiddenPoints).length > 0) &&
      summary
    ) {
      // Get all insights from summary
      const insights = summary.split("\n").filter((line) => line.trim());

      // Analyze feedback
      const preferences = analyzeFeedback(
        pointFeedback,
        hiddenPoints,
        insights
      );

      // Generate prompt enhancement
      const prompt = generatePromptEnhancement(preferences);

      // Update state
      setFeedbackPrompt(prompt);

      // If we have meaningful feedback, mark as having personalization data
      setHasPersonalizationData(prompt.length > 0);

      console.log("Generated feedback-based prompt:", prompt);
    }
  }, [pointFeedback, hiddenPoints, summary]);

  // File upload handling
  const onDrop = async (acceptedFiles) => {
    const selectedFile = acceptedFiles[0];
    if (!selectedFile) return;

    // Reset all state when a new file is uploaded
    resetState();

    // Set the new file
    setFile(selectedFile);

    try {
      setIsUploading(true);
      setIsProcessing(true);
      setUploadError(null);
      setProcessingError(null);

      // Upload and process file in one step
      const response = await api.uploadAndProcessFile(selectedFile);

      console.log("Upload and process response:", response);

      // Set a dummy file ID for compatibility
      setFileId("processed");

      // Extract insights from response and format them for the UI
      if (response.insights && response["general-insights"]) {
        // Convert sheet insights to summary format expected by UI
        const summaryLines = [];
        let pointCounter = 1;

        // Add general insights first
        if (Array.isArray(response["general-insights"])) {
          response["general-insights"].forEach(insight => {
            summaryLines.push(`${pointCounter}. ${insight}`);
            pointCounter++;
          });
        }

        const summaryText = summaryLines.join('\n');
        setSummary(summaryText);

        // Store the raw insights for potential future use
        setTables(response.insights || {});

        console.log("Processed summary:", summaryText);
      }

    } catch (error) {
      console.error("Error uploading and processing file:", error);
      setUploadError(error.response?.data?.detail || error.message || "Failed to upload and process file");
      setProcessingError(error.response?.data?.detail || error.message || "Failed to process file");
    } finally {
      setIsUploading(false);
      setIsProcessing(false);
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
        ".xlsx",
      ],
      "application/vnd.ms-excel": [".xls"],
    },
    maxFiles: 1,
    disabled: isUploading || isProcessing,
  });

  // Process file (now simplified since upload and process happen together)
  const processFile = async (id) => {
    // This function is now mostly handled in onDrop, but keeping for compatibility
    console.log("Process file called with ID:", id);
  };

  // Handle feedback for individual points
  const handlePointFeedback = (pointId, isPositive) => {
    // Update the feedback for this specific point
    setPointFeedback((prev) => ({
      ...prev,
      [pointId]: isPositive,
    }));

    console.log(
      `User gave ${
        isPositive ? "positive" : "negative"
      } feedback for point ${pointId}`
    );

    // Save feedback to backend (if we have a fileId)
    if (fileId) {
      // Use a callback to get the latest state
      setPointFeedback((prevFeedback) => {
        // Create a new feedback object with the updated value
        const updatedFeedback = {
          ...prevFeedback,
          [pointId]: isPositive,
        };

        // Send to backend (this is just a placeholder, not actually implemented)
        api
          .saveFeedback(fileId, updatedFeedback)
          .then(() => console.log("Feedback saved successfully"))
          .catch((err) => console.error("Error saving feedback:", err));

        // Return the updated state
        return updatedFeedback;
      });
    }
  };

  // Handle removing an insight point
  const handleRemovePoint = (pointId) => {
    // Mark this point as hidden
    setHiddenPoints((prev) => ({
      ...prev,
      [pointId]: true,
    }));

    console.log(`User removed insight point ${pointId}`);

    // Save this preference to backend (if we have a fileId)
    if (fileId) {
      // Use a callback to get the latest state
      setHiddenPoints((prevHiddenPoints) => {
        // Create a new hidden points object with the updated value
        const updatedHiddenPoints = {
          ...prevHiddenPoints,
          [pointId]: true,
        };

        // Send to backend (this is just a placeholder, not actually implemented)
        api
          .saveFeedback(fileId, { hiddenPoints: updatedHiddenPoints })
          .then(() => console.log("Hidden point preference saved"))
          .catch((err) =>
            console.error("Error saving hidden point preference:", err)
          );

        // Return the updated state
        return updatedHiddenPoints;
      });
    }
  };

  // Handle opening the hidden points dialog
  const handleOpenHiddenPointsDialog = () => {
    if (hiddenPointsCount > 0) {
      setIsHiddenPointsDialogOpen(true);
    }
  };

  // Handle closing the hidden points dialog
  const handleCloseHiddenPointsDialog = () => {
    setIsHiddenPointsDialogOpen(false);
  };

  // Handle restoring a specific point
  const handleRestorePoint = (pointId) => {
    // Remove this point from hidden points
    setHiddenPoints((prev) => {
      const newHiddenPoints = { ...prev };
      delete newHiddenPoints[pointId];
      return newHiddenPoints;
    });

    console.log(`Restored insight point ${pointId}`);

    // Save this preference to backend (if we have a fileId)
    if (fileId) {
      setHiddenPoints((prevHiddenPoints) => {
        // Create a new hidden points object without the restored point
        const updatedHiddenPoints = { ...prevHiddenPoints };
        delete updatedHiddenPoints[pointId];

        // Send to backend (this is just a placeholder, not actually implemented)
        api
          .saveFeedback(fileId, { hiddenPoints: updatedHiddenPoints })
          .then(() => console.log("Hidden point preference saved"))
          .catch((err) =>
            console.error("Error saving hidden point preference:", err)
          );

        return updatedHiddenPoints;
      });
    }
  };

  // Handle restoring all hidden points
  const handleRestoreAllPoints = () => {
    // Clear all hidden points
    setHiddenPoints({});
    console.log("Restored all hidden insight points");

    // Save this preference to backend (if we have a fileId)
    if (fileId) {
      api
        .saveFeedback(fileId, { hiddenPoints: {} })
        .then(() => console.log("Hidden points preference saved"))
        .catch((err) =>
          console.error("Error saving hidden points preference:", err)
        );
    }

    // Close the dialog
    setIsHiddenPointsDialogOpen(false);
  };

  // Count hidden points
  const hiddenPointsCount = Object.keys(hiddenPoints).filter(
    (key) => hiddenPoints[key]
  ).length;

  // Generate more insights
  const generateMoreInsights = async () => {
    if (!fileId || isGeneratingMore) return;

    setIsGeneratingMore(true);

    try {
      const response = await api.generateMoreInsights();

      console.log("Generate more insights response:", response);

      if (response.additional_insights) {
        // Convert the additional insights array to the format expected by the UI
        const formattedInsights = response.additional_insights.map((insight, index) =>
          `${additionalInsights.length * 5 + index + 1}. ${insight}`
        ).join('\n');

        // Add the new insights to the existing array
        setAdditionalInsights((prevInsights) => [
          ...prevInsights,
          {
            id: Date.now(), // Unique ID for each set of insights
            content: formattedInsights,
            isPersonalized: false, // Server doesn't support personalization yet
          },
        ]);
      }
    } catch (error) {
      console.error("Error generating more insights:", error);
      // Show error to user
      setSnackbarMessage(error.response?.data?.detail || "Failed to generate additional insights");
      setSnackbarSeverity("error");
      setSnackbarOpen(true);
    } finally {
      setIsGeneratingMore(false);
    }
  };

  // Helper function to filter visible points
  const getVisibleContent = () => {
    // Filter the summary to only include visible points
    const visibleSummary = summary
      .split("\n")
      .filter((line) => {
        // Skip empty lines
        if (!line.trim()) return false;

        // Check if this is a numbered point
        const match = line.match(/^(\d+)\./);
        if (match) {
          const pointId = `summary-${match[1]}`;
          // Only include if not hidden
          return !hiddenPoints[pointId];
        }
        // Include non-numbered lines
        return true;
      })
      .join("\n");

    // Filter additional insights to only include visible points
    const visibleAdditionalInsights = additionalInsights.map((insightSet) => {
      // Create a new object with the same properties
      return {
        ...insightSet,
        // Filter the content to only include visible points
        content: insightSet.content
          .split("\n")
          .filter((line) => {
            // Skip empty lines
            if (!line.trim()) return false;

            // Check if this is a numbered point
            const match = line.match(/^(\d+)\./);
            if (match) {
              const pointId = `additional-${match[1]}`;
              // Only include if not hidden
              return !hiddenPoints[pointId];
            }
            // Include non-numbered lines
            return true;
          })
          .join("\n"),
      };
    });

    return { visibleSummary, visibleAdditionalInsights };
  };

  // Handle PDF download
  const handleDownloadPDF = async () => {
    if (!summary) return;

    try {
      setIsPdfGenerating(true);

      // Show a message
      setSnackbarMessage("Generating PDF...");
      setSnackbarSeverity("info");
      setSnackbarOpen(true);

      // Get only visible content
      const { visibleSummary, visibleAdditionalInsights } = getVisibleContent();

      console.log("Visible summary for PDF:", visibleSummary);
      console.log(
        "Visible additional insights for PDF:",
        visibleAdditionalInsights
      );

      // Generate PDF with only visible points
      const pdfBlob = await generateInsightsPDF(
        "Data Insights Report",
        visibleSummary,
        visibleAdditionalInsights,
        pointFeedback,
        hiddenPoints
      );

      // Format date for filename (MMDDYYYY_HHMM)
      const now = new Date();
      const month = String(now.getMonth() + 1).padStart(2, "0");
      const day = String(now.getDate()).padStart(2, "0");
      const year = now.getFullYear();
      const hours = String(now.getHours()).padStart(2, "0");
      const minutes = String(now.getMinutes()).padStart(2, "0");
      const dateTimeForFilename = `${month}${day}${year}_${hours}${minutes}`;

      // Download the PDF with date in filename
      downloadBlob(pdfBlob, `insights-report_${dateTimeForFilename}.pdf`);

      // Show success message after a short delay to ensure download has started
      setTimeout(() => {
        setSnackbarMessage("PDF downloaded successfully!");
        setSnackbarSeverity("success");
        setSnackbarOpen(true);
      }, 300);
    } catch (error) {
      console.error("Error generating PDF:", error);

      // Show error message
      setSnackbarMessage("Failed to generate PDF. Please try again.");
      setSnackbarSeverity("error");
      setSnackbarOpen(true);
    } finally {
      setIsPdfGenerating(false);
    }
  };

  // Handle opening email dialog
  const handleOpenEmailDialog = () => {
    setIsEmailDialogOpen(true);
  };

  // Handle closing email dialog
  const handleCloseEmailDialog = () => {
    setIsEmailDialogOpen(false);
  };

  // Handle selecting email client
  const handleSelectEmailClient = async (clientType) => {
    if (!summary) return;

    try {
      setIsPdfGenerating(true);

      // Close the dialog
      handleCloseEmailDialog();

      // Show a message
      setSnackbarMessage("Preparing email...");
      setSnackbarSeverity("info");
      setSnackbarOpen(true);

      // Get only visible content
      const { visibleSummary, visibleAdditionalInsights } = getVisibleContent();

      console.log("Visible summary for email PDF:", visibleSummary);

      // Generate PDF with only visible points
      const pdfBlob = await generateInsightsPDF(
        "Data Insights Report",
        visibleSummary,
        visibleAdditionalInsights,
        pointFeedback,
        hiddenPoints
      );

      // Format date for filename (MMDDYYYY_HHMM)
      const now = new Date();
      const month = String(now.getMonth() + 1).padStart(2, "0");
      const day = String(now.getDate()).padStart(2, "0");
      const year = now.getFullYear();
      const hours = String(now.getHours()).padStart(2, "0");
      const minutes = String(now.getMinutes()).padStart(2, "0");
      const dateTimeForFilename = `${month}${day}${year}_${hours}${minutes}`;

      // Download the PDF with date in filename
      downloadBlob(pdfBlob, `insights-report_${dateTimeForFilename}.pdf`);

      // Create email URL based on client type
      let emailUrl = "";
      const subject = encodeURIComponent("Data Insights Report");
      const body = encodeURIComponent(
        "Please find attached the data insights report that was just downloaded."
      );

      switch (clientType) {
        case "gmail":
          emailUrl = `https://mail.google.com/mail/u/0/?fs=1&tf=cm&source=mailto&su=${subject}&body=${body}`;
          break;
        case "outlook":
          // Use Outlook Web App (OWA) instead of desktop app
          emailUrl = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;
          break;
        default:
          // Default to mailto protocol which should open the default email client
          emailUrl = `mailto:?subject=${subject}&body=${body}`;
      }

      // Wait a short delay to ensure the download has started before opening email client
      setTimeout(() => {
        // Open the email client
        window.open(emailUrl, "_blank");

        // Show success message
        setSnackbarMessage(
          "Email client opened. Please attach the downloaded PDF file."
        );
        setSnackbarSeverity("success");
        setSnackbarOpen(true);
      }, 500);
    } catch (error) {
      console.error("Error preparing email:", error);
      setSnackbarMessage("Failed to prepare email. Please try again.");
      setSnackbarSeverity("error");
      setSnackbarOpen(true);
    } finally {
      setIsPdfGenerating(false);
    }
  };

  // Main render function
  return (
    <Box
      sx={{
        flexGrow: 1,
        height: "100%",
        overflow: "hidden",
        display: "flex",
        flexDirection: "column",
        bgcolor: "background.default",
      }}
    >
      {/* Main content */}
      {!summary ? (
        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "flex-start",
            minHeight: "100%",
            p: 3,
            pt: 1,
            overflow: "auto",
          }}
        >
          <Typography
            variant="h3"
            gutterBottom
            className="fade-in"
            sx={{
              fontWeight: "bold",
              color: (theme) =>
                theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
              letterSpacing: "0.01em",
              position: "relative",
              display: "inline-block",
              mt: 2,
              "&::after": {
                content: '""',
                position: "absolute",
                width: "40%",
                height: "4px",
                bottom: "-10px",
                left: "30%",
                backgroundColor: (theme) =>
                  theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                borderRadius: "2px",
              },
            }}
          >
            Data Insights
          </Typography>

          <Typography
            variant="body1"
            color="text.secondary"
            className="fade-in"
            sx={{
              mb: 3,
              mt: 2,
              textAlign: "center",
              maxWidth: 600,
              opacity: 0.9,
              lineHeight: 1.6,
            }}
          >
            Advanced financial data analysis and business intelligence
          </Typography>

          <Box
            sx={{
              display: "grid",
              gridTemplateColumns: "1fr 1fr",
              gridTemplateRows: "auto auto",
              gap: 3,
              mb: 4,
              maxWidth: 700,
              mx: "auto",
              px: 2,
            }}
          >
            <Paper
              elevation={0}
              className="fade-in card-1 hover-card"
              sx={{
                p: "12px 16px",
                bgcolor: (theme) =>
                  theme.palette.mode === "dark"
                    ? "rgba(16, 163, 127, 0.15)"
                    : "#E7F7F1",
                borderRadius: 3,
                width: "100%",
                height: "110px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                position: "relative",
                overflow: "hidden",
                boxShadow: (theme) =>
                  theme.palette.mode === "dark"
                    ? "0 4px 12px rgba(0, 0, 0, 0.2)"
                    : "0 4px 12px rgba(0, 0, 0, 0.05)",
                transition: "all 0.3s ease",
                "&:hover": {
                  transform: "translateY(-4px)",
                  boxShadow: (theme) =>
                    theme.palette.mode === "dark"
                      ? "0 8px 24px rgba(0, 0, 0, 0.3)"
                      : "0 8px 24px rgba(0, 0, 0, 0.1)",
                },
                border: (theme) =>
                  `1px solid ${
                    theme.palette.mode === "dark"
                      ? "rgba(16, 163, 127, 0.3)"
                      : "rgba(16, 163, 127, 0.2)"
                  }`,
                "&::before": {
                  content: '""',
                  position: "absolute",
                  top: 0,
                  left: 0,
                  width: "100%",
                  height: "100%",
                  background:
                    "linear-gradient(45deg, transparent 65%, rgba(255,255,255,0.1) 100%)",
                  zIndex: 1,
                },
              }}
            >
              <Box sx={{ display: "flex", alignItems: "center", mb: 0.5 }}>
                <Box
                  sx={{
                    width: 6,
                    height: 24,
                    backgroundColor: (theme) =>
                      theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    borderRadius: 1,
                    mr: 1.5,
                  }}
                />
                <Typography
                  variant="subtitle1"
                  sx={{
                    fontWeight: "bold",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    letterSpacing: "0.01em",
                  }}
                >
                  Executive Summary
                </Typography>
              </Box>
              <Typography
                variant="body2"
                color="text.primary"
                sx={{
                  pl: 3,
                  opacity: 0.9,
                  lineHeight: 1.4,
                }}
              >
                Get a concise overview of your financial data
              </Typography>
            </Paper>

            <Paper
              elevation={0}
              className="fade-in card-2 hover-card"
              sx={{
                p: "12px 16px",
                bgcolor: (theme) =>
                  theme.palette.mode === "dark"
                    ? "rgba(66, 133, 244, 0.15)"
                    : "#EFF8FF",
                borderRadius: 3,
                width: "100%",
                height: "110px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                position: "relative",
                overflow: "hidden",
                boxShadow: (theme) =>
                  theme.palette.mode === "dark"
                    ? "0 4px 12px rgba(0, 0, 0, 0.2)"
                    : "0 4px 12px rgba(0, 0, 0, 0.05)",
                transition: "all 0.3s ease",
                "&:hover": {
                  transform: "translateY(-4px)",
                  boxShadow: (theme) =>
                    theme.palette.mode === "dark"
                      ? "0 8px 24px rgba(0, 0, 0, 0.3)"
                      : "0 8px 24px rgba(0, 0, 0, 0.1)",
                },
                border: (theme) =>
                  `1px solid ${
                    theme.palette.mode === "dark"
                      ? "rgba(66, 133, 244, 0.3)"
                      : "rgba(66, 133, 244, 0.2)"
                  }`,
                "&::before": {
                  content: '""',
                  position: "absolute",
                  top: 0,
                  left: 0,
                  width: "100%",
                  height: "100%",
                  background:
                    "linear-gradient(45deg, transparent 65%, rgba(255,255,255,0.1) 100%)",
                  zIndex: 1,
                },
              }}
            >
              <Box sx={{ display: "flex", alignItems: "center", mb: 0.5 }}>
                <Box
                  sx={{
                    width: 6,
                    height: 24,
                    backgroundColor: (theme) =>
                      theme.palette.mode === "dark" ? "#60a5fa" : "#4285F4",
                    borderRadius: 1,
                    mr: 1.5,
                  }}
                />
                <Typography
                  variant="subtitle1"
                  sx={{
                    fontWeight: "bold",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#60a5fa" : "#4285F4",
                    letterSpacing: "0.01em",
                  }}
                >
                  Key Metrics
                </Typography>
              </Box>
              <Typography
                variant="body2"
                color="text.primary"
                sx={{
                  pl: 3,
                  opacity: 0.9,
                  lineHeight: 1.4,
                }}
              >
                Identify important numbers and performance indicators
              </Typography>
            </Paper>

            <Paper
              elevation={0}
              className="fade-in card-3 hover-card"
              sx={{
                p: "12px 16px",
                bgcolor: (theme) =>
                  theme.palette.mode === "dark"
                    ? "rgba(251, 188, 5, 0.15)"
                    : "#FFF8E6",
                borderRadius: 3,
                width: "100%",
                height: "110px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                position: "relative",
                overflow: "hidden",
                boxShadow: (theme) =>
                  theme.palette.mode === "dark"
                    ? "0 4px 12px rgba(0, 0, 0, 0.2)"
                    : "0 4px 12px rgba(0, 0, 0, 0.05)",
                transition: "all 0.3s ease",
                "&:hover": {
                  transform: "translateY(-4px)",
                  boxShadow: (theme) =>
                    theme.palette.mode === "dark"
                      ? "0 8px 24px rgba(0, 0, 0, 0.3)"
                      : "0 8px 24px rgba(0, 0, 0, 0.1)",
                },
                border: (theme) =>
                  `1px solid ${
                    theme.palette.mode === "dark"
                      ? "rgba(251, 188, 5, 0.3)"
                      : "rgba(251, 188, 5, 0.2)"
                  }`,
                "&::before": {
                  content: '""',
                  position: "absolute",
                  top: 0,
                  left: 0,
                  width: "100%",
                  height: "100%",
                  background:
                    "linear-gradient(45deg, transparent 65%, rgba(255,255,255,0.1) 100%)",
                  zIndex: 1,
                },
              }}
            >
              <Box sx={{ display: "flex", alignItems: "center", mb: 0.5 }}>
                <Box
                  sx={{
                    width: 6,
                    height: 24,
                    backgroundColor: (theme) =>
                      theme.palette.mode === "dark" ? "#fbbf24" : "#FBBC05",
                    borderRadius: 1,
                    mr: 1.5,
                  }}
                />
                <Typography
                  variant="subtitle1"
                  sx={{
                    fontWeight: "bold",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#fbbf24" : "#FBBC05",
                    letterSpacing: "0.01em",
                  }}
                >
                  Trend Analysis
                </Typography>
              </Box>
              <Typography
                variant="body2"
                color="text.primary"
                sx={{
                  pl: 3,
                  opacity: 0.9,
                  lineHeight: 1.4,
                }}
              >
                Discover patterns across branches and time periods
              </Typography>
            </Paper>

            <Paper
              elevation={0}
              className="fade-in card-4 hover-card"
              sx={{
                p: "12px 16px",
                bgcolor: (theme) =>
                  theme.palette.mode === "dark"
                    ? "rgba(234, 67, 53, 0.15)"
                    : "#FFEBE6",
                borderRadius: 3,
                width: "100%",
                height: "110px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "center",
                position: "relative",
                overflow: "hidden",
                boxShadow: (theme) =>
                  theme.palette.mode === "dark"
                    ? "0 4px 12px rgba(0, 0, 0, 0.2)"
                    : "0 4px 12px rgba(0, 0, 0, 0.05)",
                transition: "all 0.3s ease",
                "&:hover": {
                  transform: "translateY(-4px)",
                  boxShadow: (theme) =>
                    theme.palette.mode === "dark"
                      ? "0 8px 24px rgba(0, 0, 0, 0.3)"
                      : "0 8px 24px rgba(0, 0, 0, 0.1)",
                },
                border: (theme) =>
                  `1px solid ${
                    theme.palette.mode === "dark"
                      ? "rgba(234, 67, 53, 0.3)"
                      : "rgba(234, 67, 53, 0.2)"
                  }`,
                "&::before": {
                  content: '""',
                  position: "absolute",
                  top: 0,
                  left: 0,
                  width: "100%",
                  height: "100%",
                  background:
                    "linear-gradient(45deg, transparent 65%, rgba(255,255,255,0.1) 100%)",
                  zIndex: 1,
                },
              }}
            >
              <Box sx={{ display: "flex", alignItems: "center", mb: 0.5 }}>
                <Box
                  sx={{
                    width: 6,
                    height: 24,
                    backgroundColor: (theme) =>
                      theme.palette.mode === "dark" ? "#f87171" : "#EA4335",
                    borderRadius: 1,
                    mr: 1.5,
                  }}
                />
                <Typography
                  variant="subtitle1"
                  sx={{
                    fontWeight: "bold",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#f87171" : "#EA4335",
                    letterSpacing: "0.01em",
                  }}
                >
                  Data Highlights
                </Typography>
              </Box>
              <Typography
                variant="body2"
                color="text.primary"
                sx={{
                  pl: 3,
                  opacity: 0.9,
                  lineHeight: 1.4,
                }}
              >
                Identify key data points and important findings
              </Typography>
            </Paper>
          </Box>

          <Box sx={{ width: "100%", mb: 3, mt: 1 }}>
            {isUploading || isProcessing ? (
              <Paper
                elevation={1}
                className="fade-in subtle-shadow"
                sx={{
                  borderRadius: 2,
                  p: 3,
                  width: "100%",
                  maxWidth: 500,
                  mx: "auto",
                  textAlign: "center",
                  bgcolor: (theme) =>
                    theme.palette.mode === "dark"
                      ? "rgba(255, 255, 255, 0.05)"
                      : "#f9f9f9",
                  position: "relative",
                  overflow: "hidden",
                  border: (theme) =>
                    `1px solid ${
                      theme.palette.mode === "dark"
                        ? "rgba(255, 255, 255, 0.1)"
                        : "rgba(0, 0, 0, 0.03)"
                    }`,
                }}
              >
                <Box
                  className="animated-progress-bar"
                  sx={{
                    position: "absolute",
                    bottom: 0,
                    left: 0,
                    height: 4,
                    bgcolor: (theme) =>
                      theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    width: isUploading ? "30%" : "70%",
                    animation:
                      "progressBar 1.5s ease-in-out infinite alternate",
                  }}
                />

                <Box sx={{ mb: 2 }} className="pulse-animation">
                  <CircularProgress
                    size={48}
                    thickness={3}
                    sx={{
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    }}
                  />
                </Box>

                <Typography
                  variant="h6"
                  gutterBottom
                  className="fade-in"
                  sx={{
                    fontWeight: "medium",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    mb: 2,
                  }}
                >
                  {isUploading ? "Uploading File" : "Processing Data"}
                </Typography>

                <Typography
                  variant="body1"
                  sx={{
                    mb: 2,
                    opacity: 0.9,
                  }}
                  className="fade-in"
                >
                  {isUploading
                    ? "Your file is being uploaded to the server..."
                    : "Generating insights from your data..."}
                </Typography>

                <Typography
                  variant="body2"
                  color="text.secondary"
                  sx={{
                    opacity: 0.7,
                    mb: 3,
                  }}
                >
                  {isUploading
                    ? "This should only take a few seconds"
                    : "This may take up to 30 seconds depending on file size"}
                </Typography>

                {file && (
                  <Typography
                    variant="caption"
                    className="fade-in"
                    sx={{
                      display: "block",
                      mt: 3,
                      color: "text.secondary",
                      bgcolor: (theme) =>
                        theme.palette.mode === "dark"
                          ? "rgba(255, 255, 255, 0.05)"
                          : "rgba(0, 0, 0, 0.04)",
                      py: 1,
                      px: 2,
                      borderRadius: 2,
                      maxWidth: "80%",
                      mx: "auto",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                      border: (theme) =>
                        `1px solid ${
                          theme.palette.mode === "dark"
                            ? "rgba(255, 255, 255, 0.1)"
                            : "rgba(0, 0, 0, 0.03)"
                        }`,
                    }}
                  >
                    {file.name} ({(file.size / 1024).toFixed(1)} KB)
                  </Typography>
                )}
              </Paper>
            ) : (
              <Paper
                elevation={0}
                {...getRootProps()}
                className="fade-in"
                sx={{
                  border: "2px dashed",
                  borderColor: isDragActive ? "#10a37f" : "grey.300",
                  borderRadius: 2,
                  p: 3,
                  width: "100%",
                  maxWidth: 500,
                  mx: "auto",
                  textAlign: "center",
                  cursor: "pointer",
                  bgcolor: isDragActive
                    ? "rgba(16, 163, 127, 0.08)"
                    : "transparent",
                  transition: "all 0.3s ease",
                  position: "relative",
                  overflow: "hidden",
                  "&:hover": {
                    borderColor: "#10a37f",
                    bgcolor: "rgba(16, 163, 127, 0.04)",
                    transform: "translateY(-2px)",
                    boxShadow: "0 8px 30px rgba(0, 0, 0, 0.05)",
                  },
                  "&::before": {
                    content: '""',
                    position: "absolute",
                    top: 0,
                    left: 0,
                    width: "100%",
                    height: "100%",
                    background:
                      "linear-gradient(45deg, transparent 65%, rgba(16, 163, 127, 0.05) 100%)",
                    zIndex: 1,
                  },
                }}
              >
                <input {...getInputProps()} />

                <Box
                  className={
                    isDragActive ? "pulse-animation" : "float-animation"
                  }
                >
                  <UploadFileIcon
                    sx={{
                      fontSize: 48,
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                      mb: 2,
                      transition: "all 0.3s ease",
                    }}
                  />
                </Box>

                {isDragActive ? (
                  <Typography
                    variant="body1"
                    sx={{
                      fontWeight: "medium",
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    }}
                  >
                    Drop your Excel file here...
                  </Typography>
                ) : (
                  <>
                    <Typography
                      variant="body1"
                      gutterBottom
                      sx={{
                        fontWeight: "medium",
                        mb: 1,
                      }}
                    >
                      Drag and drop an Excel file here, or click to select
                    </Typography>
                    <Typography
                      variant="body2"
                      color="text.secondary"
                      sx={{
                        opacity: 0.8,
                        px: 4,
                      }}
                    >
                      Supported formats: .xlsx, .xls
                    </Typography>
                  </>
                )}
              </Paper>
            )}

            {uploadError && (
              <Paper
                elevation={0}
                sx={{
                  mt: 2,
                  p: 2,
                  bgcolor: (theme) =>
                    theme.palette.mode === "dark"
                      ? "rgba(244, 67, 54, 0.15)"
                      : "#FFEBEE",
                  borderRadius: 2,
                  maxWidth: 500,
                  width: "100%",
                  border: (theme) =>
                    `1px solid ${
                      theme.palette.mode === "dark"
                        ? "rgba(244, 67, 54, 0.3)"
                        : "rgba(244, 67, 54, 0.1)"
                    }`,
                }}
              >
                <Typography variant="body2" color="error">
                  Error: {uploadError}
                </Typography>
              </Paper>
            )}

            {processingError && (
              <Paper
                elevation={0}
                sx={{
                  mt: 2,
                  p: 2,
                  bgcolor: (theme) =>
                    theme.palette.mode === "dark"
                      ? "rgba(244, 67, 54, 0.15)"
                      : "#FFEBEE",
                  borderRadius: 2,
                  maxWidth: 500,
                  width: "100%",
                  border: (theme) =>
                    `1px solid ${
                      theme.palette.mode === "dark"
                        ? "rgba(244, 67, 54, 0.3)"
                        : "rgba(244, 67, 54, 0.1)"
                    }`,
                }}
              >
                <Typography variant="body2" color="error">
                  Error: {processingError}
                </Typography>
              </Paper>
            )}
          </Box>
        </Box>
      ) : (
        <Box
          sx={{
            display: "flex",
            flexDirection: "column",
            height: "100%",
            bgcolor: (theme) =>
              theme.palette.mode === "dark" ? "#121212" : "#ffffff",
          }}
        >
          {/* Summary area */}
          <Box
            sx={{
              flexGrow: 1,
              overflow: "auto",
              p: 3,
              display: "flex",
              flexDirection: "column",
              bgcolor: (theme) =>
                theme.palette.mode === "dark" ? "#121212" : "#ffffff",
            }}
          >
            <Box
              sx={{
                width: "100%",
                py: 3,
                px: { xs: 2, md: 12 },
                bgcolor: (theme) =>
                  theme.palette.mode === "dark" ? "#121212" : "#ffffff",
              }}
            >
              <Box
                sx={{
                  maxWidth: 768,
                  mx: "auto",
                  "& p": {
                    m: 0,
                    mb: 1,
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  },
                  "& a": {
                    color: (theme) =>
                      theme.palette.mode === "dark"
                        ? "#90caf9"
                        : "primary.main",
                  },
                  "& table": {
                    borderCollapse: "collapse",
                    fontSize: "0.875rem",
                    mb: 2,
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                    "& th, & td": {
                      padding: "4px 8px",
                    },
                    "& th": {
                      bgcolor: (theme) =>
                        theme.palette.mode === "dark"
                          ? "rgba(255,255,255,0.08)"
                          : "rgba(0,0,0,0.02)",
                    },
                  },
                  "& ul, & ol": {
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  },
                  "& li": {
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  },
                  "& h1, & h2, & h3, & h4, & h5, & h6": {
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  },
                  "& blockquote": {
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                    borderLeft: (theme) =>
                      `4px solid ${
                        theme.palette.mode === "dark"
                          ? "rgba(255,255,255,0.1)"
                          : "rgba(0,0,0,0.1)"
                      }`,
                    paddingLeft: 2,
                    margin: 1,
                  },
                  "& code": {
                    backgroundColor: (theme) =>
                      theme.palette.mode === "dark"
                        ? "rgba(255,255,255,0.1)"
                        : "rgba(0,0,0,0.05)",
                    padding: "2px 4px",
                    borderRadius: 1,
                    fontSize: "0.9em",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  },
                }}
              >
                {/* File name and action buttons */}
                {file && (
                  <Box
                    sx={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      mb: 2,
                    }}
                  >
                    <Typography
                      variant="caption"
                      sx={{
                        color: "text.secondary",
                        bgcolor: (theme) =>
                          theme.palette.mode === "dark"
                            ? "rgba(255,255,255,0.05)"
                            : "rgba(0,0,0,0.04)",
                        py: 0.5,
                        px: 1.5,
                        borderRadius: 10,
                        display: "flex",
                        alignItems: "center",
                        gap: 0.5,
                      }}
                    >
                      <Box
                        component="span"
                        sx={{
                          width: 8,
                          height: 8,
                          borderRadius: "50%",
                          bgcolor: (theme) =>
                            theme.palette.mode === "dark"
                              ? "#4ade80"
                              : "#10a37f",
                          display: "inline-block",
                        }}
                      />
                      {file.name}
                    </Typography>

                    {/* Action buttons */}
                    <Box sx={{ display: "flex", gap: 1, alignItems: "center" }}>
                      {/* Personalization toggle (only shown if we have personalization data) */}
                      {hasPersonalizationData && (
                        <Tooltip title="Generate insights based on your feedback preferences">
                          <Box
                            sx={{
                              display: "flex",
                              alignItems: "center",
                              mr: 1,
                            }}
                          >
                            <Typography
                              variant="caption"
                              sx={{ mr: 0.5, color: "text.secondary" }}
                            >
                              Personalized
                            </Typography>
                            <Switch
                              checked={isPersonalizationEnabled}
                              onChange={(e) =>
                                setIsPersonalizationEnabled(e.target.checked)
                              }
                              color="primary"
                              size="small"
                              sx={{
                                "& .MuiSwitch-switchBase.Mui-checked": {
                                  color: (theme) =>
                                    theme.palette.mode === "dark"
                                      ? "#4ade80"
                                      : "#10a37f",
                                },
                                "& .MuiSwitch-switchBase.Mui-checked + .MuiSwitch-track":
                                  {
                                    backgroundColor: (theme) =>
                                      theme.palette.mode === "dark"
                                        ? "#4ade80"
                                        : "#10a37f",
                                  },
                              }}
                            />
                          </Box>
                        </Tooltip>
                      )}

                      {/* Show hidden points dialog button */}
                      {hiddenPointsCount > 0 && (
                        <Tooltip
                          title={`Show ${hiddenPointsCount} hidden insight${
                            hiddenPointsCount !== 1 ? "s" : ""
                          }`}
                        >
                          <IconButton
                            onClick={handleOpenHiddenPointsDialog}
                            size="small"
                            sx={{
                              color: (theme) =>
                                theme.palette.mode === "dark"
                                  ? "#4ade80"
                                  : "#10a37f",
                              bgcolor: (theme) =>
                                theme.palette.mode === "dark"
                                  ? "rgba(16, 163, 127, 0.1)"
                                  : "rgba(16, 163, 127, 0.05)",
                              "&:hover": {
                                bgcolor: (theme) =>
                                  theme.palette.mode === "dark"
                                    ? "rgba(16, 163, 127, 0.2)"
                                    : "rgba(16, 163, 127, 0.1)",
                              },
                            }}
                          >
                            <Badge
                              badgeContent={hiddenPointsCount}
                              color="error"
                              sx={{
                                "& .MuiBadge-badge": { fontSize: "0.6rem" },
                              }}
                            >
                              <RestoreIcon fontSize="small" />
                            </Badge>
                          </IconButton>
                        </Tooltip>
                      )}

                      {/* Deep Dive button */}
                      <Tooltip title="Deep Dive - View individual sheet insights">
                        <IconButton
                          onClick={() => setIsDeepDiveDialogOpen(true)}
                          size="small"
                          sx={{
                            color: (theme) =>
                              theme.palette.mode === "dark"
                                ? "#4ade80"
                                : "#10a37f",
                            bgcolor: (theme) =>
                              theme.palette.mode === "dark"
                                ? "rgba(16, 163, 127, 0.1)"
                                : "rgba(16, 163, 127, 0.05)",
                            "&:hover": {
                              bgcolor: (theme) =>
                                theme.palette.mode === "dark"
                                  ? "rgba(16, 163, 127, 0.2)"
                                  : "rgba(16, 163, 127, 0.1)",
                            },
                          }}
                        >
                          <TroubleshootIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>

                      {/* PDF download button */}
                      <Tooltip title="Download PDF">
                        <IconButton
                          onClick={handleDownloadPDF}
                          disabled={isPdfGenerating}
                          size="small"
                          sx={{
                            color: (theme) =>
                              theme.palette.mode === "dark"
                                ? "#4ade80"
                                : "#10a37f",
                            bgcolor: (theme) =>
                              theme.palette.mode === "dark"
                                ? "rgba(16, 163, 127, 0.1)"
                                : "rgba(16, 163, 127, 0.05)",
                            "&:hover": {
                              bgcolor: (theme) =>
                                theme.palette.mode === "dark"
                                  ? "rgba(16, 163, 127, 0.2)"
                                  : "rgba(16, 163, 127, 0.1)",
                            },
                          }}
                        >
                          {isPdfGenerating ? (
                            <CircularProgress
                              size={18}
                              sx={{ color: "inherit" }}
                            />
                          ) : (
                            <PictureAsPdfIcon fontSize="small" />
                          )}
                        </IconButton>
                      </Tooltip>

                      {/* Email button */}
                      <Tooltip title="Email Report">
                        <IconButton
                          onClick={handleOpenEmailDialog}
                          disabled={isPdfGenerating}
                          size="small"
                          sx={{
                            color: (theme) =>
                              theme.palette.mode === "dark"
                                ? "#4ade80"
                                : "#10a37f",
                            bgcolor: (theme) =>
                              theme.palette.mode === "dark"
                                ? "rgba(16, 163, 127, 0.1)"
                                : "rgba(16, 163, 127, 0.05)",
                            "&:hover": {
                              bgcolor: (theme) =>
                                theme.palette.mode === "dark"
                                  ? "rgba(16, 163, 127, 0.2)"
                                  : "rgba(16, 163, 127, 0.1)",
                            },
                          }}
                        >
                          <EmailIcon fontSize="small" />
                        </IconButton>
                      </Tooltip>
                    </Box>
                  </Box>
                )}

                {/* Title */}
                <Box
                  sx={{
                    display: "flex",
                    justifyContent: "flex-start",
                    alignItems: "center",
                    mb: 3,
                  }}
                >
                  <Typography
                    variant="h5"
                    sx={{
                      fontWeight: "bold",
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                    }}
                  >
                    AI Driven Insights
                  </Typography>
                </Box>

                <Box
                  sx={{
                    position: "relative",
                    bgcolor: (theme) =>
                      theme.palette.mode === "dark" ? "#121212" : "#ffffff",
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#e0e0e0" : "inherit",
                  }}
                >
                  {/* Parse and render the summary points with hover feedback */}
                  {summary && (
                    <Box>
                      {summary
                        .split("\n")
                        .filter((line) => line.trim())
                        .map((line, index, array) => {
                          // Extract point number from the line (assuming format: "1. Point text")
                          const match = line.match(/^(\d+)\./);
                          if (match) {
                            const pointId = `summary-${match[1]}`;
                            const isLastPoint =
                              index === array.length - 1 &&
                              additionalInsights.length === 0;

                            return (
                              <InsightPoint
                                key={pointId}
                                pointId={pointId}
                                content={line}
                                feedback={pointFeedback[pointId]}
                                onFeedback={handlePointFeedback}
                                isLastPoint={isLastPoint}
                                showMoreLink={isLastPoint}
                                onMoreClick={generateMoreInsights}
                                isGeneratingMore={isGeneratingMore}
                                onRemove={handleRemovePoint}
                                isHidden={hiddenPoints[pointId]}
                                showHiddenPoints={false}
                                onRestore={handleRestorePoint}
                              />
                            );
                          }
                          // For non-point lines, just render them normally
                          return <Box key={`other-${index}`}>{line}</Box>;
                        })}
                    </Box>
                  )}

                  {/* Render additional insights with hover feedback (no subheading) */}
                  {additionalInsights.length > 0 && (
                    <Box>
                      {additionalInsights.map(
                        (insightSet, setIndex, setArray) => (
                          <Box key={insightSet.id}>
                            {/* Show personalization indicator if insights are personalized */}
                            {insightSet.isPersonalized && (
                              <Box
                                sx={{
                                  display: "flex",
                                  alignItems: "center",
                                  mb: 1.5,
                                  mt: setIndex > 0 ? 2 : 0,
                                }}
                              >
                                <Chip
                                  icon={<AutoFixHighIcon />}
                                  label="Personalized Insights"
                                  size="small"
                                  sx={{
                                    bgcolor: (theme) =>
                                      theme.palette.mode === "dark"
                                        ? "rgba(16, 163, 127, 0.15)"
                                        : "#E7F7F1",
                                    color: (theme) =>
                                      theme.palette.mode === "dark"
                                        ? "#4ade80"
                                        : "#10a37f",
                                    border: (theme) =>
                                      `1px solid ${
                                        theme.palette.mode === "dark"
                                          ? "rgba(16, 163, 127, 0.3)"
                                          : "rgba(16, 163, 127, 0.2)"
                                      }`,
                                    "& .MuiChip-icon": {
                                      color: (theme) =>
                                        theme.palette.mode === "dark"
                                          ? "#4ade80"
                                          : "#10a37f",
                                    },
                                  }}
                                />
                                <Typography
                                  variant="caption"
                                  sx={{
                                    ml: 1,
                                    color: "text.secondary",
                                    fontStyle: "italic",
                                  }}
                                >
                                  Based on your feedback
                                </Typography>
                              </Box>
                            )}

                            {insightSet.content
                              .split("\n")
                              .filter((line) => line.trim())
                              .map((line, index, lineArray) => {
                                // Extract point number from the line
                                const match = line.match(/^(\d+)\./);
                                if (match) {
                                  const pointId = `additional-${match[1]}`;
                                  // Check if this is the last point of the last set
                                  const isLastPoint =
                                    setIndex === setArray.length - 1 &&
                                    index === lineArray.length - 1;

                                  return (
                                    <InsightPoint
                                      key={`${insightSet.id}-${index}`}
                                      pointId={pointId}
                                      content={line}
                                      feedback={pointFeedback[pointId]}
                                      onFeedback={handlePointFeedback}
                                      isLastPoint={isLastPoint}
                                      showMoreLink={isLastPoint}
                                      onMoreClick={generateMoreInsights}
                                      isGeneratingMore={isGeneratingMore}
                                      onRemove={handleRemovePoint}
                                      isHidden={hiddenPoints[pointId]}
                                      showHiddenPoints={false}
                                      onRestore={handleRestorePoint}
                                    />
                                  );
                                }
                                // For non-point lines, just render them normally (should be rare)
                                return (
                                  <Box key={`${insightSet.id}-other-${index}`}>
                                    {line}
                                  </Box>
                                );
                              })}
                          </Box>
                        )
                      )}
                    </Box>
                  )}

                  {/* "...more" link is now attached to the last point in the InsightPoint component */}
                </Box>
              </Box>
            </Box>
          </Box>

          {/* Footer */}
          <Box
            sx={{
              p: 2,
              bgcolor: (theme) =>
                theme.palette.mode === "dark" ? "#121212" : "#ffffff",
            }}
          >
            <Box
              sx={{
                maxWidth: 768,
                mx: "auto",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <Typography variant="caption" color="text.secondary">
                Data Insights • {new Date().getFullYear()}
              </Typography>

              <Paper
                elevation={0}
                {...getRootProps()}
                sx={{
                  border: "1px solid",
                  borderColor: (theme) =>
                    isUploading || isProcessing
                      ? theme.palette.mode === "dark"
                        ? "rgba(255,255,255,0.2)"
                        : "#ccc"
                      : theme.palette.mode === "dark"
                      ? "#4ade80"
                      : "#10a37f",
                  borderRadius: 2,
                  p: 1,
                  px: 2,
                  cursor:
                    isUploading || isProcessing ? "not-allowed" : "pointer",
                  bgcolor: (theme) =>
                    isUploading || isProcessing
                      ? theme.palette.mode === "dark"
                        ? "rgba(255,255,255,0.05)"
                        : "#f5f5f5"
                      : "transparent",
                  transition: "all 0.2s ease",
                  display: "flex",
                  alignItems: "center",
                  gap: 1,
                  minWidth: 160,
                  justifyContent: "center",
                  "&:hover": {
                    bgcolor: (theme) =>
                      isUploading || isProcessing
                        ? theme.palette.mode === "dark"
                          ? "rgba(255,255,255,0.05)"
                          : "#f5f5f5"
                        : theme.palette.mode === "dark"
                        ? "rgba(16, 163, 127, 0.15)"
                        : "rgba(16, 163, 127, 0.04)",
                  },
                }}
              >
                <input {...getInputProps()} />
                {isUploading ? (
                  <>
                    <CircularProgress
                      size={16}
                      sx={{
                        color: (theme) =>
                          theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                      }}
                    />
                    <Typography
                      variant="body2"
                      sx={{ color: "text.secondary" }}
                    >
                      Uploading...
                    </Typography>
                  </>
                ) : isProcessing ? (
                  <>
                    <CircularProgress
                      size={16}
                      sx={{
                        color: (theme) =>
                          theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                      }}
                    />
                    <Typography
                      variant="body2"
                      sx={{ color: "text.secondary" }}
                    >
                      Processing...
                    </Typography>
                  </>
                ) : (
                  <Typography
                    variant="body2"
                    sx={{
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                      fontWeight: "medium",
                    }}
                  >
                    Upload another file
                  </Typography>
                )}
              </Paper>
            </Box>
          </Box>
        </Box>
      )}

      {/* Email client selection dialog */}
      <EmailClientDialog
        open={isEmailDialogOpen}
        onClose={handleCloseEmailDialog}
        onSelect={handleSelectEmailClient}
      />

      {/* Hidden points dialog */}
      <HiddenPointsDialog
        open={isHiddenPointsDialogOpen}
        onClose={handleCloseHiddenPointsDialog}
        hiddenPoints={hiddenPoints}
        allInsights={{
          summary,
          additionalInsights,
        }}
        onRestore={handleRestorePoint}
        onRestoreAll={handleRestoreAllPoints}
      />

      {/* Snackbar for notifications */}
      <Snackbar
        open={snackbarOpen}
        autoHideDuration={6000}
        onClose={() => setSnackbarOpen(false)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert
          onClose={() => setSnackbarOpen(false)}
          severity={snackbarSeverity}
          sx={{ width: "100%" }}
        >
          {snackbarMessage}
        </Alert>
      </Snackbar>

      {/* Deep Dive Dialog */}
      <DeepDiveDialog
        open={isDeepDiveDialogOpen}
        onClose={() => setIsDeepDiveDialogOpen(false)}
      />
    </Box>
  );
};

export default MainContent;
