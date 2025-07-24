import React, { createContext, useContext, useState, useEffect } from "react";

const AppContext = createContext();

export const useAppContext = () => useContext(AppContext);

export const AppProvider = ({ children }) => {
  // Theme state
  const [darkMode, setDarkMode] = useState(() => {
    // Check if user has a preference stored in localStorage
    const savedMode = localStorage.getItem("darkMode");
    return savedMode ? JSON.parse(savedMode) : false;
  });

  // Update localStorage when darkMode changes
  useEffect(() => {
    localStorage.setItem("darkMode", JSON.stringify(darkMode));
  }, [darkMode]);

  // Toggle dark mode
  const toggleDarkMode = () => {
    setDarkMode((prevMode) => !prevMode);
  };

  // File and processing state
  const [file, setFile] = useState(null);
  const [fileId, setFileId] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingError, setProcessingError] = useState(null);

  // Results state
  const [tables, setTables] = useState({});
  const [summary, setSummary] = useState("");
  const [markdown, setMarkdown] = useState("");
  const [downloadUrl, setDownloadUrl] = useState("");
  const [additionalInsights, setAdditionalInsights] = useState([]);
  const [isGeneratingMore, setIsGeneratingMore] = useState(false);



  // Feedback state
  const [pointFeedback, setPointFeedback] = useState(() => {
    // Load saved feedback from localStorage if available
    const savedFeedback = localStorage.getItem("pointFeedback");
    return savedFeedback ? JSON.parse(savedFeedback) : {};
  });

  const [hiddenPoints, setHiddenPoints] = useState(() => {
    // Load saved hidden points from localStorage if available
    const savedHiddenPoints = localStorage.getItem("hiddenPoints");
    return savedHiddenPoints ? JSON.parse(savedHiddenPoints) : {};
  });

  // Update localStorage when pointFeedback changes
  useEffect(() => {
    localStorage.setItem("pointFeedback", JSON.stringify(pointFeedback));
  }, [pointFeedback]);

  // Update localStorage when hiddenPoints changes
  useEffect(() => {
    localStorage.setItem("hiddenPoints", JSON.stringify(hiddenPoints));
  }, [hiddenPoints]);

  // Clear current state
  const resetState = () => {
    setFile(null);
    setFileId(null);
    setIsUploading(false);
    setUploadError(null);
    setIsProcessing(false);
    setProcessingError(null);
    setTables({});
    setSummary("");
    setMarkdown("");
    setDownloadUrl("");
    setAdditionalInsights([]);
    setIsGeneratingMore(false);
    setPointFeedback({});
    setHiddenPoints({});

    // Also clear localStorage to prevent persistence across sessions
    localStorage.removeItem("pointFeedback");
    localStorage.removeItem("hiddenPoints");
  };

  return (
    <AppContext.Provider
      value={{
        // Theme state
        darkMode,
        toggleDarkMode,

        // File state
        file,
        setFile,
        fileId,
        setFileId,
        isUploading,
        setIsUploading,
        uploadError,
        setUploadError,

        // Processing state
        isProcessing,
        setIsProcessing,
        processingError,
        setProcessingError,

        // Results state
        tables,
        setTables,
        summary,
        setSummary,
        markdown,
        setMarkdown,
        downloadUrl,
        setDownloadUrl,
        additionalInsights,
        setAdditionalInsights,
        isGeneratingMore,
        setIsGeneratingMore,

        // Feedback state
        pointFeedback,
        setPointFeedback,
        hiddenPoints,
        setHiddenPoints,

        // Actions
        resetState,
      }}
    >
      {children}
    </AppContext.Provider>
  );
};
