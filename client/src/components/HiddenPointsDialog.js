import React from "react";
import {
  Dialog,
  DialogTitle,
  DialogContent,
  Button,
  Box,
  Typography,
  IconButton,
} from "@mui/material";
import CloseIcon from "@mui/icons-material/Close";
import RestoreIcon from "@mui/icons-material/Restore";

/**
 * Dialog component to display and manage hidden insight points
 */
const HiddenPointsDialog = ({
  open,
  onClose,
  hiddenPoints,
  allInsights,
  onRestore,
  onRestoreAll,
}) => {
  // Filter and prepare hidden points data
  const hiddenPointsData = React.useMemo(() => {
    const result = [];

    // Process all insights to find the hidden ones
    Object.entries(hiddenPoints).forEach(([pointId, isHidden]) => {
      if (!isHidden) return; // Skip points that aren't actually hidden

      // Find the content for this point ID
      const pointContent = findPointContent(pointId, allInsights);
      if (pointContent) {
        result.push({
          id: pointId,
          content: pointContent,
        });
      }
    });

    return result;
  }, [hiddenPoints, allInsights]);

  // Function to find the content of a point by its ID
  function findPointContent(pointId, allInsights) {
    // Extract the point number from the ID (e.g., "summary-1" -> "1")
    const match = pointId.match(/^(summary|additional)-(\d+)$/);
    if (!match) return null;

    const pointNumber = match[2];
    const pointType = match[1]; // "summary" or "additional"

    // Search in the appropriate section
    if (pointType === "summary" && allInsights.summary) {
      // Look for the point in the summary
      const lines = allInsights.summary.split("\n");
      for (const line of lines) {
        if (line.trim().startsWith(`${pointNumber}.`)) {
          return line.trim();
        }
      }
    } else if (pointType === "additional" && allInsights.additionalInsights) {
      // Look for the point in additional insights
      for (const insightSet of allInsights.additionalInsights) {
        const lines = insightSet.content.split("\n");
        for (const line of lines) {
          if (line.trim().startsWith(`${pointNumber}.`)) {
            return line.trim();
          }
        }
      }
    }

    return null;
  }

  // Handle restoring a point
  const handleRestore = (pointId) => {
    onRestore(pointId);
  };

  return (
    <Dialog
      open={open}
      onClose={onClose}
      maxWidth="md"
      fullWidth
      PaperProps={{
        sx: {
          borderRadius: 2,
          overflow: "hidden",
          width: "90%",
          maxWidth: 800,
        },
      }}
    >
      <DialogTitle
        sx={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          bgcolor: (theme) =>
            theme.palette.mode === "dark"
              ? "rgba(255,255,255,0.05)"
              : "rgba(0,0,0,0.02)",
          borderBottom: "1px solid",
          borderColor: "divider",
        }}
      >
        <Typography variant="h6" sx={{ fontWeight: 500 }}>
          Hidden Insights
        </Typography>
        <Box sx={{ display: "flex", alignItems: "center" }}>
          <Button
            size="small"
            startIcon={<RestoreIcon fontSize="small" />}
            onClick={() => onRestoreAll && onRestoreAll()}
            sx={{
              mr: 1.5,
              color: (theme) =>
                theme.palette.mode === "dark" ? "#60a5fa" : "#4285F4",
              fontSize: "0.8125rem",
            }}
          >
            Restore All
          </Button>
          <IconButton onClick={onClose} size="small">
            <CloseIcon fontSize="small" />
          </IconButton>
        </Box>
      </DialogTitle>

      <DialogContent sx={{ p: 2, mt: 0, pb: 2 }}>
        {hiddenPointsData.length === 0 ? (
          <Box sx={{ p: 3, textAlign: "center" }}>
            <Typography color="text.secondary">
              No hidden insights found.
            </Typography>
          </Box>
        ) : (
          <Box sx={{ p: 1 }}>
            {hiddenPointsData.map((point, index) => {
              // Extract the text from the content
              // Content might be in the format "1. Some insight text"
              const match = point.content.match(/^(\d+)\.\s+(.+)$/);
              let pointText = match ? match[2] : point.content;
              const pointNumber = match ? match[1] : "";

              // Check if the insight is negative by looking for the [NEG] prefix
              const isNegativeInsight = pointText.startsWith("[NEG]");

              // Remove the [NEG] prefix from the text if it exists
              if (isNegativeInsight) {
                pointText = pointText.replace("[NEG]", "").trim();
              }

              return (
                <Box
                  key={point.id}
                  sx={{
                    display: "flex",
                    alignItems: "flex-start",
                    mb: index === hiddenPointsData.length - 1 ? 0 : 0.75,
                    p: 1,
                    borderRadius: 1,
                    "&:hover": {
                      bgcolor: (theme) =>
                        theme.palette.mode === "dark"
                          ? "rgba(255,255,255,0.03)"
                          : "rgba(0,0,0,0.01)",
                    },
                  }}
                >
                  {/* Bullet point - red for negative insights, green for positive */}
                  <Box
                    component="span"
                    sx={{
                      display: "inline-block",
                      width: 8,
                      height: 8,
                      borderRadius: "50%",
                      mt: 1.5,
                      mr: 2,
                      bgcolor: isNegativeInsight
                        ? (theme) =>
                            theme.palette.mode === "dark"
                              ? "#f87171"
                              : "#EA4335"
                        : (theme) =>
                            theme.palette.mode === "dark"
                              ? "#4ade80"
                              : "#10a37f",
                    }}
                  />

                  {/* Point text */}
                  <Box sx={{ flex: 1 }}>
                    <Typography variant="body1" sx={{ fontSize: "1rem" }}>
                      {pointText}
                    </Typography>
                  </Box>

                  {/* Restore button */}
                  <Button
                    startIcon={<RestoreIcon />}
                    onClick={() => handleRestore(point.id)}
                    size="small"
                    sx={{
                      ml: 1,
                      color: (theme) =>
                        theme.palette.mode === "dark" ? "#60a5fa" : "#4285F4",
                      minWidth: 100,
                    }}
                  >
                    Restore
                  </Button>
                </Box>
              );
            })}
          </Box>
        )}
      </DialogContent>
    </Dialog>
  );
};

export default HiddenPointsDialog;
