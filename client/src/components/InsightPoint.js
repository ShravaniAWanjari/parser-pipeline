import React, { useState } from "react";
import {
  Box,
  IconButton,
  Tooltip,
  Typography,
  CircularProgress,
} from "@mui/material";
import ThumbUpIcon from "@mui/icons-material/ThumbUp";
import ThumbDownIcon from "@mui/icons-material/ThumbDown";
import CloseIcon from "@mui/icons-material/Close";
import RestoreIcon from "@mui/icons-material/Restore";
import "./InsightPoint.css";

const InsightPoint = ({
  content,
  pointId,
  feedback,
  onFeedback,
  // isLastPoint is not used but kept for API compatibility
  showMoreLink = false,
  onMoreClick,
  isGeneratingMore = false,
  onRemove,
  isHidden = false,
  showHiddenPoints = false,
  onRestore,
}) => {
  const [isHovered, setIsHovered] = useState(false);
  // We don't need darkMode anymore since it's applied at the app level

  // Extract the text from the content
  // Content might be in the format "1. Some insight text" or just plain text
  const match = content.match(/^(\d+)\.\s+(.+)$/);
  let pointText = match ? match[2] : content;

  // Check if the insight is negative by looking for the [NEG] prefix
  const isNegativeInsight = pointText.startsWith("[NEG]");

  // Remove the [NEG] prefix from the text if it exists
  if (isNegativeInsight) {
    pointText = pointText.replace("[NEG]", "").trim();
  }

  const handleMouseEnter = () => {
    setIsHovered(true);
  };

  const handleMouseLeave = () => {
    setIsHovered(false);
  };

  // If the point is hidden and we're not showing hidden points, don't render
  if (isHidden && !showHiddenPoints) {
    return null;
  }

  // Determine if this is a hidden point being shown
  const isHiddenButShown = isHidden && showHiddenPoints;

  return (
    <Box
      className={`insight-point ${isHiddenButShown ? "hidden-point" : ""}`}
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
    >
      {/* Bullet point - red for negative insights, green for positive */}
      <Box
        component="span"
        className={`bullet-point ${
          isNegativeInsight ? "negative" : "positive"
        } ${isHiddenButShown ? "faded" : ""}`}
      />

      {/* Text content with inline feedback buttons */}
      <Box className="insight-content">
        <Box sx={{ display: "inline" }}>
          {/* Point text */}
          <Box
            component="span"
            sx={{
              opacity: isHiddenButShown ? 0.6 : 1,
              textDecoration: isHiddenButShown ? "line-through" : "none",
            }}
          >
            {pointText}
          </Box>

          {/* Feedback buttons that appear on hover, inline with text */}
          {isHovered && (
            <Box component="span" className="feedback-buttons">
              {isHiddenButShown ? (
                <Tooltip title="Restore this insight">
                  <IconButton
                    onClick={() => onRestore && onRestore(pointId)}
                    size="small"
                    className="feedback-button restore"
                  >
                    <RestoreIcon sx={{ fontSize: "1.1rem" }} />
                  </IconButton>
                </Tooltip>
              ) : (
                <>
                  <Tooltip title="This insight is helpful">
                    <IconButton
                      onClick={() => onFeedback(pointId, true)}
                      size="small"
                      className={`feedback-button thumbs-up ${
                        feedback === true ? "active" : ""
                      }`}
                    >
                      <ThumbUpIcon sx={{ fontSize: "1.1rem" }} />
                    </IconButton>
                  </Tooltip>

                  <Tooltip title="This insight needs improvement">
                    <IconButton
                      onClick={() => onFeedback(pointId, false)}
                      size="small"
                      className={`feedback-button thumbs-down ${
                        feedback === false ? "active" : ""
                      }`}
                    >
                      <ThumbDownIcon sx={{ fontSize: "1.1rem" }} />
                    </IconButton>
                  </Tooltip>

                  <Tooltip title="Remove this insight">
                    <IconButton
                      onClick={() => onRemove && onRemove(pointId)}
                      size="small"
                      className="feedback-button remove"
                    >
                      <CloseIcon sx={{ fontSize: "1.1rem" }} />
                    </IconButton>
                  </Tooltip>
                </>
              )}
            </Box>
          )}

          {/* "...more" link attached to the last point - inline with text */}
          {showMoreLink && (
            <>
              {isGeneratingMore ? (
                <Box component="span" className="generating-indicator">
                  <CircularProgress size={12} className="progress-indicator" />
                  <Typography
                    component="span"
                    variant="body2"
                    color="text.secondary"
                    className="generating-text"
                  >
                    generating...
                  </Typography>
                </Box>
              ) : (
                <Typography
                  component="span"
                  variant="body2"
                  onClick={onMoreClick}
                  className="more-link"
                >
                  ...more
                </Typography>
              )}
            </>
          )}
        </Box>
      </Box>
    </Box>
  );
};

export default InsightPoint;
