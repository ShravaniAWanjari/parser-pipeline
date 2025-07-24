import React from "react";
import {
  Dialog,
  DialogTitle,
  DialogContent,
  Box,
  IconButton,
  Typography,
  Tooltip,
} from "@mui/material";
import GoogleIcon from "@mui/icons-material/Google";
import MicrosoftIcon from "@mui/icons-material/Microsoft";
import EmailIcon from "@mui/icons-material/Email";
import CloseIcon from "@mui/icons-material/Close";
import "./EmailClientDialog.css";

/**
 * Dialog for selecting an email client
 */
const EmailClientDialog = ({ open, onClose, onSelectClient }) => {
  // We don't need darkMode anymore since it's applied at the app level

  return (
    <Dialog
      open={open}
      onClose={onClose}
      maxWidth="sm"
      slotProps={{
        paper: {
          className: "email-dialog-paper",
        },
      }}
    >
      <DialogTitle className="email-dialog-title">
        <Box className="email-dialog-title-content">
          <EmailIcon className="email-icon" />
          <Typography variant="h6" sx={{ fontWeight: "medium" }}>
            Select Email Client
          </Typography>
        </Box>
        <IconButton onClick={onClose} size="small">
          <CloseIcon fontSize="small" />
        </IconButton>
      </DialogTitle>

      <DialogContent className="email-dialog-content">
        <Box className="email-clients-container">
          <Box className="email-client-option">
            <Tooltip title="Gmail">
              <IconButton
                onClick={() => onSelectClient("gmail")}
                className="gmail-button"
              >
                <GoogleIcon className="client-icon" />
              </IconButton>
            </Tooltip>
            <Typography variant="body1" className="client-label">
              Gmail
            </Typography>
          </Box>

          <Box className="email-client-option">
            <Tooltip title="Outlook Web">
              <IconButton
                onClick={() => onSelectClient("outlook")}
                className="outlook-button"
              >
                <MicrosoftIcon className="client-icon" />
              </IconButton>
            </Tooltip>
            <Typography variant="body1" className="client-label">
              Outlook Web
            </Typography>
          </Box>
        </Box>

        <Typography
          variant="caption"
          color="text.secondary"
          className="dialog-footer-text"
        >
          Select an email client to open. The PDF will be downloaded
          automatically.
        </Typography>
      </DialogContent>
    </Dialog>
  );
};

export default EmailClientDialog;
