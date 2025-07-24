import React, { useState } from "react";
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Button,
  Box,
  Typography,
  CircularProgress,
  Alert,
  Checkbox,
  FormControlLabel,
  Divider,
} from "@mui/material";
import EmailIcon from "@mui/icons-material/Email";
import CloseIcon from "@mui/icons-material/Close";
import gmailAttachImage from "../assets/gmail-attach.svg";

/**
 * Email form component for sending insights via email
 */
const EmailForm = ({
  open,
  onClose,
  onSend,
  loading = false,
  error = null,
  success = false,
}) => {
  // Form state
  const [email, setEmail] = useState("");
  const [subject, setSubject] = useState("Data Insights Report");
  const [message, setMessage] = useState(
    "Please find attached the data insights report."
  );
  const [includeFeedback, setIncludeFeedback] = useState(true);

  // Email validation
  const isValidEmail = (email) => {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  };

  // Form validation
  const isFormValid = email && isValidEmail(email) && subject && message;

  // Handle form submission
  const handleSubmit = (e) => {
    e.preventDefault();
    if (isFormValid) {
      onSend({
        email,
        subject,
        message,
        includeFeedback,
      });
    }
  };

  // Reset form when dialog is closed
  const handleClose = () => {
    if (!loading) {
      onClose();
      // Reset form after a short delay to avoid visual glitches
      setTimeout(() => {
        setEmail("");
        setSubject("Data Insights Report");
        setMessage("Please find attached the data insights report.");
        setIncludeFeedback(true);
      }, 300);
    }
  };

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      maxWidth="sm"
      fullWidth
      PaperProps={{
        sx: {
          borderRadius: 2,
          boxShadow: "0 8px 32px rgba(0, 0, 0, 0.1)",
        },
      }}
    >
      <DialogTitle
        sx={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          borderBottom: "1px solid",
          borderColor: "divider",
          pb: 1,
        }}
      >
        <Box sx={{ display: "flex", alignItems: "center" }}>
          <EmailIcon
            sx={{
              mr: 1,
              color: (theme) =>
                theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
            }}
          />
          <Typography variant="h6">Send Insights via Email</Typography>
        </Box>
        <Button
          onClick={handleClose}
          disabled={loading}
          sx={{ minWidth: "auto", p: 0.5 }}
        >
          <CloseIcon />
        </Button>
      </DialogTitle>

      <DialogContent sx={{ pt: 2 }}>
        <form onSubmit={handleSubmit}>
          <Alert severity="info" sx={{ mb: 2, mt: 1 }}>
            <Typography variant="body2" sx={{ fontWeight: "medium", mb: 1 }}>
              Quick Email Process:
            </Typography>
            <ol style={{ margin: 0, paddingLeft: "1.5rem" }}>
              <li>
                The PDF report will be automatically downloaded to your computer
              </li>
              <li>Gmail will open in a new tab with your message pre-filled</li>
              <li>
                Click the attachment icon in Gmail and select the downloaded PDF
                file
              </li>
              <li>Send your email as usual</li>
            </ol>
          </Alert>

          {error && (
            <Alert severity="error" sx={{ mb: 2, mt: 1 }}>
              {error}
            </Alert>
          )}

          {success && (
            <Alert severity="success" sx={{ mb: 2, mt: 1 }}>
              Success! PDF downloaded and Gmail opened. Please attach the PDF
              file to complete your email.
            </Alert>
          )}

          <TextField
            label="Recipient Email"
            type="email"
            fullWidth
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            margin="normal"
            required
            error={email && !isValidEmail(email)}
            helperText={
              email && !isValidEmail(email)
                ? "Please enter a valid email address"
                : ""
            }
            disabled={loading || success}
            autoFocus
          />

          <TextField
            label="Subject"
            fullWidth
            value={subject}
            onChange={(e) => setSubject(e.target.value)}
            margin="normal"
            required
            disabled={loading || success}
          />

          <TextField
            label="Message"
            fullWidth
            value={message}
            onChange={(e) => setMessage(e.target.value)}
            margin="normal"
            multiline
            rows={4}
            required
            disabled={loading || success}
          />

          <FormControlLabel
            control={
              <Checkbox
                checked={includeFeedback}
                onChange={(e) => setIncludeFeedback(e.target.checked)}
                disabled={loading || success}
                sx={{
                  color: (theme) =>
                    theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                  "&.Mui-checked": {
                    color: (theme) =>
                      theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
                  },
                }}
              />
            }
            label="Include feedback in the report"
            sx={{ mt: 1 }}
          />

          <Divider sx={{ my: 2 }} />

          <Typography variant="subtitle2" sx={{ mb: 1, fontWeight: "medium" }}>
            How to attach the PDF in Gmail:
          </Typography>

          <Box sx={{ display: "flex", justifyContent: "center", mb: 1 }}>
            <img
              src={gmailAttachImage}
              alt="Gmail attachment guide"
              style={{
                maxWidth: "100%",
                height: "auto",
                border: "1px solid #e0e0e0",
                borderRadius: "4px",
              }}
            />
          </Box>
        </form>
      </DialogContent>

      <DialogActions
        sx={{ px: 3, py: 2, borderTop: "1px solid", borderColor: "divider" }}
      >
        <Button
          onClick={handleClose}
          disabled={loading}
          variant="outlined"
          sx={{
            borderColor: (theme) =>
              theme.palette.mode === "dark"
                ? "rgba(255,255,255,0.2)"
                : "rgba(0,0,0,0.2)",
            color: "text.secondary",
            "&:hover": {
              borderColor: (theme) =>
                theme.palette.mode === "dark"
                  ? "rgba(255,255,255,0.3)"
                  : "rgba(0,0,0,0.3)",
              backgroundColor: "transparent",
            },
          }}
        >
          Cancel
        </Button>
        <Button
          onClick={handleSubmit}
          disabled={!isFormValid || loading || success}
          variant="contained"
          startIcon={
            loading ? (
              <CircularProgress size={20} color="inherit" />
            ) : (
              <EmailIcon />
            )
          }
          sx={{
            bgcolor: (theme) =>
              theme.palette.mode === "dark" ? "#4ade80" : "#10a37f",
            "&:hover": {
              bgcolor: (theme) =>
                theme.palette.mode === "dark" ? "#3bca6d" : "#0e8c6d",
            },
            "&.Mui-disabled": {
              bgcolor: (theme) =>
                theme.palette.mode === "dark"
                  ? "rgba(74, 222, 128, 0.3)"
                  : "rgba(16, 163, 127, 0.3)",
            },
          }}
        >
          {loading ? "Preparing..." : "Open in Gmail"}
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default EmailForm;
