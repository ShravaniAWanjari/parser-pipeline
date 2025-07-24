import React, { useState } from "react";
import {
  Box,
  List,
  ListItem,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  Divider,
  Button,
  Typography,
  IconButton,
  TextField,
  Menu,
  MenuItem,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
} from "@mui/material";
import AddIcon from "@mui/icons-material/Add";
import ChatIcon from "@mui/icons-material/Chat";
import DeleteIcon from "@mui/icons-material/Delete";
import EditIcon from "@mui/icons-material/Edit";
import MoreVertIcon from "@mui/icons-material/MoreVert";
import { useAppContext } from "../context/AppContext";

const Sidebar = () => {
  const {
    chatSessions,
    activeChatId,
    setActiveChatId,
    createNewChat,
    deleteChat,
    updateChatTitle,
  } = useAppContext();

  // State for chat menu
  const [menuAnchorEl, setMenuAnchorEl] = useState(null);
  const [selectedChatId, setSelectedChatId] = useState(null);
  const isMenuOpen = Boolean(menuAnchorEl);

  // State for rename dialog
  const [renameDialogOpen, setRenameDialogOpen] = useState(false);
  const [newTitle, setNewTitle] = useState("");

  // State for delete confirmation dialog
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);

  // Handle opening the menu
  const handleMenuOpen = (event, chatId) => {
    event.stopPropagation();
    setMenuAnchorEl(event.currentTarget);
    setSelectedChatId(chatId);
  };

  // Handle closing the menu
  const handleMenuClose = () => {
    setMenuAnchorEl(null);
  };

  // Handle opening the rename dialog
  const handleRenameClick = () => {
    const chat = chatSessions.find((chat) => chat.id === selectedChatId);
    if (chat) {
      setNewTitle(chat.title);
      setRenameDialogOpen(true);
    }
    handleMenuClose();
  };

  // Handle saving the new title
  const handleRenameSave = () => {
    if (newTitle.trim()) {
      updateChatTitle(selectedChatId, newTitle.trim());
    }
    setRenameDialogOpen(false);
  };

  // Handle opening the delete confirmation dialog
  const handleDeleteClick = () => {
    setDeleteDialogOpen(true);
    handleMenuClose();
  };

  // Handle confirming deletion
  const handleDeleteConfirm = () => {
    deleteChat(selectedChatId);
    setDeleteDialogOpen(false);
  };

  return (
    <Box
      sx={{
        width: 250,
        height: "100vh",
        bgcolor: "#202123",
        color: "white",
        display: "flex",
        flexDirection: "column",
      }}
    >
      {/* New Chat Button */}
      <Box sx={{ p: 2 }}>
        <Button
          variant="outlined"
          fullWidth
          startIcon={<AddIcon />}
          onClick={createNewChat}
          sx={{
            color: "white",
            borderColor: "rgba(255,255,255,0.2)",
            "&:hover": {
              borderColor: "white",
              bgcolor: "rgba(255,255,255,0.1)",
            },
            justifyContent: "flex-start",
            textTransform: "none",
            fontWeight: "normal",
          }}
        >
          New chat
        </Button>
      </Box>

      <Divider sx={{ bgcolor: "rgba(255,255,255,0.1)" }} />

      {/* Chat Sessions List */}
      <Box sx={{ flexGrow: 1, overflow: "auto" }}>
        <List>
          {chatSessions.map((chat) => (
            <ListItem
              key={chat.id}
              disablePadding
              secondaryAction={
                <IconButton
                  edge="end"
                  size="small"
                  onClick={(e) => handleMenuOpen(e, chat.id)}
                  sx={{
                    color: "rgba(255,255,255,0.5)",
                    "&:hover": { color: "white" },
                    // Always show the three dots
                    opacity: { xs: 1, sm: 0.7 },
                    ".MuiListItemButton-root:hover &": {
                      opacity: 1,
                    },
                  }}
                >
                  <MoreVertIcon fontSize="small" />
                </IconButton>
              }
              sx={{
                ".MuiListItemSecondaryAction-root": {
                  right: 8,
                },
              }}
            >
              <ListItemButton
                selected={chat.id === activeChatId}
                onClick={() => setActiveChatId(chat.id)}
                sx={{
                  py: 1.5,
                  borderRadius: "4px",
                  mx: 0.5,
                  "&.Mui-selected": {
                    bgcolor: "#343541",
                    "&:hover": {
                      bgcolor: "#343541",
                    },
                    // Add a subtle left border to indicate active chat
                    borderLeft: "3px solid #10a37f",
                    pl: "13px", // Compensate for the border to keep text aligned
                  },
                  "&:hover": {
                    bgcolor: "rgba(255,255,255,0.1)",
                  },
                }}
              >
                <ListItemIcon
                  sx={{
                    minWidth: 36,
                    color: chat.id === activeChatId ? "#10a37f" : "white",
                  }}
                >
                  <ChatIcon fontSize="small" />
                </ListItemIcon>
                <ListItemText
                  primary={
                    <Typography
                      variant="body2"
                      noWrap
                      sx={{
                        textOverflow: "ellipsis",
                        overflow: "hidden",
                        whiteSpace: "nowrap",
                        fontWeight:
                          chat.id === activeChatId ? "medium" : "normal",
                      }}
                    >
                      {chat.title || "New chat"}
                    </Typography>
                  }
                />
              </ListItemButton>
            </ListItem>
          ))}
        </List>
      </Box>

      {/* Footer */}
      <Box sx={{ p: 2, borderTop: "1px solid rgba(255,255,255,0.1)" }}>
        <Typography variant="caption" sx={{ opacity: 0.7 }}>
          TowerLoan Data Insights
        </Typography>
      </Box>

      {/* Chat Options Menu */}
      <Menu
        anchorEl={menuAnchorEl}
        open={isMenuOpen}
        onClose={handleMenuClose}
        PaperProps={{
          sx: {
            bgcolor: "#202123",
            color: "white",
            border: "1px solid rgba(255,255,255,0.1)",
            boxShadow: "0 4px 20px rgba(0,0,0,0.4)",
            mt: 0.5,
            "& .MuiMenuItem-root": {
              py: 1.5,
              px: 2,
            },
          },
        }}
        transformOrigin={{ horizontal: "right", vertical: "top" }}
        anchorOrigin={{ horizontal: "right", vertical: "bottom" }}
      >
        <MenuItem onClick={handleRenameClick} sx={{ color: "white" }}>
          <EditIcon fontSize="small" sx={{ mr: 2 }} />
          Rename
        </MenuItem>
        <MenuItem onClick={handleDeleteClick} sx={{ color: "white" }}>
          <DeleteIcon fontSize="small" sx={{ mr: 2 }} />
          Delete
        </MenuItem>
      </Menu>

      {/* Rename Dialog */}
      <Dialog
        open={renameDialogOpen}
        onClose={() => setRenameDialogOpen(false)}
        PaperProps={{
          sx: {
            bgcolor: "#202123",
            color: "white",
            border: "1px solid rgba(255,255,255,0.1)",
          },
        }}
      >
        <DialogTitle>Rename chat</DialogTitle>
        <DialogContent>
          <TextField
            autoFocus
            margin="dense"
            label="Chat name"
            fullWidth
            variant="outlined"
            value={newTitle}
            onChange={(e) => setNewTitle(e.target.value)}
            sx={{
              "& .MuiOutlinedInput-root": {
                color: "white",
                "& fieldset": {
                  borderColor: "rgba(255,255,255,0.2)",
                },
                "&:hover fieldset": {
                  borderColor: "rgba(255,255,255,0.5)",
                },
                "&.Mui-focused fieldset": {
                  borderColor: "primary.main",
                },
              },
              "& .MuiInputLabel-root": {
                color: "rgba(255,255,255,0.7)",
              },
            }}
            InputLabelProps={{
              sx: { color: "rgba(255,255,255,0.7)" },
            }}
          />
        </DialogContent>
        <DialogActions>
          <Button
            onClick={() => setRenameDialogOpen(false)}
            sx={{ color: "white" }}
          >
            Cancel
          </Button>
          <Button onClick={handleRenameSave} sx={{ color: "primary.main" }}>
            Save
          </Button>
        </DialogActions>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <Dialog
        open={deleteDialogOpen}
        onClose={() => setDeleteDialogOpen(false)}
        PaperProps={{
          sx: {
            bgcolor: "#202123",
            color: "white",
            border: "1px solid rgba(255,255,255,0.1)",
          },
        }}
      >
        <DialogTitle>Delete chat</DialogTitle>
        <DialogContent>
          <Typography>
            Are you sure you want to delete this chat? This action cannot be
            undone.
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button
            onClick={() => setDeleteDialogOpen(false)}
            sx={{ color: "white" }}
          >
            Cancel
          </Button>
          <Button onClick={handleDeleteConfirm} color="error">
            Delete
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
};

export default Sidebar;
