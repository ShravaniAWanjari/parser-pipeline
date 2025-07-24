import React from "react";
import {
  AppBar,
  Toolbar,
  Typography,
  IconButton,
  Box,
  useTheme,
  useMediaQuery,
  Tooltip,
} from "@mui/material";
import Brightness4Icon from "@mui/icons-material/Brightness4";
import Brightness7Icon from "@mui/icons-material/Brightness7";
import { useAppContext } from "../context/AppContext";
import "./Header.css";

const Header = () => {
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down("sm"));
  const {
    darkMode,
    toggleDarkMode,
    resetState,
  } = useAppContext();

  // Function to handle click on title and redirect to home page
  const handleTitleClick = () => {
    resetState();
  };

  return (
    <AppBar
      position="static"
      color="default"
      elevation={0}
      className="app-header"
    >
      <Toolbar className="header-toolbar">
        {/* Logo and Title */}
        <Box className="logo-container">
          <Tooltip title="Go to Home Page">
            <Typography
              variant={isMobile ? "h6" : "h5"}
              component="h1"
              onClick={handleTitleClick}
              className="app-title fade-in"
            >
              Data Insights
            </Typography>
          </Tooltip>
        </Box>

        {/* Right side actions */}
        <Box className="actions-container">
          {/* Dark mode toggle */}
          <Tooltip
            title={darkMode ? "Switch to light mode" : "Switch to dark mode"}
          >
            <IconButton
              onClick={toggleDarkMode}
              color="inherit"
              className="theme-toggle-button pulse-animation"
            >
              {darkMode ? <Brightness7Icon /> : <Brightness4Icon />}
            </IconButton>
          </Tooltip>
        </Box>
      </Toolbar>
    </AppBar>
  );
};

export default Header;
