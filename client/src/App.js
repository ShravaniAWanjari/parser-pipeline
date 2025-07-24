import React from "react";
import { CssBaseline, Box } from "@mui/material";
import { createTheme, ThemeProvider } from "@mui/material/styles";
import { AppProvider } from "./context/AppContext";
import MainContent from "./components/MainContent";
import Header from "./components/Header";
import { useAppContext } from "./context/AppContext";
import "./animations.css";
import "./darkMode.css";

// Create theme based on mode (light/dark)
const createAppTheme = (mode) =>
  createTheme({
    palette: {
      mode,
      primary: {
        main: "#10a37f", // Primary green
        light: mode === "dark" ? "#4ade80" : "#34d399",
        dark: mode === "dark" ? "#059669" : "#047857",
        contrastText: "#ffffff",
      },
      secondary: {
        main: "#4285F4", // Blue
        light: "#60a5fa",
        dark: "#2563eb",
        contrastText: "#ffffff",
      },
      info: {
        main: "#FBBC05", // Yellow
        light: "#fcd34d",
        dark: "#d97706",
      },
      error: {
        main: "#EA4335", // Red
        light: "#f87171",
        dark: "#dc2626",
      },
      background: {
        default: mode === "dark" ? "#121212" : "#f8fafc",
        paper: mode === "dark" ? "#1e1e1e" : "#ffffff",
      },
      text: {
        primary: mode === "dark" ? "#f5f5f5" : "#334155",
        secondary: mode === "dark" ? "#aaaaaa" : "#64748b",
      },
      divider:
        mode === "dark" ? "rgba(255, 255, 255, 0.08)" : "rgba(0, 0, 0, 0.06)",
    },
    typography: {
      fontFamily: '"Inter", "Helvetica Neue", "Helvetica", "Arial", sans-serif',
      h1: {
        fontWeight: 700,
      },
      h2: {
        fontWeight: 700,
      },
      h3: {
        fontWeight: 600,
      },
      h4: {
        fontWeight: 600,
      },
      h5: {
        fontWeight: 600,
      },
      h6: {
        fontWeight: 600,
      },
      subtitle1: {
        fontWeight: 500,
      },
      subtitle2: {
        fontWeight: 500,
      },
      body1: {
        lineHeight: 1.6,
      },
      body2: {
        lineHeight: 1.6,
      },
    },
    shape: {
      borderRadius: 10,
    },
    components: {
      MuiPaper: {
        styleOverrides: {
          root: {
            boxShadow:
              mode === "dark"
                ? "0 4px 20px rgba(0, 0, 0, 0.3)"
                : "0 4px 20px rgba(0, 0, 0, 0.05)",
            borderRadius: 10,
          },
          elevation1: {
            boxShadow:
              mode === "dark"
                ? "0 2px 10px rgba(0, 0, 0, 0.2)"
                : "0 2px 10px rgba(0, 0, 0, 0.03)",
          },
        },
      },
      MuiAppBar: {
        styleOverrides: {
          root: {
            boxShadow:
              mode === "dark"
                ? "0 2px 10px rgba(0, 0, 0, 0.5)"
                : "0 2px 10px rgba(0, 0, 0, 0.03)",
          },
        },
      },
      MuiButton: {
        styleOverrides: {
          root: {
            textTransform: "none",
            fontWeight: 500,
            borderRadius: 8,
            padding: "8px 16px",
          },
          contained: {
            boxShadow: "none",
            "&:hover": {
              boxShadow: "0 4px 12px rgba(0, 0, 0, 0.15)",
            },
          },
        },
      },
      MuiTextField: {
        styleOverrides: {
          root: {
            "& .MuiOutlinedInput-root": {
              borderRadius: 8,
            },
          },
        },
      },
      MuiCard: {
        styleOverrides: {
          root: {
            borderRadius: 12,
            overflow: "hidden",
          },
        },
      },
      MuiCardContent: {
        styleOverrides: {
          root: {
            padding: 20,
          },
        },
      },
      MuiChip: {
        styleOverrides: {
          root: {
            borderRadius: 6,
          },
        },
      },
    },
  });

// App wrapper that uses context
function AppWithTheme() {
  const { darkMode } = useAppContext();
  const theme = createAppTheme(darkMode ? "dark" : "light");

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Box
        className={darkMode ? "dark" : ""}
        sx={{
          height: "100vh",
          display: "flex",
          flexDirection: "column",
          bgcolor: "background.default",
        }}
      >
        <Header />
        <Box sx={{ flexGrow: 1, overflow: "hidden" }}>
          <MainContent />
        </Box>
      </Box>
    </ThemeProvider>
  );
}

function App() {
  return (
    <AppProvider>
      <AppWithTheme />
    </AppProvider>
  );
}

export default App;
