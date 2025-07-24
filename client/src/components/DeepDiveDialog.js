import React, { useState, useEffect } from "react";
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Typography,
  Box,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  CircularProgress,
  Alert,
  Chip,
  IconButton,
  Tooltip,
  Card,
  CardContent,
  Divider,
  LinearProgress,
  Fade,
  Slide,
} from "@mui/material";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import CloseIcon from "@mui/icons-material/Close";
import InsightsIcon from "@mui/icons-material/Insights";
import TableChartIcon from "@mui/icons-material/TableChart";
import TrendingUpIcon from "@mui/icons-material/TrendingUp";
import AssessmentIcon from "@mui/icons-material/Assessment";
import api from "../services/api";

const DeepDiveDialog = ({ open, onClose }) => {
  const [sheetInsights, setSheetInsights] = useState({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [expandedPanels, setExpandedPanels] = useState({});

  // Load sheet insights when dialog opens
  useEffect(() => {
    if (open) {
      loadSheetInsights();
    }
  }, [open]);

  const loadSheetInsights = async () => {
    setLoading(true);
    setError(null);
    
    try {
      const response = await api.getSheetInsights();
      console.log("Sheet insights response:", response);
      
      if (response.sheet_insights) {
        setSheetInsights(response.sheet_insights);
        
        // Auto-expand the first panel
        const firstSheet = Object.keys(response.sheet_insights)[0];
        if (firstSheet) {
          setExpandedPanels({ [firstSheet]: true });
        }
      }
    } catch (error) {
      console.error("Error loading sheet insights:", error);
      setError(error.response?.data?.detail || "Failed to load sheet insights");
    } finally {
      setLoading(false);
    }
  };

  const handleAccordionChange = (sheetName) => (event, isExpanded) => {
    setExpandedPanels(prev => ({
      ...prev,
      [sheetName]: isExpanded
    }));
  };

  const handleClose = () => {
    setSheetInsights({});
    setExpandedPanels({});
    setError(null);
    onClose();
  };

  return (
    <Dialog
      open={open}
      onClose={handleClose}
      maxWidth="lg"
      fullWidth
      PaperProps={{
        sx: {
          borderRadius: 3,
          maxHeight: '85vh',
          background: (theme) =>
            theme.palette.mode === 'dark'
              ? 'linear-gradient(135deg, #1e1e1e 0%, #2d2d2d 100%)'
              : 'linear-gradient(135deg, #ffffff 0%, #f8fafc 100%)',
          boxShadow: (theme) =>
            theme.palette.mode === 'dark'
              ? '0 25px 50px -12px rgba(0, 0, 0, 0.8)'
              : '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
        }
      }}
      TransitionComponent={Slide}
      TransitionProps={{
        direction: "up",
        timeout: 400,
      }}
    >
      <DialogTitle
        sx={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          pb: 2,
          pt: 3,
          px: 3,
          background: (theme) =>
            theme.palette.mode === 'dark'
              ? 'linear-gradient(90deg, rgba(16, 163, 127, 0.1) 0%, rgba(74, 222, 128, 0.05) 100%)'
              : 'linear-gradient(90deg, rgba(16, 163, 127, 0.05) 0%, rgba(74, 222, 128, 0.02) 100%)',
          borderBottom: (theme) => `1px solid ${theme.palette.divider}`,
        }}
      >
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
          <Box
            sx={{
              p: 1.5,
              borderRadius: 2,
              background: (theme) =>
                theme.palette.mode === 'dark'
                  ? 'linear-gradient(135deg, #10a37f 0%, #4ade80 100%)'
                  : 'linear-gradient(135deg, #10a37f 0%, #34d399 100%)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            <TrendingUpIcon sx={{ color: 'white', fontSize: 24 }} />
          </Box>
          <Box>
            <Typography variant="h5" component="div" sx={{ fontWeight: 700, mb: 0.5 }}>
               Analysis
            </Typography>
            <Typography variant="body2" color="text.secondary">
              Explore detailed insights from each data sheet
            </Typography>
          </Box>
        </Box>
        <Tooltip title="Close">
          <IconButton
            onClick={handleClose}
            size="small"
            sx={{
              bgcolor: (theme) => theme.palette.action.hover,
              '&:hover': {
                bgcolor: (theme) => theme.palette.action.selected,
                transform: 'scale(1.1)',
              },
              transition: 'all 0.2s ease-in-out',
            }}
          >
            <CloseIcon />
          </IconButton>
        </Tooltip>
      </DialogTitle>

      <DialogContent
        dividers
        sx={{
          p: 0,
          bgcolor: (theme) => theme.palette.background.default,
        }}
      >
        {loading && (
          <Box sx={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            py: 8,
            gap: 3
          }}>
            <CircularProgress
              size={48}
              sx={{
                color: (theme) => theme.palette.primary.main,
              }}
            />
            <Typography variant="h6" color="text.secondary">
              Loading sheet insights...
            </Typography>
            <LinearProgress
              sx={{
                width: '200px',
                borderRadius: 1,
                height: 6,
              }}
            />
          </Box>
        )}

        {error && (
          <Box sx={{ p: 3 }}>
            <Alert
              severity="error"
              sx={{
                borderRadius: 2,
                '& .MuiAlert-icon': {
                  fontSize: '1.5rem',
                }
              }}
            >
              <Typography variant="subtitle2" sx={{ fontWeight: 600, mb: 0.5 }}>
                Failed to Load Insights
              </Typography>
              {error}
            </Alert>
          </Box>
        )}

        {!loading && !error && Object.keys(sheetInsights).length === 0 && (
          <Box sx={{ p: 4, textAlign: 'center' }}>
            <AssessmentIcon
              sx={{
                fontSize: 64,
                color: 'text.disabled',
                mb: 2
              }}
            />
            <Typography variant="h6" color="text.secondary" sx={{ mb: 1 }}>
              No Sheet Insights Available
            </Typography>
            <Typography variant="body2" color="text.disabled">
              Please upload and process an Excel file first to view detailed insights.
            </Typography>
          </Box>
        )}

        {!loading && !error && Object.keys(sheetInsights).length > 0 && (
          <Box sx={{ p: 3 }}>
            {/* Header Stats */}
            <Card
              sx={{
                mb: 3,
                background: (theme) =>
                  theme.palette.mode === 'dark'
                    ? 'linear-gradient(135deg, rgba(16, 163, 127, 0.1) 0%, rgba(74, 222, 128, 0.05) 100%)'
                    : 'linear-gradient(135deg, rgba(16, 163, 127, 0.05) 0%, rgba(74, 222, 128, 0.02) 100%)',
                border: (theme) => `1px solid ${theme.palette.divider}`,
                borderRadius: 2,
              }}
            >
              <CardContent sx={{ py: 2 }}>
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                    <TableChartIcon color="primary" />
                    <Typography variant="h6" sx={{ fontWeight: 600 }}>
                      {Object.keys(sheetInsights).length}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Data Sheets
                    </Typography>
                  </Box>
                  <Divider orientation="vertical" flexItem />
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                    <InsightsIcon color="primary" />
                    <Typography variant="h6" sx={{ fontWeight: 600 }}>
                      {Object.values(sheetInsights).reduce((total, insights) =>
                        total + (Array.isArray(insights) ? insights.length : 0), 0
                      )}
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Total Insights
                    </Typography>
                  </Box>
                </Box>
              </CardContent>
            </Card>

            {/* Sheet Insights */}
            <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
              {Object.entries(sheetInsights).map(([sheetName, insights], sheetIndex) => (
                <Fade
                  key={sheetName}
                  in={true}
                  timeout={300 + sheetIndex * 100}
                >
                  <Card
                    sx={{
                      borderRadius: 2,
                      overflow: 'hidden',
                      border: (theme) => `1px solid ${theme.palette.divider}`,
                      '&:hover': {
                        boxShadow: (theme) =>
                          theme.palette.mode === 'dark'
                            ? '0 8px 25px rgba(0, 0, 0, 0.3)'
                            : '0 8px 25px rgba(0, 0, 0, 0.1)',
                        transform: 'translateY(-2px)',
                      },
                      transition: 'all 0.3s ease-in-out',
                    }}
                  >
                    <Accordion
                      expanded={expandedPanels[sheetName] || false}
                      onChange={handleAccordionChange(sheetName)}
                      sx={{
                        boxShadow: 'none',
                        '&:before': { display: 'none' },
                        '& .MuiAccordionSummary-root': {
                          minHeight: 'auto',
                        }
                      }}
                    >
                      <AccordionSummary
                        expandIcon={
                          <ExpandMoreIcon
                            sx={{
                              color: 'primary.main',
                              transition: 'transform 0.3s ease',
                            }}
                          />
                        }
                        sx={{
                          background: (theme) =>
                            theme.palette.mode === 'dark'
                              ? 'rgba(255, 255, 255, 0.02)'
                              : 'rgba(0, 0, 0, 0.02)',
                          '&:hover': {
                            background: (theme) =>
                              theme.palette.mode === 'dark'
                                ? 'rgba(255, 255, 255, 0.05)'
                                : 'rgba(0, 0, 0, 0.04)',
                          },
                          py: 2,
                          px: 3,
                        }}
                      >
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 2, width: '100%' }}>
                          <Box
                            sx={{
                              width: 40,
                              height: 40,
                              borderRadius: 2,
                              background: (theme) =>
                                theme.palette.mode === 'dark'
                                  ? 'linear-gradient(135deg, #10a37f 0%, #4ade80 100%)'
                                  : 'linear-gradient(135deg, #10a37f 0%, #34d399 100%)',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              color: 'white',
                              fontWeight: 700,
                              fontSize: '1.1rem',
                            }}
                          >
                            {sheetIndex + 1}
                          </Box>
                          <Box sx={{ flex: 1 }}>
                            <Typography variant="h6" sx={{ fontWeight: 600, mb: 0.5 }}>
                              {sheetName}
                            </Typography>
                            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                              <Chip
                                label={`${Array.isArray(insights) ? insights.length : 0} insights`}
                                size="small"
                                color="primary"
                                variant="outlined"
                                sx={{
                                  borderRadius: 1,
                                  fontWeight: 500,
                                }}
                              />
                              <Typography variant="caption" color="text.secondary">
                                Click to explore
                              </Typography>
                            </Box>
                          </Box>
                        </Box>
                      </AccordionSummary>
                      <AccordionDetails sx={{ p: 0 }}>
                        <Box sx={{ p: 3, pt: 0 }}>
                          {Array.isArray(insights) ? (
                            <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1.5, mt: 2 }}>
                              {insights.map((insight, index) => (
                                <Fade
                                  key={index}
                                  in={expandedPanels[sheetName] || false}
                                  timeout={200 + index * 100}
                                >
                                  <Box sx={{ display: 'flex', gap: 2, alignItems: 'flex-start' }}>
                                    <Box
                                      sx={{
                                        width: 6,
                                        height: 6,
                                        borderRadius: '50%',
                                        background: (theme) =>
                                          theme.palette.mode === 'dark'
                                            ? 'linear-gradient(135deg, #10a37f 0%, #4ade80 100%)'
                                            : 'linear-gradient(135deg, #10a37f 0%, #34d399 100%)',
                                        flexShrink: 0,
                                        mt: 0.75, // Align with first line of text
                                      }}
                                    />
                                    <Typography
                                      variant="body1"
                                      sx={{
                                        lineHeight: 1.6,
                                        flex: 1,
                                        color: 'text.primary',
                                      }}
                                    >
                                      {insight}
                                    </Typography>
                                  </Box>
                                </Fade>
                              ))}
                            </Box>
                          ) : (
                            <Box sx={{ textAlign: 'center', py: 3 }}>
                              <Typography variant="body2" color="text.secondary">
                                No insights available for this sheet.
                              </Typography>
                            </Box>
                          )}
                        </Box>
                      </AccordionDetails>
                    </Accordion>
                  </Card>
                </Fade>
              ))}
            </Box>
          </Box>
        )}
      </DialogContent>

      <DialogActions
        sx={{
          px: 3,
          py: 2.5,
          background: (theme) =>
            theme.palette.mode === 'dark'
              ? 'rgba(255, 255, 255, 0.02)'
              : 'rgba(0, 0, 0, 0.02)',
          borderTop: (theme) => `1px solid ${theme.palette.divider}`,
        }}
      >
        <Button
          onClick={handleClose}
          variant="contained"
          size="large"
          sx={{
            borderRadius: 2,
            px: 4,
            py: 1,
            background: (theme) =>
              theme.palette.mode === 'dark'
                ? 'linear-gradient(135deg, #10a37f 0%, #4ade80 100%)'
                : 'linear-gradient(135deg, #10a37f 0%, #34d399 100%)',
            '&:hover': {
              background: (theme) =>
                theme.palette.mode === 'dark'
                  ? 'linear-gradient(135deg, #059669 0%, #22c55e 100%)'
                  : 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
              transform: 'translateY(-1px)',
              boxShadow: (theme) =>
                theme.palette.mode === 'dark'
                  ? '0 8px 25px rgba(16, 163, 127, 0.3)'
                  : '0 8px 25px rgba(16, 163, 127, 0.2)',
            },
            transition: 'all 0.2s ease-in-out',
          }}
        >
          Close
        </Button>
      </DialogActions>
    </Dialog>
  );
};

export default DeepDiveDialog;
