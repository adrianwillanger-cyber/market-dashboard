# 📊 Market Dashboard

A web-based market dashboard that visualizes key market data through an interactive, browser-friendly interface. Built with HTML and powered by a Python data pipeline.

## 🔗 Live Demo

👉 [View the Dashboard](https://adrianwillanger-cyber.github.io/market-dashboard/)

## 📁 Project Structure

| File | Description |
|------|-------------|
| `index.html` | Main entry point for the dashboard |
| `market_dashboard.html` | Core dashboard layout and visualizations |
| `_js_chunk.js` | JavaScript modules for interactivity and chart rendering |
| `dashboard_data.json` | Primary dataset powering the dashboard |
| `dashboard_data2.json` | Supplementary dataset |
| `build_report.py` | Python script to process and generate dashboard data |

## 🚀 Getting Started

### View the Dashboard
Simply open `index.html` in any modern web browser, or visit the live demo link above.

### Rebuild the Data
To regenerate the dashboard data files:

```bash
python build_report.py
