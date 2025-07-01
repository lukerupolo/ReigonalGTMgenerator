# ReigonalGTMgenerator

This repository contains tools for regional Go-To-Market (GTM) analysis and presentation generation.

## Files

### `app.py` - GTM Presentation Generator
A Streamlit web application for creating regional GTM presentation decks. This tool:
- Uploads and analyzes global GTM presentations
- Generates regional market insights using OpenAI API
- Creates customized PowerPoint presentations for specific regions
- Includes timeline planning and activation planning features

**Usage:**
```bash
streamlit run app.py
```

### `ReigonalGTMgenerator.py` - Data Analysis Tool
A Python module for analyzing regional customer feedback data. This tool provides:
- Sentiment analysis of customer comments
- Topic classification functionality
- Pie chart visualization of topic coverage
- Regional sentiment distribution charts
- Comprehensive analysis reporting

**Usage:**
```python
from ReigonalGTMgenerator import RegionalGTMAnalyzer

analyzer = RegionalGTMAnalyzer()
analyzer.load_data(your_data)
df_result = analyzer.analyze_comments()
analyzer.pie_chart_topic_coverage()
```

## Installation

Install required dependencies:
```bash
pip install -r requirements.txt
```

## Purpose

These tools serve different but complementary purposes:
- `app.py`: For creating presentation materials for GTM strategies
- `ReigonalGTMgenerator.py`: For analyzing customer feedback and market data

Both tools support regional analysis but focus on different aspects of the GTM process.