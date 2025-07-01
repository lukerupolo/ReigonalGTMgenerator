#!/usr/bin/env python3
"""
ReigonalGTMgenerator.py - Regional GTM Data Analysis Tool
This file contains functionality for analyzing customer feedback data by region,
including sentiment analysis, topic classification, and pie chart visualizations.
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from typing import List, Dict, Any
import re
from collections import Counter

class RegionalGTMAnalyzer:
    """
    A class for analyzing regional GTM (Go-To-Market) data with sentiment analysis,
    topic classification, and visualization capabilities.
    """
    
    def __init__(self):
        self.df_result = None
        self.sentiment_keywords = {
            'positive': ['good', 'great', 'excellent', 'amazing', 'love', 'perfect', 'best', 'awesome', 'fantastic'],
            'negative': ['bad', 'terrible', 'awful', 'hate', 'worst', 'horrible', 'disappointing', 'poor'],
            'neutral': ['okay', 'fine', 'average', 'normal', 'standard']
        }
        
    def load_data(self, data: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Load and prepare data for analysis.
        Expected format: [{"region": "region_name", "comment": "text", ...}, ...]
        """
        self.df_result = pd.DataFrame(data)
        return self.df_result
    
    def classify_sentiment(self, comment: str) -> str:
        """
        Classify sentiment of a comment based on keyword matching.
        Returns: 'positive', 'negative', or 'neutral'
        """
        comment_lower = comment.lower()
        
        positive_count = sum(1 for word in self.sentiment_keywords['positive'] if word in comment_lower)
        negative_count = sum(1 for word in self.sentiment_keywords['negative'] if word in comment_lower)
        
        if positive_count > negative_count:
            return 'positive'
        elif negative_count > positive_count:
            return 'negative'
        else:
            return 'neutral'
    
    def classify_topic(self, comment: str, topics: List[str] = None) -> str:
        """
        Classify the topic/theme of a comment based on keyword matching.
        """
        if topics is None:
            topics = ['product', 'service', 'pricing', 'support', 'delivery', 'quality']
        
        comment_lower = comment.lower()
        topic_keywords = {
            'product': ['product', 'item', 'goods', 'feature', 'functionality'],
            'service': ['service', 'help', 'assistance', 'staff', 'team'],
            'pricing': ['price', 'cost', 'expensive', 'cheap', 'value', 'money'],
            'support': ['support', 'help desk', 'customer service', 'assistance'],
            'delivery': ['delivery', 'shipping', 'transport', 'arrival', 'logistics'],
            'quality': ['quality', 'durability', 'reliability', 'performance']
        }
        
        topic_scores = {}
        for topic in topics:
            if topic in topic_keywords:
                score = sum(1 for keyword in topic_keywords[topic] if keyword in comment_lower)
                topic_scores[topic] = score
        
        if topic_scores:
            return max(topic_scores.items(), key=lambda x: x[1])[0]
        else:
            return 'general'
    
    def analyze_comments(self, topics: List[str] = None) -> pd.DataFrame:
        """
        Perform sentiment analysis and topic classification on comments.
        """
        if self.df_result is None:
            raise ValueError("No data loaded. Please call load_data() first.")
        
        # Add sentiment column
        self.df_result['sentiment'] = self.df_result['comment'].apply(self.classify_sentiment)
        
        # Add topic column
        self.df_result['topic'] = self.df_result['comment'].apply(
            lambda x: self.classify_topic(x, topics)
        )
        
        return self.df_result
    
    def pie_chart_topic_coverage(self, region: str = None, save_path: str = None) -> None:
        """
        Create a pie chart showing topic coverage distribution.
        This function was supposedly removed but is being restored.
        """
        if self.df_result is None:
            raise ValueError("No data loaded. Please call load_data() and analyze_comments() first.")
        
        # Filter by region if specified
        data = self.df_result
        if region:
            data = data[data['region'] == region]
            title = f"Topic Coverage - {region}"
        else:
            title = "Topic Coverage - All Regions"
        
        # Count topics
        topic_counts = data['topic'].value_counts()
        
        # Create pie chart
        fig, ax = plt.subplots(figsize=(10, 8))
        colors = plt.cm.Set3(np.linspace(0, 1, len(topic_counts)))
        
        wedges, texts, autotexts = ax.pie(
            topic_counts.values, 
            labels=topic_counts.index,
            autopct='%1.1f%%',
            colors=colors,
            startangle=90
        )
        
        ax.set_title(title, fontsize=16, fontweight='bold')
        
        # Make percentage text more readable
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        plt.show()
    
    def sentiment_by_region_chart(self, save_path: str = None) -> None:
        """
        Create a bar chart showing sentiment distribution by region.
        """
        if self.df_result is None:
            raise ValueError("No data loaded. Please call analyze_comments() first.")
        
        # Create sentiment summary by region
        sentiment_summary = self.df_result.groupby(['region', 'sentiment']).size().unstack(fill_value=0)
        
        # Create bar chart
        fig, ax = plt.subplots(figsize=(12, 6))
        sentiment_summary.plot(kind='bar', ax=ax, color=['red', 'gray', 'green'])
        
        ax.set_title('Sentiment Analysis by Region', fontsize=16, fontweight='bold')
        ax.set_xlabel('Region', fontsize=12)
        ax.set_ylabel('Number of Comments', fontsize=12)
        ax.legend(title='Sentiment')
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
        
        plt.show()
    
    def generate_summary_report(self) -> Dict[str, Any]:
        """
        Generate a comprehensive summary report of the analysis.
        """
        if self.df_result is None:
            raise ValueError("No data loaded. Please call analyze_comments() first.")
        
        total_comments = len(self.df_result)
        regions = self.df_result['region'].unique()
        
        # Sentiment summary
        sentiment_summary = self.df_result['sentiment'].value_counts().to_dict()
        
        # Topic summary
        topic_summary = self.df_result['topic'].value_counts().to_dict()
        
        # Regional breakdown
        regional_breakdown = {}
        for region in regions:
            region_data = self.df_result[self.df_result['region'] == region]
            regional_breakdown[region] = {
                'total_comments': len(region_data),
                'sentiment_breakdown': region_data['sentiment'].value_counts().to_dict(),
                'topic_breakdown': region_data['topic'].value_counts().to_dict()
            }
        
        return {
            'total_comments': total_comments,
            'regions_analyzed': list(regions),
            'overall_sentiment': sentiment_summary,
            'overall_topics': topic_summary,
            'regional_breakdown': regional_breakdown
        }

def main():
    """
    Example usage of the RegionalGTMAnalyzer.
    """
    # Sample data
    sample_data = [
        {"region": "North America", "comment": "The product quality is excellent and support team is great"},
        {"region": "Europe", "comment": "Pricing seems too expensive for the value provided"},
        {"region": "Asia Pacific", "comment": "Delivery was slow but the product is good"},
        {"region": "North America", "comment": "Customer service was terrible and unhelpful"},
        {"region": "Europe", "comment": "Great product features and fast delivery"},
        {"region": "Asia Pacific", "comment": "Average quality, okay pricing"}
    ]
    
    # Initialize analyzer
    analyzer = RegionalGTMAnalyzer()
    
    # Load and analyze data
    analyzer.load_data(sample_data)
    df_result = analyzer.analyze_comments()
    
    print("Analysis Results:")
    print(df_result[['region', 'comment', 'sentiment', 'topic']])
    
    # Generate visualizations
    analyzer.pie_chart_topic_coverage()
    analyzer.sentiment_by_region_chart()
    
    # Generate summary report
    report = analyzer.generate_summary_report()
    print("\nSummary Report:")
    for key, value in report.items():
        print(f"{key}: {value}")

if __name__ == "__main__":
    main()