#!/usr/bin/env python3
"""
Title and Abstract Screening Script for Systematic Reviews
This script automates the initial screening phase of systematic reviews.
"""

import pandas as pd
import re
from pathlib import Path
import os
from difflib import SequenceMatcher

def normalize_text(text):
    """Standardize text for comparison."""
    if not isinstance(text, str):
        return ""
    # Convert to lowercase and remove extra whitespace
    text = re.sub(r'\s+', ' ', text.lower().strip())
    # Remove punctuation
    text = re.sub(r'[^\w\s]', '', text)
    return text

def text_similarity(text1, text2):
    """Calculate similarity between two texts."""
    return SequenceMatcher(None, normalize_text(text1), normalize_text(text2)).ratio()

def parse_endnote_entries(content):
    """Parse EndNote text export into structured entries."""
    entries = []
    current_entry = {}
    
    lines = content.split('\n')
    current_tag = None
    
    for line in lines:
        line = line.strip()
        
        # Skip empty lines
        if not line:
            if current_entry:
                entries.append(current_entry)
                current_entry = {}
            continue
            
        # Check for new tag
        if re.match(r'^%[A-Z]', line):
            current_tag = line[:2]
            content = line[2:].strip()
        else:
            content = line
            
        if current_tag:
            if current_tag not in current_entry:
                current_entry[current_tag] = content
            else:
                current_entry[current_tag] += ' ' + content
                
    # Add last entry if exists
    if current_entry:
        entries.append(current_entry)
        
    return entries

def check_primary_condition(entry):
    """Check for primary condition terms."""
    terms = ['hfpef', 'heart failure with preserved ejection fraction', 'diastolic heart failure']
    text = (entry.get('%T', '') + ' ' + entry.get('%X', '')).lower()
    return any(term in text for term in terms)

def check_pathway_terms(entry):
    """Check for specific pathway/mechanism terms."""
    categories = {
        'metabolism': ['fatty acid', 'glucose', 'metabolic', 'oxidation'],
        'inflammation': ['inflammation', 'cytokine', 'immune'],
        'fibrosis': ['fibrosis', 'collagen', 'extracellular matrix']
    }
    
    text = (entry.get('%T', '') + ' ' + entry.get('%X', '')).lower()
    matches = {cat: sum(1 for term in terms if term in text)
              for cat, terms in categories.items()}
    return any(count >= 1 for count in matches.values())

def check_methodology(entry):
    """Check for specific methodological terms."""
    categories = {
        'clinical': ['patient', 'clinical trial', 'cohort'],
        'animal': ['mouse', 'rat', 'animal model'],
        'molecular': ['cell culture', 'protein expression', 'gene expression']
    }
    
    text = (entry.get('%T', '') + ' ' + entry.get('%X', '')).lower()
    matches = {cat: sum(1 for term in terms if term in text)
              for cat, terms in categories.items()}
    return any(count >= 1 for count in matches.values())

def filter_articles(entries):
    """Filter and categorize articles based on criteria."""
    categories = {
        'included': [],
        'systematic_reviews': [],
        'narrative_reviews': [],
        'duplicates': [],
        'unrelated': [],
        'methodology_lacking': []
    }
    
    # Check for duplicates first
    processed_titles = set()
    for entry in entries:
        title = entry.get('%T', '').lower()
        
        # Skip empty titles
        if not title:
            continue
            
        # Check for duplicates
        if title in processed_titles:
            categories['duplicates'].append(entry)
            continue
            
        processed_titles.add(title)
        
        # Check if it's a review
        if 'review' in title.lower():
            if 'systematic review' in title.lower():
                categories['systematic_reviews'].append(entry)
            else:
                categories['narrative_reviews'].append(entry)
            continue
            
        # Check primary condition
        if not check_primary_condition(entry):
            categories['unrelated'].append(entry)
            continue
            
        # Check pathway terms and methodology
        if check_pathway_terms(entry) and check_methodology(entry):
            categories['included'].append(entry)
        else:
            categories['methodology_lacking'].append(entry)
            
    return categories

def create_excel_report(categories, output_file='articles_classification.xlsx'):
    """Generate Excel report with results."""
    with pd.ExcelWriter(output_file) as writer:
        # Create summary sheet
        summary_data = {
            'Category': list(categories.keys()),
            'Count': [len(articles) for articles in categories.values()]
        }
        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
        
        # Create detailed sheets for each category
        for category, articles in categories.items():
            if not articles:  # Skip empty categories
                continue
                
            df_data = []
            for article in articles:
                df_data.append({
                    'Title': article.get('%T', ''),
                    'Authors': article.get('%A', ''),
                    'Year': article.get('%D', ''),
                    'Journal': article.get('%J', ''),
                    'Abstract': article.get('%X', '')
                })
                
            if df_data:  # Only create sheet if there's data
                pd.DataFrame(df_data).to_excel(writer, sheet_name=category.title(), index=False)

def main():
    """Main execution function."""
    try:
        # Get input file
        print("Choose input method:")
        print("1. Use file from current directory")
        print("2. Enter full file path")
        choice = input("Enter your choice (1 or 2): ")
        
        if choice == '1':
            filename = input("Enter the filename (including .txt extension): ")
            file_path = filename
        else:
            file_path = input("Enter the full file path: ")
            
        # Read and process file
        with open(file_path, 'r', encoding='utf-8') as file:
            content = file.read()
            
        # Parse entries
        entries = parse_endnote_entries(content)
        print(f"Found {len(entries)} entries to process")
        
        # Filter and categorize
        categories = filter_articles(entries)
        
        # Create report
        create_excel_report(categories)
        
        # Print summary
        print("\nProcessing complete! Summary of results:")
        for category, articles in categories.items():
            print(f"{category.title()}: {len(articles)} articles")
            
        print("\nDone! Check articles_classification.xlsx for detailed results.")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        raise

if __name__ == "__main__":
    main()