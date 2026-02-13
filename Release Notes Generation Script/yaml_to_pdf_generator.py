#!/usr/bin/env python3
"""
YAML to PDF Converter for CCDC Release Notes

This script reads release notes from a YAML file and converts them into a 
professionally formatted PDF document with NIH branding and styling.

Features:
- Converts YAML release notes to PDF format
- Includes NIH branding and proper formatting
- Handles HTML content from YAML fullText fields
- Professional layout with headers, sections, and styling
- Page numbering with total page count
- SVG logo support with proper aspect ratio maintenance
- Customizable PDF metadata (title, author, subject, creator)

Usage:
    python3 yaml_to_pdf_generator.py

Requirements:
    - site_announcement_log.yaml file in the same directory (or newsData.yaml for legacy format)
    - CCDC_Logo.svg file (optional, for logo display)
"""

import yaml
import os
import sys
from datetime import datetime
from io import BytesIO

# PDF generation libraries
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import Color, HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.lib import colors

# HTML parsing for content formatting
from bs4 import BeautifulSoup
import re

# SVG to PNG conversion
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing
import io

class ReleaseNotesPDFGenerator:
    def __init__(self, yaml_file_path, output_path=None, pdf_metadata=None):
        """
        Initialize the PDF generator with YAML file path.
        
        Args:
            yaml_file_path (str): Path to the YAML file containing release notes
            output_path (str): Path for the output PDF file (optional)
            pdf_metadata (dict): PDF metadata dictionary with keys:
                - Title: PDF title
                - Author: PDF author
                - Subject: PDF subject
                - Creator: Content creator
                - Producer: PDF producer (optional, defaults to 'ReportLab PDF Library')
        """
        self.yaml_file_path = yaml_file_path
        self.output_path = output_path or "CCDC_Release_Notes.pdf"
        self.release_notes = []
        self.total_pages = 0
        self.current_page = 0
        self.logo_drawing = None  # Cache for converted logo
        
        # Set default PDF metadata if not provided
        self.pdf_metadata = pdf_metadata or {
            'Title': 'CCDC Release Notes',
            'Author': 'National Cancer Institute',
            'Subject': 'CCDC Release Notes and Updates',
            'Creator': 'CCDC Release Notes Generator',
            'Producer': 'ReportLab PDF Library',
        }

        print(self.pdf_metadata)
        
        # Define colors based on the screenshot
        self.nih_blue = HexColor('#2f5496')
        self.nih_red = HexColor('#BA1F40')
        self.dark_gray = HexColor('#606061')
        self.light_gray = HexColor('#f5f5f5')
        
        # Initialize styles
        self.setup_styles()
        
    def setup_styles(self):
        """Set up paragraph styles for the PDF."""
        self.styles = getSampleStyleSheet()
        
        # Title style (Update Title)
        self.styles.add(ParagraphStyle(
            name='ReleaseTitle',
            parent=self.styles['Title'],
            fontSize=24,
            textColor=self.nih_blue,
            spaceAfter=12,
            alignment=TA_LEFT,
            fontName='Helvetica-Bold'
        ))
        
        # Date style (Date of Release)
        self.styles.add(ParagraphStyle(
            name='ReleaseDate',
            parent=self.styles['Normal'],
            fontSize=14,
            textColor=colors.black,
            spaceAfter=20,
            alignment=TA_LEFT,
            fontName='Helvetica'
        ))
        
        # Section header style
        self.styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=self.nih_blue,
            spaceBefore=20,
            spaceAfter=8,
            alignment=TA_LEFT,
            fontName='Helvetica-Bold'
        ))
        
        # Subsection header style
        self.styles.add(ParagraphStyle(
            name='SubsectionHeader',
            parent=self.styles['Heading2'],
            fontSize=13,
            textColor=self.nih_blue,
            spaceBefore=12,
            spaceAfter=6,
            alignment=TA_LEFT,
            fontName='Helvetica-Bold'
        ))
        
        # Normal paragraph style
        self.styles.add(ParagraphStyle(
            name='ReleaseContent',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            spaceAfter=4,
            alignment=TA_JUSTIFY,
            fontName='Helvetica',
            leftIndent=0,
            rightIndent=0
        ))
        
        # List item style
        self.styles.add(ParagraphStyle(
            name='ListItem',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            spaceAfter=2,
            alignment=TA_LEFT,
            fontName='Helvetica',
            leftIndent=20,
            bulletIndent=10
        ))
        
        # Nested list item style (for sub-items)
        self.styles.add(ParagraphStyle(
            name='NestedListItem',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            spaceAfter=2,
            alignment=TA_LEFT,
            fontName='Helvetica',
            leftIndent=40,
            bulletIndent=30
        ))
        
        # Deeply nested list item style (for third level and beyond)
        self.styles.add(ParagraphStyle(
            name='DeeplyNestedListItem',
            parent=self.styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            spaceAfter=2,
            alignment=TA_LEFT,
            fontName='Helvetica',
            leftIndent=60,
            bulletIndent=10
        ))
        
        # Footer style
        self.styles.add(ParagraphStyle(
            name='Footer',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))
        
        # Page number style
        self.styles.add(ParagraphStyle(
            name='PageNumber',
            parent=self.styles['Normal'],
            fontSize=9,
            textColor=colors.black,
            alignment=TA_RIGHT,
            fontName='Helvetica'
        ))

    def convert_date_format(self, date_str):
        """
        Convert date from YYYY-MM-DD format to "Month Date, Year" format.
        
        Args:
            date_str (str): Date in YYYY-MM-DD format
            
        Returns:
            str: Date in "Month Date, Year" format
        """
        try:
            # Parse the date
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            # Format as "Month Date, Year"
            return date_obj.strftime('%B %d, %Y')
        except Exception as e:
            print(f"Warning: Could not parse date '{date_str}': {e}")
            return date_str
    
    def load_yaml_data(self):
        """Load and parse the YAML file in site_announcement_log.yaml format."""
        try:
            with open(self.yaml_file_path, 'r', encoding='utf-8') as file:
                data = yaml.safe_load(file)
            
            # Check if it's the new format (list of lists) or old format (dict with releaseNotesList)
            if isinstance(data, list):
                # New format: site_announcement_log.yaml format
                # Structure: [Type, Title, Version, Date, Data Type, Highlight, Full Text]
                self.release_notes = []
                
                for entry in data:
                    # Skip non-list entries and entries that don't have enough fields
                    if not isinstance(entry, list) or len(entry) < 7:
                        continue
                    
                    # Skip if it's the header row (first item is "Type")
                    first_item = str(entry[0]).strip() if entry[0] else ""
                    if first_item == "Type":
                        continue
                    
                    # Extract fields: [Type, Title, Version, Date, Data Type, Highlight, Full Text]
                    type_val = entry[0] if len(entry) > 0 else ""
                    title = entry[1] if len(entry) > 1 else ""
                    version = entry[2] if len(entry) > 2 else ""
                    date_yyyy_mm_dd = entry[3] if len(entry) > 3 else ""
                    data_type = entry[4] if len(entry) > 4 else ""
                    highlight = entry[5] if len(entry) > 5 else ""
                    full_text = entry[6] if len(entry) > 6 else ""
                    
                    # Convert date format from YYYY-MM-DD to "Month Date, Year"
                    date_formatted = self.convert_date_format(date_yyyy_mm_dd) if date_yyyy_mm_dd else "Unknown Date"
                    
                    # Create release note dictionary
                    release_note = {
                        'type': type_val,
                        'title': title,
                        'version': version,
                        'date': date_formatted,
                        'dataType': data_type,
                        'slug': highlight,
                        'fullText': full_text,
                        'img': 'updateImgReleaseNotes'
                    }
                    
                    self.release_notes.append(release_note)
                
                print(f"Loaded {len(self.release_notes)} release notes entries from site_announcement_log.yaml format")
                
            elif isinstance(data, dict) and 'releaseNotesList' in data:
                # Old format: releaseNotesList structure
                self.release_notes = data['releaseNotesList']
                print(f"Loaded {len(self.release_notes)} release notes entries from releaseNotesList format")
            else:
                raise ValueError("YAML file format not recognized. Expected either a list of lists (site_announcement_log.yaml format) or a dict with 'releaseNotesList' key.")
                
        except Exception as e:
            print(f"Error loading YAML file: {e}")
            sys.exit(1)

    def convert_links_in_html(self, html_content, element):
        """
        Convert <a> tags in HTML content to ReportLab link format.
        
        Args:
            html_content (str): HTML content string
            element: BeautifulSoup element to search for links
            
        Returns:
            str: HTML content with links converted to ReportLab format
        """
        if not html_content or not element:
            return html_content
        
        # Find all <a> tags in the element
        links = element.find_all('a', recursive=True)
        
        if links:
            # Process links in reverse order to maintain positions
            for link in reversed(links):
                href = link.get('href', '')
                link_text = link.get_text().strip()
                if href and link_text:
                    # Create ReportLab link format (blue, underlined)
                    # Wrap link text in font tag for blue color, and underline for styling
                    # The link tag must be outermost for clickability
                    reportlab_link = f'<link href="{href}"><font color="blue"><u>{link_text}</u></font></link>'
                    # Replace the <a> tag with ReportLab link format
                    link_html = str(link)
                    if link_html in html_content:
                        html_content = html_content.replace(link_html, reportlab_link, 1)
            
            # Clean up any remaining spaces after links
            html_content = re.sub(r'(</link>)\s+', r'\1', html_content)
        
        return html_content

    def process_list_item(self, li, level=0, processed_elements=None):
        """
        Recursively process a list item and its nested lists.
        
        Args:
            li: BeautifulSoup list item element
            level: Current nesting level (0 = top level, 1 = nested, 2 = deeply nested, etc.)
            processed_elements: Set of processed element IDs
            
        Returns:
            list: List of ReportLab Paragraph elements
        """
        if processed_elements is None:
            processed_elements = set()
        
        elements = []
        
        # Mark this li and its descendants as processed
        for child in li.descendants:
            processed_elements.add(id(child))
        processed_elements.add(id(li))
        
        # Check if this li contains a nested ul
        nested_ul = li.find('ul')
        
        if nested_ul:
            # Mark the nested ul and all its descendants as processed
            for nested_desc in nested_ul.descendants:
                processed_elements.add(id(nested_desc))
            processed_elements.add(id(nested_ul))
            
            # Extract the main text (before the nested ul)
            p_tag = li.find('p')
            if p_tag:
                # Get the HTML content of the p tag
                p_html = str(p_tag)
                # Remove the nested ul from the HTML
                nested_ul_html = str(nested_ul)
                p_html = p_html.replace(nested_ul_html, '')
                
                # Convert links first, before removing other tags
                p_html = self.convert_links_in_html(p_html, p_tag)
                
                # Convert <strong> to <b> for ReportLab
                p_html = p_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                # Convert <em> to <i> for ReportLab
                p_html = p_html.replace('<em>', '<i>').replace('</em>', '</i>')
                # Remove other HTML tags but keep <b>, <i>, <link>, <font>, and <u> tags
                # Match tags that are NOT <b>, </b>, <i>, </i>, <link>, </link>, <u>, </u>, <font>, or </font>
                p_html = re.sub(r'<(?!/?(?:b|i|link|u|font)[\s>])[^>]+>', '', p_html)
                # Clean up extra spaces and remove symbol characters
                p_html = re.sub(r'\s+', ' ', p_html)
                p_html = p_html.replace('•', '').replace('·', '').strip()
                
                main_text = p_html
            else:
                # Fallback: extract text preserving formatting from direct children
                # Get HTML content before the nested ul
                li_html = str(li)
                nested_ul_html = str(nested_ul)
                # Find position of nested ul and extract everything before it
                ul_pos = li_html.find(nested_ul_html)
                if ul_pos > 0:
                    main_text_html = li_html[:ul_pos]
                    
                    # Convert links first, before removing other tags
                    main_text_html = self.convert_links_in_html(main_text_html, li)
                    
                    # Convert <strong> to <b> for ReportLab
                    main_text_html = main_text_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                    # Convert <em> to <i> for ReportLab
                    main_text_html = main_text_html.replace('<em>', '<i>').replace('</em>', '</i>')
                    # Remove <li> tag and other HTML tags but keep <b>, <i>, <link>, <font>, and <u> tags
                    main_text_html = re.sub(r'</?li[^>]*>', '', main_text_html)
                    main_text_html = re.sub(r'<(?!/?(?:b|i|link|u|font)[\s>])[^>]+>', '', main_text_html)
                    # Clean up extra spaces
                    main_text_html = re.sub(r'\s+', ' ', main_text_html)
                    main_text = main_text_html.strip()
                else:
                    # Fallback to plain text extraction
                    main_text_parts = []
                    for child in li.children:
                        if child == nested_ul:
                            break
                        if hasattr(child, 'get_text'):
                            text = child.get_text().strip()
                            if text:
                                main_text_parts.append(text)
                        elif hasattr(child, 'string') and child.string and child.string.strip():
                            main_text_parts.append(child.string.strip())
                    main_text = ' '.join(main_text_parts).strip()
            
            # Clean up extra spaces and remove empty content
            main_text = ' '.join(main_text.split())
            main_text = main_text.replace('&nbsp;', ' ').strip()
            
            if main_text:
                # Clean up the text - remove any symbol characters that might be present
                main_text = main_text.replace('•', '').replace('·', '').strip()
                
                # Choose style based on level
                if level == 0:
                    elements.append(Paragraph(f"• {main_text}", self.styles['ListItem']))
                elif level == 1:
                    elements.append(Paragraph(f"• {main_text}", self.styles['NestedListItem']))
                else:
                    # Level 2+ - use deeply nested style (more indentation)
                    elements.append(Paragraph(f"• {main_text}", self.styles['DeeplyNestedListItem']))
            
            # Recursively process nested list items
            for nested_li in nested_ul.find_all('li', recursive=False):
                nested_elements = self.process_list_item(nested_li, level + 1, processed_elements)
                elements.extend(nested_elements)
            
            # Extract text that comes AFTER the nested ul (within the same li)
            # This handles cases where there are additional items after third-level bullets
            after_text_parts = []
            found_nested_ul = False
            for child in li.children:
                if child == nested_ul:
                    found_nested_ul = True
                    continue
                if found_nested_ul:
                    # We're now processing content after the nested ul
                    if hasattr(child, 'get_text'):
                        text = child.get_text().strip()
                        if text:
                            after_text_parts.append(text)
                    elif hasattr(child, 'string') and child.string and child.string.strip():
                        after_text_parts.append(child.string.strip())
                    elif hasattr(child, 'name') and child.name == 'p':
                        # Extract text from p tag, preserving bold and italic formatting
                        p_html = str(child)
                        
                        # Convert links first, before removing other tags
                        p_html = self.convert_links_in_html(p_html, child)
                        
                        p_html = p_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                        p_html = p_html.replace('<em>', '<i>').replace('</em>', '</i>')
                        p_html = re.sub(r'<(?!/?(?:b|i|link|u|font)[\s>])[^>]+>', '', p_html)
                        p_html = re.sub(r'\s+', ' ', p_html)
                        text = p_html.strip()
                        if text:
                            after_text_parts.append(text)
            
            # Process any text that came after the nested ul
            if after_text_parts:
                after_text = ' '.join(after_text_parts).strip()
                after_text = after_text.replace('&nbsp;', ' ').strip()
                # Split by common separators to create separate items
                # Look for patterns like "Updated resource description" followed by "Updated resource and contact"
                # These are likely separate items that should be split
                if after_text:
                    # Try to split into separate items if there are multiple "Updated" statements
                    # Split on patterns like "Updated" at the start of a new phrase
                    items = re.split(r'(?=\bUpdated\b)', after_text)
                    for item in items:
                        item = item.strip()
                        if item:
                            # Clean up the text
                            item = item.replace('•', '').replace('·', '').strip()
                            if item:
                                # Add as a second-level item (same level as the one with nested items)
                                if level == 0:
                                    elements.append(Paragraph(f"• {item}", self.styles['NestedListItem']))
                                elif level == 1:
                                    elements.append(Paragraph(f"• {item}", self.styles['NestedListItem']))
                                else:
                                    elements.append(Paragraph(f"• {item}", self.styles['DeeplyNestedListItem']))
        else:
            # Regular li without nested ul - extract text preserving bold and italic formatting
            p_tag = li.find('p')
            if p_tag:
                # Get HTML and preserve <strong> and <i> tags
                p_html = str(p_tag)
                
                # Convert links first, before removing other tags
                p_html = self.convert_links_in_html(p_html, p_tag)
                
                # Convert <strong> to <b> for ReportLab
                p_html = p_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                # Convert <em> to <i> for ReportLab
                p_html = p_html.replace('<em>', '<i>').replace('</em>', '</i>')
                # Remove other HTML tags but keep <b>, <i>, <link>, and <u> tags
                p_html = re.sub(r'<(?!/?(?:b|i|link|u)[\s>])[^>]+>', '', p_html)
                # Clean up extra spaces
                p_html = re.sub(r'\s+', ' ', p_html)
                text = p_html.strip()
            else:
                # Fallback: check if li has direct children with formatting (like <span> tags)
                # Get the HTML content of the li element itself
                li_html = str(li)
                # Remove any nested <ul> tags first
                nested_ul = li.find('ul')
                if nested_ul:
                    nested_ul_html = str(nested_ul)
                    li_html = li_html.replace(nested_ul_html, '')
                
                # Convert links first, before removing other tags
                li_html = self.convert_links_in_html(li_html, li)
                
                # Convert <strong> to <b> for ReportLab
                li_html = li_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                # Convert <em> to <i> for ReportLab
                li_html = li_html.replace('<em>', '<i>').replace('</em>', '</i>')
                # Remove the <li> tag itself and other HTML tags but keep <b>, <i>, <link>, <font>, and <u> tags
                li_html = re.sub(r'</?li[^>]*>', '', li_html)  # Remove <li> tags
                li_html = re.sub(r'<(?!/?(?:b|i|link|u|font)[\s>])[^>]+>', '', li_html)  # Remove other tags except <b>, <i>, <link>, <font>, and <u>
                # Clean up extra spaces
                li_html = re.sub(r'\s+', ' ', li_html)
                text = li_html.strip()
                
                # If we still don't have text, fall back to plain text extraction
                if not text:
                    text = li.get_text().strip()
            
            if text:
                # Choose style based on level
                if level == 0:
                    elements.append(Paragraph(f"• {text}", self.styles['ListItem']))
                elif level == 1:
                    elements.append(Paragraph(f"• {text}", self.styles['NestedListItem']))
                else:
                    # Level 2+ - use deeply nested style
                    elements.append(Paragraph(f"• {text}", self.styles['DeeplyNestedListItem']))
        
        return elements

    def parse_html_content(self, html_content):
        """
        Parse HTML content and convert to ReportLab elements.
        
        Args:
            html_content (str): HTML content to parse
            
        Returns:
            list: List of ReportLab elements
        """
        elements = []
        
        if not html_content:
            return elements
            
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Track which elements we've already processed to avoid duplication
            processed_elements = set()
            
            for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'ul', 'li']):
                # Skip if already processed
                if id(element) in processed_elements:
                    continue
                
                if element.name == 'p':
                    # Skip paragraphs that are inside list items (they'll be handled by the ul/li processing)
                    if element.find_parent('li') is not None:
                        continue
                    
                    # Check if paragraph contains a nested ul - if so, exclude the ul content
                    nested_ul = element.find('ul')
                    if nested_ul:
                        # Extract HTML content (preserving formatting) from elements before the ul
                        # Build HTML from child elements before the ul to preserve formatting
                        html_parts = []
                        text_parts = []
                        for child in element.children:
                            if child == nested_ul:
                                break  # Stop at nested ul
                            # Get HTML representation of this child to preserve formatting
                            if hasattr(child, '__str__'):
                                child_html = str(child)
                                if child_html.strip():
                                    html_parts.append(child_html)
                            # Also get text for validation
                            if hasattr(child, 'get_text'):
                                child_text = child.get_text().strip()
                                if child_text:
                                    text_parts.append(child_text)
                            elif hasattr(child, 'string') and child.string:
                                text_parts.append(child.string.strip())
                        
                        if html_parts:
                            # Combine HTML parts and wrap in a p tag structure for processing
                            p_html = '<p>' + ' '.join(html_parts) + '</p>'
                            text = ' '.join(text_parts).strip()
                        else:
                            # Fallback: use text only
                            text = ' '.join(text_parts).strip()
                            p_html = None
                    else:
                        # No nested ul - preserve HTML formatting (italic, bold, links)
                        # Get the HTML content and convert to ReportLab format
                        p_html = str(element)
                        text = element.get_text().strip()
                    
                    # Clean up - remove any text that looks like it's from a list
                    # If text contains multiple resource names or dataset updates, it's probably duplicated content
                    # BUT: Skip this check for introductory paragraphs (those that come before the first heading)
                    if text:
                        # Check if this paragraph comes before any heading in the document
                        # Get all headings in document order
                        all_headings = soup.find_all(['h1', 'h2', 'h3'])
                        is_first_paragraph = True
                        if all_headings:
                            # Check if this paragraph appears before the first heading
                            first_heading = all_headings[0]
                            # Get all elements in document order
                            all_elements = soup.find_all(['p', 'h1', 'h2', 'h3'])
                            element_index = all_elements.index(element) if element in all_elements else -1
                            first_heading_index = all_elements.index(first_heading) if first_heading in all_elements else -1
                            if element_index >= 0 and first_heading_index >= 0:
                                is_first_paragraph = element_index < first_heading_index
                        
                        # Only skip if it's NOT the first paragraph and has multiple keywords
                        if not is_first_paragraph:
                            resource_keywords = ['Childhood Cancer Data Initiative', 'dbGaP', 'GENIE', 'Kids First Data Resource', 
                                                'Updated dataset', 'Moved dataset', 'Dataset', 'was replaced']
                            keyword_count = sum(1 for keyword in resource_keywords if keyword in text)
                            # If it has multiple keywords, it's likely duplicated list content - skip it
                            if keyword_count > 2:
                                processed_elements.add(id(element))
                                continue
                    
                    if text:
                        # Preserve HTML formatting while processing
                        # Start with HTML content if available, otherwise use plain text
                        if p_html:
                            # Convert HTML to ReportLab format
                            # Convert <strong> and <b> to <b> for ReportLab
                            final_text = p_html.replace('<strong>', '<b>').replace('</strong>', '</b>')
                            # Convert <em> to <i> for ReportLab
                            final_text = final_text.replace('<em>', '<i>').replace('</em>', '</i>')
                        else:
                            final_text = text
                        
                        # Convert links to clickable hyperlinks
                        # Find all <a> tags in the paragraph and convert them
                        links = element.find_all('a', recursive=True)
                        
                        if links:
                            # Process links in the HTML
                            for link in reversed(links):
                                href = link.get('href', '')
                                link_text = link.get_text().strip()
                                if href and link_text:
                                    # Create ReportLab link format (blue, underlined)
                                    # Wrap link text in font tag for blue color, and underline for styling
                                    # The link tag must be outermost for clickability
                                    reportlab_link = f'<link href="{href}"><font color="blue"><u>{link_text}</u></font></link>'
                                    # Replace the <a> tag with ReportLab link format
                                    link_html = str(link)
                                    if link_html in final_text:
                                        final_text = final_text.replace(link_html, reportlab_link, 1)
                            
                            # Clean up any remaining spaces after links
                            final_text = re.sub(r'(</link>)\s+', r'\1', final_text)
                        
                        # Remove remaining HTML tags except <b>, <i>, <link>, <font>, and <u>
                        # Keep <b>, <i>, <link>, <font>, and <u> tags, remove everything else
                        if p_html:
                            final_text = re.sub(r'<(?!/?(?:b|i|link|u|font)[\s>])[^>]+>', '', final_text)
                            # Clean up extra whitespace but preserve structure
                            final_text = re.sub(r'\s+', ' ', final_text).strip()
                        
                        # Check for inline styles
                        style = element.get('style', '')
                        if 'color: #2f5496' in style and 'font-size: 16pt' in style:
                            # This is a section header
                            elements.append(Paragraph(final_text, self.styles['SectionHeader']))
                        elif 'color: #2f5496' in style and 'font-size: 13pt' in style:
                            # This is a subsection header
                            elements.append(Paragraph(final_text, self.styles['SubsectionHeader']))
                        else:
                            elements.append(Paragraph(final_text, self.styles['ReleaseContent']))
                    processed_elements.add(id(element))
                
                elif element.name in ['h1', 'h2', 'h3']:
                    # Handle headers
                    text = element.get_text().strip()
                    if text:
                        if element.name == 'h1':
                            elements.append(Paragraph(text, self.styles['SectionHeader']))
                        else:
                            elements.append(Paragraph(text, self.styles['SubsectionHeader']))
                    processed_elements.add(id(element))
                
                elif element.name == 'ul':
                    # Skip if this ul is nested inside another ul (it will be handled by the parent ul)
                    if element.find_parent('ul') is not None:
                        continue
                    
                    # Handle unordered lists - process all ul elements but avoid duplication
                    for li in element.find_all('li', recursive=False):  # Only direct children
                        # Use recursive helper to process list items at any depth
                        li_elements = self.process_list_item(li, level=0, processed_elements=processed_elements)
                        elements.extend(li_elements)
                    
                    # Mark the ul and all its descendants as processed
                    for descendant in element.descendants:
                        processed_elements.add(id(descendant))
                    processed_elements.add(id(element))
                
                elif element.name == 'li':
                    # Skip list items that are inside a ul (they're already handled by ul processing)
                    if element.find_parent('ul') is not None:
                        continue
                    
                    # Handle individual list items (only if not inside ul)
                    text = element.get_text().strip()
                    if text:
                        elements.append(Paragraph(f"• {text}", self.styles['ListItem']))
                    processed_elements.add(id(element))
                        
        except Exception as e:
            print(f"Error parsing HTML content: {e}")
            # Fallback: treat as plain text
            plain_text = re.sub(r'<[^>]+>', '', html_content)
            if plain_text.strip():
                elements.append(Paragraph(plain_text.strip(), self.styles['ReleaseContent']))
        
        return elements

    def convert_svg_to_drawing(self, svg_path, target_height=50):
        """
        Convert SVG file to ReportLab Drawing object, maintaining original aspect ratio.
        Clips content to the viewBox to avoid rendering elements outside the visible area.
        
        Args:
            svg_path (str): Path to the SVG file
            target_height (int): Target height for the logo (width will be calculated to maintain aspect ratio)
            
        Returns:
            Drawing: ReportLab Drawing object or None if conversion fails
        """
        try:
            # Convert SVG to ReportLab Drawing
            drawing = svg2rlg(svg_path)
            
            if drawing:
                # Maintain original aspect ratio
                if drawing.width and drawing.height:
                    # Calculate scale factor based on target height
                    scale_factor = target_height / drawing.height
                    new_width = drawing.width * scale_factor
                    new_height = target_height
                    
                    # Apply scaling
                    drawing.scale(scale_factor, scale_factor)
                    drawing.width = new_width
                    drawing.height = new_height
                    
                    print(f"Successfully converted SVG to Drawing: {svg_path} (scaled to {new_width:.1f}x{new_height})")
                else:
                    # Fallback if dimensions not available
                    drawing.width = 400
                    drawing.height = target_height
                
                return drawing
            else:
                print(f"Failed to convert SVG: {svg_path}")
                return None
            
        except Exception as e:
            print(f"Error converting SVG to Drawing: {e}")
            return None

    def get_logo_drawing(self, target_height=50):
        """
        Get the logo drawing, converting from SVG if not already cached.
        
        Args:
            target_height (int): Target height for the logo (width will be calculated to maintain aspect ratio)
            
        Returns:
            Drawing: ReportLab Drawing object or None if conversion fails
        """
        if self.logo_drawing is None:
            svg_logo_path = os.path.join(os.path.dirname(__file__), 'CCDC_Logo.svg')
            if os.path.exists(svg_logo_path):
                self.logo_drawing = self.convert_svg_to_drawing(svg_logo_path, target_height)
        return self.logo_drawing

    def create_table_of_contents(self):
        """Create an interactive table of contents with clickable links"""
        from reportlab.platypus import Paragraph
        
        # Create header row - simple black and white styling
        toc_data = [
            [Paragraph('<b>Version</b>', self.styles['ListItem']), 
             Paragraph('<b>Date</b>', self.styles['ListItem'])]
        ]
        
        for i, note in enumerate(self.release_notes):
            version = note.get('version', 'N/A')
            date = note.get('date', 'Unknown Date')
            
            # Create clickable link on the version number
            anchor_name = f"release_{i}"
            clickable_version = Paragraph(f'<link href="#{anchor_name}" color="blue">{version}</link>', self.styles['ListItem'])
            
            # Create date cell as paragraph
            date_para = Paragraph(date, self.styles['ListItem'])
            
            toc_data.append([clickable_version, date_para])
        
        # Create table with simple black and white styling and cell borders
        toc_table = Table(toc_data, colWidths=[1.5*inch, 2.5*inch], hAlign='LEFT')
        toc_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ]))
        
        return toc_table

    def create_header_footer(self, canvas, doc):
        """
        Create header and footer for each page.
        
        Args:
            canvas: ReportLab canvas object
            doc: Document object
        """
        # Get page dimensions
        page_width, page_height = letter
        
        # Header with NIH logo or text
        try:
            # Try to use SVG logo first (with caching)
            drawing = self.get_logo_drawing(target_height=50)
            if drawing:
                # Render the drawing to the canvas
                renderPDF.draw(drawing, canvas, 50, page_height - 80)
            else:
                # Fallback to PNG logo if SVG not found
                png_logo_path = os.path.join(os.path.dirname(__file__), 'nih_logo.png')
                if os.path.exists(png_logo_path):
                    canvas.drawImage(png_logo_path, 50, page_height - 80, width=400, height=50)
                else:
                    raise Exception("No logo files found")
        except Exception as e:
            print(f"Warning: Could not load logo: {e}")
            # Draw text header as fallback
            canvas.setFont("Helvetica-Bold", 16)
            canvas.setFillColor(self.nih_blue)
            canvas.drawString(50, page_height - 60, "NATIONAL CANCER INSTITUTE")
        
        # Draw horizontal line under header
        canvas.setStrokeColor(self.nih_blue)
        canvas.setLineWidth(1)
        canvas.line(50, page_height - 90, page_width - 50, page_height - 90)
        
        # Footer
        footer_y = 50
        canvas.setFont("Helvetica", 9)
        canvas.setFillColor(colors.black)
        
        # Left footer text
        footer_text = "U.S. Department of Health and Human Services | National Institutes of Health | National Cancer Institute"
        canvas.drawString(50, footer_y, footer_text)
        
        # Right footer - page number
        page_num = canvas.getPageNumber()
        page_text = f"Page {page_num} of {self.total_pages}"
        text_width = canvas.stringWidth(page_text, "Helvetica", 9)
        canvas.drawString(page_width - 50 - text_width, footer_y, page_text)
        
        # Draw horizontal line above footer
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(0.5)
        canvas.line(50, footer_y + 15, page_width - 50, footer_y + 15)

    def build_story(self):
        """Build the story content (can be called multiple times)."""
        story = []
        
        # Add Table of Contents
        story.append(Paragraph("Table of Contents", self.styles['SectionHeader']))
        story.append(Spacer(1, 0.2*inch))
        
        # Create and add table of contents
        toc_table = self.create_table_of_contents()
        story.append(toc_table)
        story.append(PageBreak())
        
        # Process each release note
        for i, note in enumerate(self.release_notes):
            print(f"Processing release note {i+1}/{len(self.release_notes)}: {note.get('title', 'Unknown')}")
            
            # Add release title with anchor for TOC links
            title = note.get('title', 'Unknown Release')
            anchor_name = f"release_{i}"
            # Create a paragraph with an anchor that can be linked to from the TOC
            title_paragraph = Paragraph(f'<a name="{anchor_name}"></a>{title}', self.styles['ReleaseTitle'])
            story.append(title_paragraph)
            
            # Add release date
            date = note.get('date', 'Unknown Date')
            story.append(Paragraph(f"<b>DATE OF RELEASE:</b> {date.upper()}", self.styles['ReleaseDate']))
            
            # Add small spacing after date
            story.append(Spacer(1, 0.05*inch))
            
            # Add content
            full_text = note.get('fullText', '')
            if full_text:
                content_elements = self.parse_html_content(full_text)
                story.extend(content_elements)
            
            # Add page break between releases
            if i < len(self.release_notes) - 1:
                story.append(PageBreak())
        
        return story

    def generate_pdf(self):
        """Generate the PDF document."""
        print("Generating PDF...")
        
        # Two-pass approach to get accurate page count
        # First pass: Build to a temporary file to count pages
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_path = temp_file.name
        temp_file.close()
        
        temp_doc = SimpleDocTemplate(
            temp_path,
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=100,
            bottomMargin=80
        )
        
        # Track page count during first pass
        page_count = [0]  # Use list to allow modification in nested function
        
        def count_pages(canvas, doc):
            page_count[0] = canvas.getPageNumber()
        
        # Build story for first pass
        story = self.build_story()
        
        # Build temporary PDF to count pages
        temp_doc.build(story, onFirstPage=count_pages, onLaterPages=count_pages)
        
        # Get the actual page count
        self.total_pages = page_count[0]
        print(f"Total pages detected: {self.total_pages}")
        
        # Clean up temporary file
        try:
            os.remove(temp_path)
        except:
            pass
        
        # Second pass: Build the actual PDF with correct page count
        doc = SimpleDocTemplate(
            self.output_path,
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=100,
            bottomMargin=80
        )
        
        def add_header_footer(canvas, doc):
            # Set PDF metadata on the canvas
            canvas.setTitle(self.pdf_metadata.get('Title', ''))
            canvas.setAuthor(self.pdf_metadata.get('Author', ''))
            canvas.setSubject(self.pdf_metadata.get('Subject', ''))
            canvas.setCreator(self.pdf_metadata.get('Creator', ''))
            canvas.setProducer(self.pdf_metadata.get('Producer', 'ReportLab PDF Library'))
            if 'Keywords' in self.pdf_metadata:
                canvas.setKeywords(self.pdf_metadata['Keywords'])
            
            # Add header and footer
            self.create_header_footer(canvas, doc)
        
        # Rebuild story for second pass
        story = self.build_story()
        
        # Build the actual PDF
        try:
            doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)
            print(f"PDF generated successfully: {self.output_path}")
        except Exception as e:
            print(f"Error building PDF: {e}")
            raise

def main():
    """Main function to run the script."""
    # Get the directory of the script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Default file paths - try site_announcement_log.yaml first, then newsData.yaml
    yaml_file = os.path.join(script_dir, 'site_announcement_log.yaml')
    if not os.path.exists(yaml_file):
        # Fallback to legacy format
        yaml_file = os.path.join(script_dir, 'newsData.yaml')
    
    output_file = os.path.join(script_dir, 'CCDC_Release_Notes.pdf')
    
    # Check if YAML file exists
    if not os.path.exists(yaml_file):
        print(f"Error: YAML file not found. Tried:")
        print(f"  - {os.path.join(script_dir, 'site_announcement_log.yaml')}")
        print(f"  - {os.path.join(script_dir, 'newsData.yaml')}")
        sys.exit(1)
    
    # Create PDF generator
    generator = ReleaseNotesPDFGenerator(yaml_file, output_file)
    
    # Load data and generate PDF
    generator.load_yaml_data()
    generator.generate_pdf()
    
    print("Done!")

if __name__ == "__main__":
    main()
