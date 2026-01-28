#!/usr/bin/env python3
"""
Simple test script to verify YAML parsing from site_announcement_log.yaml
This doesn't require PDF generation libraries.
"""

import yaml
import os
from datetime import datetime

def convert_date_format(date_str):
    """Convert date from YYYY-MM-DD format to "Month Date, Year" format."""
    try:
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        return date_obj.strftime('%B %d, %Y')
    except Exception as e:
        print(f"Warning: Could not parse date '{date_str}': {e}")
        return date_str

def load_yaml_data(yaml_file_path):
    """Load and parse the YAML file in site_announcement_log.yaml format."""
    try:
        with open(yaml_file_path, 'r', encoding='utf-8') as file:
            data = yaml.safe_load(file)
        
        release_notes = []
        
        if isinstance(data, list):
            # New format: site_announcement_log.yaml format
            # Structure: [Type, Title, Version, Date, Data Type, Highlight, Full Text]
            
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
                date_formatted = convert_date_format(date_yyyy_mm_dd) if date_yyyy_mm_dd else "Unknown Date"
                
                # Create release note dictionary
                release_note = {
                    'type': type_val,
                    'title': title,
                    'version': version,
                    'date': date_formatted,
                    'dataType': data_type,
                    'slug': highlight,
                    'fullText': full_text[:100] + "..." if len(full_text) > 100 else full_text,  # Truncate for display
                    'img': 'updateImgReleaseNotes'
                }
                
                release_notes.append(release_note)
            
            print(f"✓ Successfully loaded {len(release_notes)} release notes entries from site_announcement_log.yaml format\n")
            
            # Display first few entries
            print("Sample entries:")
            print("=" * 80)
            for i, note in enumerate(release_notes[:3], 1):
                print(f"\nEntry {i}:")
                print(f"  Title: {note['title']}")
                print(f"  Version: {note['version']}")
                print(f"  Date: {note['date']}")
                print(f"  Data Type: {note['dataType']}")
                print(f"  Highlight: {note['slug']}")
                print(f"  Full Text (preview): {note['fullText'][:150]}...")
            
            if len(release_notes) > 3:
                print(f"\n... and {len(release_notes) - 3} more entries")
            
            return release_notes
            
        elif isinstance(data, dict) and 'releaseNotesList' in data:
            # Old format: releaseNotesList structure
            release_notes = data['releaseNotesList']
            print(f"✓ Loaded {len(release_notes)} release notes entries from releaseNotesList format")
            return release_notes
        else:
            raise ValueError("YAML file format not recognized. Expected either a list of lists (site_announcement_log.yaml format) or a dict with 'releaseNotesList' key.")
            
    except Exception as e:
        print(f"✗ Error loading YAML file: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    yaml_file = os.path.join(script_dir, 'site_announcement_log.yaml')
    
    if not os.path.exists(yaml_file):
        print(f"✗ Error: YAML file not found at {yaml_file}")
        exit(1)
    
    print(f"Reading from: {yaml_file}\n")
    release_notes = load_yaml_data(yaml_file)
    
    if release_notes:
        print(f"\n✓ Success! Parsed {len(release_notes)} release notes successfully.")
    else:
        print("\n✗ Failed to parse release notes.")
        exit(1)

