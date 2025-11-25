"""
Configuration-based Excel Level Processor
Allows easy configuration of file paths and column mappings.
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
import re
import json
from datetime import datetime

# Default configuration
DEFAULT_CONFIG = {
    "files": {
        "aptive_file": "Aptive_BOOM_v01.xlsb",
        "bom_template": "BOM_Template.xlsx",
        "output_prefix": "BOM_Processed"
    },
    "column_mapping": {
        "aptive_source_columns": [0, 1, 3, 4],  # A, B, D, E (0-indexed)
        "bom_target_columns": ["E", "N", "O", "P"]
    },
    "settings": {
        "start_row": 2,
        "skip_header": True,
        "add_blank_between_levels": True
    }
}

class ConfigurableExcelProcessor:
    def __init__(self, config_file=None, workspace_path=""):
        """Initialize with configuration"""
        self.workspace_path = workspace_path or os.getcwd()
        self.config = self.load_config(config_file)
        
    def load_config(self, config_file):
        """Load configuration from file or use defaults"""
        if config_file and os.path.exists(config_file):
            with open(config_file, 'r') as f:
                config = json.load(f)
        else:
            config = DEFAULT_CONFIG.copy()
            
        return config
    
    def get_file_path(self, filename):
        """Get full file path"""
        return os.path.join(self.workspace_path, filename)
    
    def process_excel_files(self):
        """Main processing function"""
        try:
            # Get file paths
            aptive_path = self.get_file_path(self.config["files"]["aptive_file"])
            template_path = self.get_file_path(self.config["files"]["bom_template"])
            
            # Generate output path
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{self.config['files']['output_prefix']}_{timestamp}.xlsx"
            output_path = self.get_file_path(output_filename)
            
            print(f"📂 Working directory: {self.workspace_path}")
            print(f"📖 Aptive file: {self.config['files']['aptive_file']}")
            print(f"📋 Template file: {self.config['files']['bom_template']}")
            print(f"💾 Output file: {output_filename}")
            print()
            
            # Check if files exist
            if not os.path.exists(aptive_path):
                raise FileNotFoundError(f"Aptive file not found: {aptive_path}")
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            
            # Read Aptive data
            print("📖 Reading Aptive data...")
            if aptive_path.endswith('.xlsb'):
                df = pd.read_excel(aptive_path, engine='pyxlsb')
            else:
                df = pd.read_excel(aptive_path)
            
            print(f"✓ Loaded {len(df)} rows")
            
            # Process data by levels
            print("🔄 Processing level data...")
            level_data = self.extract_level_data(df)
            
            # Write to BOM template
            print("✏️ Writing to BOM template...")
            self.write_to_bom(level_data, template_path, output_path)
            
            print(f"✅ Processing completed!")
            print(f"📁 Output: {output_filename}")
            
            return output_path
            
        except Exception as e:
            print(f"❌ Error: {str(e)}")
            raise
    
    def extract_level_data(self, df):
        """Extract and organize data by levels"""
        level_data = {}
        source_cols = self.config["column_mapping"]["aptive_source_columns"]
        
        start_idx = 1 if self.config["settings"]["skip_header"] else 0
        
        for index in range(start_idx, len(df)):
            row = df.iloc[index]
            
            # Get level from first column
            level_value = str(row.iloc[source_cols[0]]).strip()
            
            # Skip empty rows
            if pd.isna(row.iloc[source_cols[0]]) or level_value == "" or level_value == "nan":
                continue
            
            # Extract numeric level
            numbers = re.findall(r'\d+', level_value)
            if not numbers:
                continue
                
            level = int(numbers[0])
            
            # Extract data from specified columns
            data_item = {
                'level_text': level_value,
                'data': []
            }
            
            # Get data from remaining source columns
            for col_idx in source_cols[1:]:
                if col_idx < len(row):
                    value = str(row.iloc[col_idx]) if not pd.isna(row.iloc[col_idx]) else ""
                    data_item['data'].append(value)
                else:
                    data_item['data'].append("")
            
            # Group by level
            if level not in level_data:
                level_data[level] = []
            level_data[level].append(data_item)
        
        # Print summary
        for level in sorted(level_data.keys()):
            print(f"  Level {level}: {len(level_data[level])} items")
        
        return level_data
    
    def write_to_bom(self, level_data, template_path, output_path):
        """Write processed data to BOM template"""
        workbook = load_workbook(template_path)
        worksheet = workbook.active
        
        target_cols = self.config["column_mapping"]["bom_target_columns"]
        current_row = self.config["settings"]["start_row"]
        
        # Write data level by level
        for level in sorted(level_data.keys()):
            print(f"  Writing Level {level}...")
            
            for item in level_data[level]:
                # Write level text to first target column
                worksheet[f'{target_cols[0]}{current_row}'] = item['level_text']
                
                # Write data to remaining columns
                for i, value in enumerate(item['data']):
                    if i + 1 < len(target_cols):
                        col = target_cols[i + 1]
                        worksheet[f'{col}{current_row}'] = value
                
                current_row += 1
            
            # Add blank row between levels
            if self.config["settings"]["add_blank_between_levels"]:
                current_row += 1
        
        # Save workbook
        workbook.save(output_path)

def create_sample_config():
    """Create a sample configuration file"""
    config_path = "excel_config.json"
    with open(config_path, 'w') as f:
        json.dump(DEFAULT_CONFIG, f, indent=4)
    
    print(f"📝 Created sample configuration: {config_path}")
    print("Edit this file to customize file paths and column mappings")
    
    return config_path

def main():
    """Main function"""
    workspace = r"d:\SAP"
    
    print("🚀 Configurable Excel Level Processor")
    print("=" * 40)
    
    # Check if config file exists, create if not
    config_file = os.path.join(workspace, "excel_config.json")
    if not os.path.exists(config_file):
        print("📝 Creating default configuration file...")
        os.chdir(workspace)
        create_sample_config()
        print()
    
    # Create processor and run
    processor = ConfigurableExcelProcessor(config_file, workspace)
    
    try:
        output_file = processor.process_excel_files()
        print(f"\n🎉 Success! Output file created: {os.path.basename(output_file)}")
        
    except Exception as e:
        print(f"\n💥 Processing failed: {str(e)}")

if __name__ == "__main__":
    main()