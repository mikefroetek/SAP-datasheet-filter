"""
Excel Level-wise Data Processor
This script processes Aptive BOOM Excel files and categorizes data by levels
into a BOM Template Excel file.

Author: Python Assistant
Date: October 28, 2025
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os
from typing import Dict, List, Tuple
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelLevelProcessor:
    def __init__(self, aptive_file_path: str, bom_template_path: str, output_path: str = None):
        """
        Initialize the Excel Level Processor
        
        Args:
            aptive_file_path (str): Path to the Aptive BOOM Excel file
            bom_template_path (str): Path to the BOM Template Excel file
            output_path (str): Path for the output file (optional)
        """
        self.aptive_file_path = aptive_file_path
        self.bom_template_path = bom_template_path
        self.output_path = output_path or self._generate_output_path()
        
        # Column mappings
        self.aptive_columns = {
            'level': 'A',      # Szint (Level)
            'col_b': 'B',      # Column B data
            'col_d': 'D',      # Column D data
            'col_e': 'E'       # Column E data
        }
        
        self.bom_template_columns = {
            'col_e': 'E',      # Maps to Aptive Column A (Level)
            'col_n': 'N',      # Maps to Aptive Column B
            'col_o': 'O',      # Maps to Aptive Column D
            'col_p': 'P'       # Maps to Aptive Column E
        }
        
    def _generate_output_path(self) -> str:
        """Generate output file path based on input files"""
        base_path = os.path.dirname(self.bom_template_path)
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(base_path, f"BOM_Processed_{timestamp}.xlsx")
    
    def read_aptive_data(self) -> pd.DataFrame:
        """
        Read data from Aptive Excel file
        
        Returns:
            pd.DataFrame: Processed Aptive data
        """
        try:
            logger.info(f"Reading Aptive file: {self.aptive_file_path}")
            
            # Handle different Excel formats
            if self.aptive_file_path.endswith('.xlsb'):
                # For .xlsb files, we need to use openpyxl or xlrd
                df = pd.read_excel(self.aptive_file_path, engine='pyxlsb')
            else:
                df = pd.read_excel(self.aptive_file_path)
            
            logger.info(f"Successfully read {len(df)} rows from Aptive file")
            logger.info(f"Columns found: {list(df.columns)}")
            
            return df
            
        except Exception as e:
            logger.error(f"Error reading Aptive file: {str(e)}")
            raise
    
    def process_level_data(self, df: pd.DataFrame) -> Dict[int, List[Dict]]:
        """
        Process data by levels from the Aptive DataFrame
        
        Args:
            df (pd.DataFrame): Aptive data
            
        Returns:
            Dict[int, List[Dict]]: Data organized by levels
        """
        try:
            logger.info("Processing level-wise data categorization")
            
            # Get the relevant columns (A, B, D, E)
            # Assuming first column is A (level/Szint), then B, C, D, E...
            columns = df.columns.tolist()
            
            if len(columns) < 5:
                raise ValueError("Not enough columns in the Aptive file")
            
            # Extract data from columns A, B, D, E (indices 0, 1, 3, 4)
            level_col = df.iloc[:, 0]      # Column A (Szint/Level)
            col_b_data = df.iloc[:, 1]     # Column B
            col_d_data = df.iloc[:, 3]     # Column D
            col_e_data = df.iloc[:, 4]     # Column E
            
            # Organize data by levels
            level_data = {}
            
            for idx in range(len(df)):
                # Get level value and clean it
                level_val = level_col.iloc[idx]
                
                # Skip if level is NaN or empty
                if pd.isna(level_val) or str(level_val).strip() == '':
                    continue
                
                # Extract numeric level (handle cases like "Level 1", "1", etc.)
                try:
                    if isinstance(level_val, str):
                        # Extract number from string
                        import re
                        numbers = re.findall(r'\d+', str(level_val))
                        if numbers:
                            level = int(numbers[0])
                        else:
                            continue
                    else:
                        level = int(level_val)
                except (ValueError, TypeError):
                    logger.warning(f"Could not parse level value: {level_val}")
                    continue
                
                # Create data entry
                data_entry = {
                    'original_level': level_val,
                    'level': level,
                    'col_b': col_b_data.iloc[idx] if not pd.isna(col_b_data.iloc[idx]) else '',
                    'col_d': col_d_data.iloc[idx] if not pd.isna(col_d_data.iloc[idx]) else '',
                    'col_e': col_e_data.iloc[idx] if not pd.isna(col_e_data.iloc[idx]) else ''
                }
                
                # Add to level group
                if level not in level_data:
                    level_data[level] = []
                level_data[level].append(data_entry)
            
            # Log level statistics
            for level, data in level_data.items():
                logger.info(f"Level {level}: {len(data)} items")
            
            return level_data
            
        except Exception as e:
            logger.error(f"Error processing level data: {str(e)}")
            raise
    
    def write_to_bom_template(self, level_data: Dict[int, List[Dict]]) -> None:
        """
        Write processed data to BOM Template file
        
        Args:
            level_data (Dict[int, List[Dict]]): Processed level data
        """
        try:
            logger.info(f"Writing data to BOM template: {self.output_path}")
            
            # Load the BOM template
            workbook = load_workbook(self.bom_template_path)
            
            # Assuming we work with the first worksheet
            if len(workbook.worksheets) > 0:
                worksheet = workbook.worksheets[0]
            else:
                worksheet = workbook.active
            
            current_row = 2  # Start from row 2 (assuming row 1 has headers)
            
            # Process each level in order
            for level in sorted(level_data.keys()):
                logger.info(f"Writing Level {level} data starting at row {current_row}")
                
                level_items = level_data[level]
                
                for item in level_items:
                    # Write data to columns E, N, O, P
                    worksheet[f'E{current_row}'] = item['original_level']  # Aptive Column A -> BOM Column E
                    worksheet[f'N{current_row}'] = item['col_b']           # Aptive Column B -> BOM Column N
                    worksheet[f'O{current_row}'] = item['col_d']           # Aptive Column D -> BOM Column O
                    worksheet[f'P{current_row}'] = item['col_e']           # Aptive Column E -> BOM Column P
                    
                    current_row += 1
                
                # Add a blank row between levels for better readability
                current_row += 1
            
            # Save the workbook
            workbook.save(self.output_path)
            logger.info(f"Successfully saved processed data to: {self.output_path}")
            
        except Exception as e:
            logger.error(f"Error writing to BOM template: {str(e)}")
            raise
    
    def process_files(self) -> str:
        """
        Main method to process the files
        
        Returns:
            str: Path to the output file
        """
        try:
            logger.info("Starting Excel level processing")
            
            # Step 1: Read Aptive data
            aptive_df = self.read_aptive_data()
            
            # Step 2: Process level-wise data
            level_data = self.process_level_data(aptive_df)
            
            # Step 3: Write to BOM template
            self.write_to_bom_template(level_data)
            
            logger.info("Excel level processing completed successfully")
            return self.output_path
            
        except Exception as e:
            logger.error(f"Error in processing files: {str(e)}")
            raise

def main():
    """Main function to run the Excel processor"""
    
    # File paths (modify these as needed)
    aptive_file = r"d:\SAP\Aptive_BOOM_v01.xlsb"
    bom_template = r"d:\SAP\BOM_Template.xlsx"
    
    # Check if files exist
    if not os.path.exists(aptive_file):
        print(f"Error: Aptive file not found: {aptive_file}")
        return
    
    if not os.path.exists(bom_template):
        print(f"Error: BOM Template file not found: {bom_template}")
        return
    
    try:
        # Create processor instance
        processor = ExcelLevelProcessor(aptive_file, bom_template)
        
        # Process files
        output_file = processor.process_files()
        
        print(f"\n✓ Processing completed successfully!")
        print(f"✓ Output file created: {output_file}")
        
    except Exception as e:
        print(f"\n✗ Error occurred during processing: {str(e)}")
        logger.exception("Detailed error information:")

if __name__ == "__main__":
    main()