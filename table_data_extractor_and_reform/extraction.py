import base64
import json
import os
import time
import random
import openpyxl
from typing import Dict, Any
from openpyxl.styles import Alignment, PatternFill
from copy import copy
import google.generativeai as genai
from dotenv import load_dotenv
import cv2

# Load environment variables
load_dotenv()

###########################################
# PDF Processing Functions
###########################################

def rotate_image(image_path):
    """Rotate the image 90 degrees counterclockwise."""
    try:
        # Read the image
        image = cv2.imread(image_path)
        if image is None:
            print(f"Error: Unable to read image {image_path}.")
            return None

        # Rotate the image 90 degrees counterclockwise
        rotated_image = cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)

        # Save the rotated image to a temporary file
        rotated_image_path = "rotated_" + os.path.basename(image_path)
        cv2.imwrite(rotated_image_path, rotated_image)

        return rotated_image_path
    except Exception as e:
        print(f"Error rotating image {image_path}: {str(e)}")
        return None

def process_image(image_path, max_retries=5, initial_delay=1):
    """Process image directly with Gemini API and return extracted data."""
    
    for attempt in range(max_retries):      
        try:
            # Read and encode image
            try:
                with open(image_path, 'rb') as file:
                    image_bytes = file.read()
                    image_base64 = base64.b64encode(image_bytes).decode('utf-8')
            except Exception as e:
                print(f"Error reading image {image_path}: {str(e)}")
                return None

            # Configure Gemini
            genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
            model = genai.GenerativeModel('gemini-1.5-flash')

            # Create the prompt
            prompt = """
You are an expert table data extractor specializing in structural engineering measurements. Your task is to extract information from any table image with these characteristics:
1. Headers are in the bottom row
2. Data flows from bottom to top in each column
3. Cells may be merged vertically
4. CRITICAL: All output lists MUST have the same length as the Section list

CRITICAL EXTRACTION RULES:

1. LIST LENGTH REQUIREMENT:
   - Count the total number of sections (N) from the Section column
   - EVERY list in the output JSON MUST have exactly N values
   - This is the most important rule and applies to all columns without exception

2. HEADER AND MEASUREMENT IDENTIFICATION:
   - Examine ONLY the bottom-most row for headers
   - For headers containing "ft" or "feet":
     * Preserve the entire measurement unit in the header
     * Keep any numerical values that are part of the header (e.g., "Face Width (ft) 17")
   - For headers with measurements:
     * Maintain the exact format of measurement specifications
     * Keep parenthetical units intact (e.g., "(ft)", "(in)")

3. DIMENSIONAL VALUE HANDLING:
   - For columns with "ft" or feet measurements:
     * Preserve exact numerical values
     * Maintain decimal precision as shown
     * Keep fractional representations if present
     * For merged cells, ensure measurement consistency across repeated values
   - For panel spacing (e.g., "8 @ 10"):
     * Keep the exact format "quantity @ spacing"
     * Preserve spacing measurements precisely
     * Repeat these patterns based on section spans

4. VALUE REPETITION AND MEASUREMENT CONSISTENCY:
   For each column, especially those with measurements:
   - Single section span: Use the value once
   - Multiple section span: 
     * Repeat the EXACT measurement for each spanned section
     * Maintain consistent units and precision
     * Keep spacing patterns intact (e.g., "@ 10 ft")
   - Incomplete column fill:
     * Repeat the last valid measurement for remaining sections
     * Maintain measurement format consistency

5. MERGED CELLS WITH MEASUREMENTS:
   - Count sections spanned by each merged cell
   - For dimensional values:
     * Repeat the exact measurement for each section
     * Maintain precision and format
     * Preserve unit consistency
   - For panel specifications:
     * Keep the complete pattern (e.g., "8 @ 10")
     * Repeat for all relevant sections

6. QUALITY CHECKS FOR MEASUREMENTS:
   - Verify each measurement list has exactly N values
   - Confirm dimensional consistency within columns
   - Check for proper unit preservation
   - Validate measurement patterns in repeated values
   - Ensure spacing patterns are maintained

OUTPUT FORMAT:
{
  "Column1_Header": [value1, value2, ..., valueN],
  "Column2_Header": [value1, value2, ..., valueN],
  ...
}

Where:
- Headers preserve measurement units and specifications
- Each list contains exactly N values
- Measurements maintain their exact format and precision
- Panel spacing patterns are preserved
- Values are ordered from bottom to top

CRITICAL RULES:
- Every list MUST have exactly N values (where N is the section count)
- Preserve all measurement units and formats
- Maintain exact numerical precision
- Keep spacing patterns intact
- Repeat measurements accurately for merged cells
- Ensure dimensional consistency in repeated values
- Include all columns from the table

Return ONLY the JSON output. Do not include any other text, explanations, or markdown formatting."""
            
            # Create image part object
            image_part = {
                "mime_type": "image/jpeg",  # Will work for both JPEG and PNG
                "data": image_base64
            }
            
            # Generate content with both prompt and image
            response = model.generate_content(
                [prompt, image_part],
                generation_config={
                    'temperature': 0.1,
                    'top_p': 0.1,
                    'top_k': 16,
                }
            )
            
            # Process response
            response_text = response.text.strip()
            
            try:
                # Clean response text
                response_text = response_text.strip('```json').strip('```').strip()
                
                # Parse JSON and return directly
                return json.loads(response_text)
                
            except json.JSONDecodeError as e:
                print(f"JSON parsing error: {str(e)}")
                if attempt < max_retries - 1:
                    continue
                return None
                
        except Exception as e:
            if "429" in str(e):
                wait_time = (initial_delay * (2 ** attempt) + random.random())
                print(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds (attempt {attempt + 1}/{max_retries})")
                time.sleep(wait_time)
            else:
                print(f"Error processing image: {str(e)}")
                if attempt < max_retries - 1:
                    continue
                return None
    
    print(f"Failed to process image after {max_retries} retries.")
    return None

def save_json_output(data: Dict[str, Any], output_path: str):
    """Save the extracted data as a JSON file."""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"Successfully saved JSON output to: {output_path}")
    except Exception as e:
        print(f"Error saving JSON output: {str(e)}")

def process_all_tables():
    """Process all table images and generate JSON output."""
    
    # Define input and output directories
    input_dir = "input_tables"  # Directory containing your table images
    output_dir = "output_json"  # Directory for JSON output
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Get all image files from input directory
    image_extensions = ('.jpg', '.jpeg', '.png')
    image_files = [f for f in os.listdir(input_dir) 
                  if f.lower().endswith(image_extensions)]
    
    if not image_files:
        print(f"No image files found in {input_dir}")
        print(f"Supported formats: {', '.join(image_extensions)}")
        return
    
    print(f"Found {len(image_files)} image files to process")
    
    # Process each image file
    for index, image_file in enumerate(image_files, 1):
        image_path = os.path.join(input_dir, image_file)
        
        try:
            print(f"\nProcessing file {index} of {len(image_files)}: {image_file}")
            
            # Process image and get JSON data
            extracted_data = process_image(image_path)
            
            if not extracted_data:
                print(f"Failed to extract data from {image_file}. Skipping to next file.")
                continue
            
            # Create output JSON file name
            json_filename = os.path.splitext(image_file)[0] + '_data.json'
            output_path = os.path.join(output_dir, json_filename)
            
            # Save JSON output
            save_json_output(extracted_data, output_path)
            
            # Also print the JSON to console
            print("\nExtracted Data:")
            print(json.dumps(extracted_data, indent=2))
            
        except Exception as e:
            print(f"Error processing {image_file}: {str(e)}")
            continue
    
    print("\nAll files processed!")

def main():
    """Main entry point of the script."""
    try:
        # Check if GEMINI_API_KEY is set
        if not os.getenv("GEMINI_API_KEY"):
            print("Error: GEMINI_API_KEY not found in environment variables.")
            print("Please create a .env file with your Gemini API key.")
            return
        
        # Create input directory if it doesn't exist
        if not os.path.exists("input_tables"):
            os.makedirs("input_tables")
            print("Created 'input_tables' directory.")
            print("Please place your table images (JPG, JPEG, or PNG) in the 'input_tables' directory and run the script again.")
            return
        
        # Process all tables
        process_all_tables()
        
    except Exception as e:
        print(f"Error in main process: {str(e)}")

if __name__ == "__main__":
    main()
