ChatGPT Canvas to PowerPoint Converter
🚀 Features

Smart Content Extraction: Multiple strategies to reliably extract content from ChatGPT Canvas
Advanced PowerPoint Generation: Professional slides with proper formatting, tables, code blocks, and lists
Speaker Notes Support: Automatically extracts and adds speaker notes to slides
Table Processing: Handles complex tables with proper formatting and styling
Code Block Support: Preserves code formatting with monospace fonts
Batch Processing: Convert multiple canvas URLs at once
Interactive CLI: User-friendly command-line interface
Robust Error Handling: Comprehensive error recovery and logging
Safe File Handling: Automatic filename sanitization and unique naming

📋 Requirements
System Requirements

Python 3.7+
Chrome browser (for web scraping)
Internet connection

Command Line Mode
Single URL Conversion
bash# Basic conversion
python script.py -u "https://chatgpt.com/share/your-canvas-url"

# With custom output directory

python script.py -u "https://chatgpt.com/share/your-canvas-url" -o "./presentations"

# With custom filename

python script.py -u "https://chatgpt.com/share/your-canvas-url" -f "my-presentation.pptx"

# Verbose logging

python script.py -u "https://chatgpt.com/share/your-canvas-url" -v
Batch Processing
bash# Create a text file with URLs (one per line)
echo "https://chatgpt.com/share/url1" > urls.txt
echo "https://chatgpt.com/share/url2" >> urls.txt

# Process all URLs

python script.py -b urls.txt -o "./batch\_output"
Command Line Options
-u, --url       ChatGPT canvas URL to convert
-o, --output    Output directory
-f, --filename  Custom output filename
-b, --batch     File containing URLs for batch processing
-v, --verbose   Enable verbose logging
--log-level     Set logging level (DEBUG, INFO, WARNING, ERROR)

Font Customization
pythonfont\_fallbacks: Dict\[str, str] = {
'default': 'Calibri',
'code': 'Courier New',
'math': 'Cambria Math',
'fallback': 'Arial'
}



# Key Features:



**Multiple Mapping Formats:** Supports JSON, CSV, and TXT mapping files



**Batch Processing:** Processes all mappings automatically without user interaction



**Flexible Positioning**: Supports preset positions (center, corners) and custom coordinates



**Error Handling:** Validates mappings and provides detailed feedback



**Sample Generator:** Run with --create-samples to generate example mapping files



# **Mapping File Formats:**



# JSON Format:



###### json\[

###### &nbsp; {"image\_number": 1, "slide\_number": 1, "position": "center"},

###### &nbsp; {"image\_number": 2, "slide\_number": 2, "position": "custom", "left": 2.0, "top": 1.5, "width": 5.0, "height": 4.0}

###### ]



# CSV Format:



###### csvimage\_number,slide\_number,position,left,top,width,height

###### 1,1,center,,,4.0,3.0

###### 2,2,custom,2.0,1.5,5.0,4.0



#### Usage:



## Generate samples: **python image\_auto.py --create-samples**

## Run normally: **python image\_auto.py**

## Provide: PPT file, images folder, and mapping document



