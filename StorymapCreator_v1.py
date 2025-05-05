import arcpy
import os
import json
import tempfile
import zipfile
import xml.etree.ElementTree as ET
import shutil
from datetime import datetime
import re
import mimetypes
from typing import Dict, List, Tuple, Any, Optional, Union
import pandas as pd

from arcgis.gis import GIS
# Resolve Image class naming conflict by using aliases
from PIL import Image as PILImage
from arcgis.apps.storymap import StoryMap, Image as StoryImage, Video, Audio, Text, TextStyles, Table, Code, Separator, Language


# Constants for namespaces in DOCX files
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'v': 'urn:schemas-microsoft-com:vml',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'ind': 'http://schemas.openxmlformats.org/drawingml/2006/indicator'
}

# Global debug settings
DEBUG_LEVEL = "none"  # Default debug level
DEBUG_OUTPUT_FOLDER = None
LOG_FILE_PATH = None
JSON_FILE_PATH = None

def main():
    """Main function that orchestrates the entire process."""
    try:
        # Get parameters from ArcGIS Pro
        content_file = arcpy.GetParameterAsText(0)  # file to be parsed
        storymap_title = arcpy.GetParameterAsText(1)  # title of portal item and main heading of storymap
        storymap_tags = arcpy.GetParameterAsText(2)  # Optional - portal item tags
        storymap_summary = arcpy.GetParameterAsText(3)  # Optional - portal item summary and storymap main subtitle
        storymap_description = arcpy.GetParameterAsText(4)  # Optional - portal item description and added as first text block of storymap
        storymap_cover_image = arcpy.GetParameterAsText(5)  # Story map cover image
        config_path = arcpy.GetParameterAsText(6)  # path to config.json - takes precedence over following
        username = arcpy.GetParameterAsText(7)  # if you dont have config json username 
        password = arcpy.GetParameterAsText(8)  # if you dont have config json password
        arcgis_url = arcpy.GetParameterAsText(9)  # if you dont have config json URL of server
        
        # Parse tags into a list
        if storymap_tags:
            tags = [tag.strip() for tag in storymap_tags.split(',')]
        else:
            tags = []
            
        # Get credentials and configure debug settings
        credentials = get_credentials(config_path, username, password, arcgis_url)
        
        # Initialize debug settings - pass the storymap title for potential log filenames
        initialize_debug_settings(credentials, storymap_title)
        
        # Connect to portal
        log_message("Connecting to ArcGIS Portal...", "none")
        gis = connect_to_portal(credentials)
        
        # Parse input file
        log_message(f"Parsing content file: {content_file}", "none")
        content_blocks = parse_content_file(content_file)
        
        # Log summary of parsed blocks
        if isinstance(content_blocks, tuple) and len(content_blocks) == 2:
            blocks, temp_dir = content_blocks
            log_message(f"Parsed {len(blocks)} content blocks", "none")
        else:
            blocks = content_blocks
            log_message(f"Parsed {len(blocks)} content blocks", "none")
        
        # Count block types for summary
        block_counts = {}
        for block in blocks:
            block_type = block.get('type', 'unknown')
            if block_type in block_counts:
                block_counts[block_type] += 1
            else:
                block_counts[block_type] = 1
        
        log_message(f"Content block summary: {', '.join([f'{count} {type}(s)' for type, count in block_counts.items()])}", "none")
        
        # Create StoryMap with placeholders
        log_message("Creating StoryMap...", "none")
        storymap_item, placeholder_ids, parsed_blocks, image_dimensions = create_storymap(
            gis, 
            storymap_title, 
            tags, 
            storymap_summary, 
            storymap_description, 
            storymap_cover_image, 
            content_blocks
        )
        
        # Update StoryMap with actual content
        log_message("Updating StoryMap with actual content...", "none")
        update_storymap_json(storymap_item, placeholder_ids, parsed_blocks, image_dimensions)
        
        # Generate URLs
        if 'arcgis.com' in credentials['arcgis_url'].lower():
            edit_url = f"https://storymaps.arcgis.com/stories/{storymap_item.id}/edit"
            view_url = f"https://storymaps.arcgis.com/stories/{storymap_item.id}"
        else:
            portal_url = credentials['arcgis_url'].rstrip('/')
            edit_url = f"{portal_url}/apps/storymaps/stories/{storymap_item.id}/edit"
            view_url = f"{portal_url}/apps/storymaps/stories/{storymap_item.id}"
            
        log_message(f"StoryMap created successfully!", "none")
        log_message(f"Edit URL: {edit_url}", "none")
        log_message(f"View URL: {view_url}", "none")
        
        # If debug is full, save the StoryMap JSON to a file
        if DEBUG_LEVEL == "full" and JSON_FILE_PATH:
            storymap_data = storymap_item.get_data()
            save_storymap_json(storymap_data, JSON_FILE_PATH)
            log_message(f"StoryMap JSON saved to: {JSON_FILE_PATH}", "none")
            log_message(f"Debug log saved to: {LOG_FILE_PATH}", "none")
            
        return storymap_item
        
    except Exception as e:
        log_message(f"An error occurred: {str(e)}", "none", is_error=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_error=True)
        raise e


def initialize_debug_settings(credentials, storymap_title):
    """Initialize debug settings from credentials dictionary."""
    global DEBUG_LEVEL, DEBUG_OUTPUT_FOLDER, LOG_FILE_PATH, JSON_FILE_PATH
    
    # Set debug level from credentials
    DEBUG_LEVEL = credentials.get('debug', "none")
    if DEBUG_LEVEL not in ["none", "basic", "full"]:
        log_message(f"Invalid debug level '{DEBUG_LEVEL}', defaulting to 'none'", "none", is_warning=True)
        DEBUG_LEVEL = "none"
    
    # If debug level is full, set up log and JSON file paths
    if DEBUG_LEVEL == "full":
        DEBUG_OUTPUT_FOLDER = credentials.get('full_debug_output_folder')
        
        if not DEBUG_OUTPUT_FOLDER:
            log_message("Debug level is 'full' but no output folder specified. Using temp directory.", "none", is_warning=True)
            DEBUG_OUTPUT_FOLDER = tempfile.gettempdir()
        
        # Create the output folder if it doesn't exist
        if not os.path.exists(DEBUG_OUTPUT_FOLDER):
            try:
                os.makedirs(DEBUG_OUTPUT_FOLDER)
                log_message(f"Created debug output folder: {DEBUG_OUTPUT_FOLDER}", "basic")
            except Exception as e:
                log_message(f"Failed to create debug output folder: {str(e)}", "none", is_warning=True)
                DEBUG_OUTPUT_FOLDER = tempfile.gettempdir()
                log_message(f"Using temp directory instead: {DEBUG_OUTPUT_FOLDER}", "none", is_warning=True)
        
        # Create log and JSON file paths
        base_filename = create_safe_filename(storymap_title)
        LOG_FILE_PATH, JSON_FILE_PATH = generate_debug_file_paths(base_filename, DEBUG_OUTPUT_FOLDER)
        
        # Initialize log file with header
        with open(LOG_FILE_PATH, 'w', encoding='utf-8') as log_file:
            log_file.write(f"=== StoryMap Creator Debug Log ===\n")
            log_file.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log_file.write(f"StoryMap Title: {storymap_title}\n")
            log_file.write(f"Debug Level: {DEBUG_LEVEL}\n")
            log_file.write(f"='='='='='='='='='='='='='='='='='='='='=\n\n")
    
    log_message(f"Debug level set to: {DEBUG_LEVEL}", "basic")


def create_safe_filename(title):
    """Create a safe filename from the StoryMap title."""
    if not title:
        return "storymap"
    
    # Replace spaces with underscores and remove unsafe characters
    safe_title = re.sub(r'[\\/*?:"<>|]', '', title.strip())
    safe_title = safe_title.replace(' ', '_')
    safe_title = re.sub(r'__+', '_', safe_title)  # Replace multiple underscores with single
    
    # Add date suffix
    date_suffix = datetime.now().strftime('%Y-%m-%d')
    return f"{safe_title}_{date_suffix}"


def generate_debug_file_paths(base_filename, output_folder):
    """Generate unique file paths for log and JSON files."""
    log_path = os.path.join(output_folder, f"{base_filename}.txt")
    json_path = os.path.join(output_folder, f"{base_filename}.json")
    
    # Check if files already exist and add incremental numbers if needed
    counter = 2
    while os.path.exists(log_path) or os.path.exists(json_path):
        log_path = os.path.join(output_folder, f"{base_filename}_{counter}.txt")
        json_path = os.path.join(output_folder, f"{base_filename}_{counter}.json")
        counter += 1
    
    return log_path, json_path


def log_message(message, min_level="none", is_error=False, is_warning=False):
    """
    Log a message based on the debug level.
    
    Parameters:
        message (str): The message to log
        min_level (str): Minimum debug level required to show this message ("none", "basic", or "full")
        is_error (bool): Whether this is an error message
        is_warning (bool): Whether this is a warning message
    """
    # Determine if we should display the message based on debug level
    level_priority = {"none": 0, "basic": 1, "full": 2}
    current_level = level_priority.get(DEBUG_LEVEL, 0)
    required_level = level_priority.get(min_level, 0)
    
    # Always log errors and warnings regardless of level
    should_display = current_level >= required_level or is_error or is_warning
    
    if should_display:
        # Log to ArcGIS Pro
        if is_error:
            arcpy.AddError(message)
        elif is_warning:
            arcpy.AddWarning(message)
        else:
            arcpy.AddMessage(message)
    
    # Write to log file if debug level is full
    if DEBUG_LEVEL == "full" and LOG_FILE_PATH:
        try:
            with open(LOG_FILE_PATH, 'a', encoding='utf-8') as log_file:
                timestamp = datetime.now().strftime('%H:%M:%S')
                prefix = ""
                if is_error:
                    prefix = "[ERROR] "
                elif is_warning:
                    prefix = "[WARNING] "
                
                log_file.write(f"[{timestamp}] {prefix}{message}\n")
        except Exception as e:
            arcpy.AddWarning(f"Failed to write to log file: {str(e)}")


def save_storymap_json(data, file_path):
    """Save StoryMap JSON data to a file."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        log_message(f"Failed to save StoryMap JSON: {str(e)}", "none", is_warning=True)
        return False


def get_credentials(config_path=None, username=None, password=None, arcgis_url=None):
    """Retrieve credentials from config file or ArcGIS Pro parameters."""
    credentials = {
        'username': None,
        'password': None,
        'arcgis_url': None,
        'debug': "none",
        'full_debug_output_folder': None
    }
    
    # Set direct parameters first
    if username:
        credentials['username'] = username
    if password:
        credentials['password'] = password
    if arcgis_url:
        credentials['arcgis_url'] = arcgis_url
    
    # Override with config file if provided (config file takes precedence)
    if config_path and os.path.exists(config_path):
        config = load_config(config_path)
        if config:
            credentials.update(config)
            log_message("Loaded credentials from config file", "basic")
        
    # Validate credentials
    if not all([credentials['username'], credentials['password'], credentials['arcgis_url']]):
        missing = [k for k, v in credentials.items() if not v and k in ['username', 'password', 'arcgis_url']]
        error_msg = f"Missing credentials: {', '.join(missing)}"
        log_message(error_msg, "none", is_error=True)
        raise ValueError(error_msg)
        
    return credentials


def load_config(config_path):
    """Load credentials from a config.json file."""
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        # Validate required fields
        required_fields = ['username', 'password', 'arcgis_url']
        if not all(field in config for field in required_fields):
            missing = [field for field in required_fields if field not in config]
            log_message(f"Config file is missing fields: {', '.join(missing)}", "none", is_warning=True)
            return None
        
        # Log debug settings from config
        if 'debug' in config:
            log_message(f"Debug level found in config: {config['debug']}", "basic")
        
        if 'full_debug_output_folder' in config:
            log_message(f"Debug output folder found in config: {config['full_debug_output_folder']}", "basic")
            
        return config
    except Exception as e:
        log_message(f"Error loading config file: {str(e)}", "none", is_warning=True)
        return None


def connect_to_portal(credentials):
    """Connect to ArcGIS portal using provided credentials."""
    try:
        gis = GIS(
            credentials['arcgis_url'],
            credentials['username'],
            credentials['password']
        )
        log_message(f"Connected to {credentials['arcgis_url']} as {credentials['username']}", "none")
        return gis
    except Exception as e:
        log_message(f"Failed to connect to ArcGIS portal: {str(e)}", "none", is_error=True)
        raise


def parse_content_file(file_path):
    """Parse content file based on its extension."""
    if not os.path.exists(file_path):
        log_message(f"File not found: {file_path}", "none", is_error=True)
        raise FileNotFoundError(f"File not found: {file_path}")
    
    extension = os.path.splitext(file_path)[1].lower()
    
    if extension == '.docx':
        log_message("Detected DOCX file format", "basic")
        return parse_docx(file_path)
    elif extension in ['.html', '.htm']:
        log_message("Detected HTML file format", "basic")
        return parse_html(file_path)
    else:
        log_message(f"Unsupported file type: {extension}", "none", is_error=True)
        raise ValueError(f"Unsupported file type: {extension}")


def parse_html(file_path):
    """Parse HTML file and extract content blocks."""
    try:
        from bs4 import BeautifulSoup
        
        with open(file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
        log_message(f"HTML file size: {len(html_content)} bytes", "basic")
        soup = BeautifulSoup(html_content, 'html.parser')
        
        blocks = []
        
        # Process paragraphs, headers, images, etc.
        log_message("Parsing HTML elements...", "basic")
        for element in soup.body.find_all(recursive=False):
            block = process_html_element(element)
            if block:
                blocks.append(block)
                log_message(f"Processed HTML element: {element.name} -> {block['type']}", "full")
        
        log_message(f"HTML parsing complete. Found {len(blocks)} content blocks", "basic")
        return blocks
    except ImportError:
        log_message("BeautifulSoup not installed. HTML parsing requires BeautifulSoup library.", "none", is_warning=True)
        return []
    except Exception as e:
        log_message(f"Error parsing HTML file: {str(e)}", "none", is_error=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_error=True)
        return []


def process_html_element(element):
    """Process a single HTML element and convert to StoryMap block."""
    tag_name = element.name.lower() if hasattr(element, 'name') else None
    
    if not tag_name:
        return None
    
    log_message(f"Processing HTML element: {tag_name}", "full")
    
    # Handle different element types
    if tag_name in ['h1', 'h2']:
        return create_text_block('h2', element.get_text())
    elif tag_name == 'h3':
        return create_text_block('h3', element.get_text())
    elif tag_name == 'h4':
        return create_text_block('h4', element.get_text())
    elif tag_name == 'p':
        return create_text_block('paragraph', str(element))
    elif tag_name == 'img':
        return create_image_block(
            element.get('src'),
            caption=element.get('alt'),
            display='standard'
        )
    elif tag_name == 'hr':
        return create_separator_block()
    elif tag_name == 'pre':
        code_content = element.get_text()
        return create_code_block(code_content)
    elif tag_name in ['ul', 'ol']:
        list_items = []
        for li in element.find_all('li', recursive=False):
            list_items.append(li.get_text())
        
        list_type = 'bullet-list' if tag_name == 'ul' else 'numbered-list'
        return create_text_block(list_type, '\n'.join(list_items))
    elif tag_name == 'table':
        rows = []
        for tr in element.find_all('tr', recursive=False):
            cells = []
            for td in tr.find_all(['td', 'th'], recursive=False):
                cells.append(td.get_text().strip())
            if cells:
                rows.append(cells)
        if rows:
            return create_table_block(rows)
    
    log_message(f"HTML element type {tag_name} not supported or empty", "full")
    return None

def parse_docx(file_path):
    """Parse DOCX file and extract content blocks."""
    log_message("Parsing DOCX file...", "none")
    temp_dir = None
    
    try:
        # Create a persistent temporary directory
        temp_dir = tempfile.mkdtemp()
        log_message(f"Created temporary directory: {temp_dir}", "full")
        
        # Extract DOCX file
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            log_message("DOCX file extracted successfully", "full")
            
        # Parse document.xml
        document_path = os.path.join(temp_dir, 'word', 'document.xml')
        tree = ET.parse(document_path)
        root = tree.getroot()
        log_message("Parsed document.xml", "full")
        
        # Parse relationships to find images and hyperlinks
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()
        log_message("Parsed document relationships", "full")
        
        # Map relationship IDs to targets (images and hyperlinks)
        image_rels = {}
        hyperlink_rels = {}
        
        for rel in rels_root.findall('.//{*}Relationship'):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type')
            rel_target = rel.get('Target')
            
            # Check if it's an image relationship
            if 'image' in rel_type.lower():
                image_rels[rel_id] = rel_target
                log_message(f"Found image relationship: {rel_id} -> {rel_target}", "full")
            
            # Check if it's a hyperlink relationship
            elif 'hyperlink' in rel_type.lower():
                hyperlink_rels[rel_id] = rel_target
                log_message(f"Found hyperlink relationship: {rel_id} -> {rel_target}", "full")
        
        log_message(f"Found {len(image_rels)} image relationships and {len(hyperlink_rels)} hyperlink relationships", "basic")
        
        # Get media files from the DOCX
        media_files = {}
        for rel_target in image_rels.values():
            if rel_target.startswith('media/'):
                # Normalize path separators
                normalized_target = rel_target.replace('/', os.path.sep)
                source_path = os.path.join(temp_dir, 'word', normalized_target)
                
                # Check if file exists and save it if it does
                if os.path.exists(source_path):
                    media_files[rel_target] = source_path
                    log_message(f"Found media file: {source_path}", "full")
                else:
                    # Try with different path separators
                    alt_path = os.path.join(temp_dir, 'word', rel_target)
                    if os.path.exists(alt_path):
                        media_files[rel_target] = alt_path
                        log_message(f"Found media file (alt path): {alt_path}", "full")
                    else:
                        log_message(f"Media file not found: {source_path} or {alt_path}", "basic", is_warning=True)
        
        log_message(f"Found {len(media_files)} media files", "basic")
        
        # Process document body
        blocks = []
        body = root.find('.//{%s}body' % NAMESPACES['w'])
        if body:
            blocks = process_docx_body(body, NAMESPACES, image_rels, media_files, hyperlink_rels)
            log_message(f"Processed document body, found {len(blocks)} content blocks", "basic")
        else:
            log_message("No document body found", "basic", is_warning=True)
        
        return blocks, temp_dir
        
    except Exception as e:
        log_message(f"Error parsing DOCX file: {str(e)}", "none", is_error=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_error=True)
        raise
    finally:
        # Don't delete the temp directory immediately, it will be cleaned up by the system later
        pass


def process_docx_body(body, namespaces, image_rels, media_files, hyperlink_rels=None):
    """Process the body of a DOCX document with improved caption handling."""
    if hyperlink_rels is None:
        hyperlink_rels = {}
        
    blocks = []
    
    # Add debug to show namespace mapping
    log_message(f"DEBUG: Using namespaces: {namespaces}", "basic")
    
    # Debug count all mc:AlternateContent elements
    alt_content_elements = []
    for elem in body.iter():
        if elem.tag.endswith('AlternateContent'):
            alt_content_elements.append(elem)
    
    log_message(f"DEBUG: Found {len(alt_content_elements)} mc:AlternateContent elements in document", "basic")
    
    # Debug find all images with different approaches
    blip_elements = []
    for elem in body.iter():
        if elem.tag.endswith('blip'):
            blip_elements.append(elem)
    
    log_message(f"DEBUG: Found {len(blip_elements)} blip elements in document", "basic")
    
    # Track list collection and current element type
    current_list_items = []
    current_list_type = None
    in_list_context = False
    
    # Track previous element for caption association
    previous_element_type = None
    previous_block = None
    pending_caption = None
    
    # Additional tracking variables for caption handling
    image_has_caption = False  # Flag to track if an image already has a caption
    
    # Process each paragraph-level element
    log_message("Processing DOCX body elements...", "full")
    element_count = 0
    
    for element in body:
        try:
            element_count += 1
            tag = element.tag.split('}')[-1]
            
            log_message(f"Processing element {element_count}: {tag}", "full")
            
            # Specifically check for <mc:AlternateContent> at the top level
            if tag.endswith('AlternateContent'):
                log_message(f"DEBUG: Processing top-level AlternateContent element", "basic")
                # Try to extract an image from this complex structure
                choice = element.find('.//mc:Choice', namespaces)
                if choice is not None:
                    drawing = choice.find('.//w:drawing', namespaces)
                    if drawing is not None:
                        log_message(f"DEBUG: Found drawing in AlternateContent", "basic")
                        image_block = process_docx_image(drawing, namespaces, image_rels, media_files)
                        if image_block:
                            blocks.append(image_block)
                            previous_element_type = 'image'
                            previous_block = image_block
                            image_has_caption = False  # Reset caption flag for new image
                            continue
            
            # Normal element processing
            if tag == 'p':
                # Process any pending caption first
                if pending_caption:
                    log_message(f"DEBUG: Processing pending caption: {pending_caption[:30]}...", "basic")
                    if previous_block and 'caption' not in previous_block:
                        previous_block['caption'] = pending_caption
                        log_message(f"DEBUG: Added pending caption to previous {previous_element_type}", "basic")
                        # Mark that this image has a caption
                        if previous_element_type == 'image':
                            image_has_caption = True
                    pending_caption = None
                
                # Check if paragraph is a caption for a previous element
                style_elem = element.find('.//{%s}pStyle' % namespaces['w'])
                style_name = style_elem.get('{%s}val' % namespaces['w']) if style_elem is not None else None
                
                # Only check for captions if the previous element was an image or table
                # and (for images) it doesn't already have a caption
                if previous_element_type in ['image', 'table'] and (previous_element_type != 'image' or not image_has_caption):
                    is_caption = is_caption_paragraph(element, namespaces, style_name, previous_element_type)
                    
                    if is_caption:
                        # Extract caption text
                        caption_text, _ = extract_formatted_text(element, namespaces, hyperlink_rels)
                        log_message(f"DEBUG: Detected caption for {previous_element_type}: {caption_text[:30]}...", "basic")
                        
                        # Add caption to the last block
                        if previous_block:
                            previous_block['caption'] = caption_text
                            log_message(f"DEBUG: Added caption to previous {previous_element_type}", "basic")
                            # Mark that this image has a caption
                            if previous_element_type == 'image':
                                image_has_caption = True
                        else:
                            # Store caption for next element
                            pending_caption = caption_text
                            log_message(f"DEBUG: Stored pending caption", "basic")
                        continue  # Skip adding this paragraph as a separate block
                
                # Check if paragraph is a list item
                list_info = get_paragraph_list_info(element, namespaces)
                
                if list_info['is_list_item']:
                    # Process list item
                    log_message(f"Processing list item, level: {list_info['level']}, type: {list_info['list_type']}", "full")
                    
                    # Extract text with formatting
                    text_content, _ = extract_formatted_text(element, namespaces, hyperlink_rels)
                    
                    # Get list information
                    level = list_info['level']
                    list_type = list_info['list_type']
                    
                    # Determine if this is a continuation of the current list or a new list
                    if not in_list_context:
                        # Starting a new list
                        in_list_context = True
                        current_list_type = list_type if level == 0 else None
                    elif level == 0:
                        # New root level - determine if we should process the current list or continue it
                        if current_list_type is None:
                            current_list_type = list_type
                    
                    # Add item to the current list
                    current_list_items.append({
                        'text': text_content,
                        'level': level,
                        'type': list_type,
                        'element_index': element_count
                    })
                    
                    previous_element_type = 'list'
                    previous_block = None
                    
                else:
                    # Non-list paragraph - process any collected list items
                    if in_list_context and current_list_items:
                        # Process the collected list
                        log_message(f"Found non-list paragraph after list. Processing collected {len(current_list_items)} items.", "full")
                        
                        # Process list items
                        all_list_items = {'current_list': current_list_items}
                        num_id_to_type = {'current_list': current_list_type or get_predominant_list_type(current_list_items)}
                        num_id_level_map = {'current_list': min(item['level'] for item in current_list_items)}
                        
                        list_blocks = process_docx_lists(
                            all_list_items, 
                            num_id_to_type,
                            num_id_level_map
                        )
                        
                        # Add list blocks to the main blocks
                        blocks.extend(list_blocks)
                        log_message("Added list block to blocks list", "full")
                        
                        # Reset list tracking
                        current_list_items = []
                        current_list_type = None
                        in_list_context = False
                    
                    # Process regular paragraph
                    paragraph_block = process_docx_paragraph(element, namespaces, image_rels, media_files, hyperlink_rels)
                    if paragraph_block:
                        # Determine element type for caption detection
                        element_type = 'text'
                        if paragraph_block.get('type') == 'image':
                            element_type = 'image'
                            image_has_caption = False  # Reset caption flag for new image
                            
                        blocks.append(paragraph_block)
                        log_message(f"Added paragraph block of type {paragraph_block.get('type')}", "full")
                        previous_element_type = element_type
                        previous_block = paragraph_block
            
            elif tag == 'tbl':
                # Process any collected list items
                if in_list_context and current_list_items:
                    # Process the collected list
                    log_message(f"Found table after list. Processing collected {len(current_list_items)} items.", "full")
                    
                    # Process list items
                    all_list_items = {'current_list': current_list_items}
                    num_id_to_type = {'current_list': current_list_type or get_predominant_list_type(current_list_items)}
                    num_id_level_map = {'current_list': min(item['level'] for item in current_list_items)}
                    
                    list_blocks = process_docx_lists(
                        all_list_items, 
                        num_id_to_type,
                        num_id_level_map
                    )
                    
                    # Add list blocks to the main blocks
                    blocks.extend(list_blocks)
                    log_message("Added list block before table", "full")
                    
                    # Reset list tracking
                    current_list_items = []
                    current_list_type = None
                    in_list_context = False
                
                # Process table
                table_block = process_docx_table(element, namespaces, hyperlink_rels)
                if table_block:
                    blocks.append(table_block)
                    log_message("Added table block", "full")
                    previous_element_type = 'table'
                    previous_block = table_block
            
            # Handle directly embedded drawings (images)
            elif tag == 'drawing':
                log_message(f"DEBUG: Processing direct drawing element", "basic")
                image_block = process_docx_image(element, namespaces, image_rels, media_files)
                if image_block:
                    blocks.append(image_block)
                    log_message("Added direct drawing block", "full")
                    previous_element_type = 'image'
                    previous_block = image_block
                    image_has_caption = False  # Reset caption flag for new image
            
        except Exception as elem_err:
            log_message(f"Error processing element {element_count}: {str(elem_err)}", "basic", is_warning=True)
            import traceback
            log_message(traceback.format_exc(), "full", is_warning=True)
            continue
    
    # Process any remaining list items at the end of the document
    if in_list_context and current_list_items:
        # Process the collected list
        log_message(f"End of document. Processing remaining {len(current_list_items)} list items.", "full")
        
        # Process list items
        all_list_items = {'current_list': current_list_items}
        num_id_to_type = {'current_list': current_list_type or get_predominant_list_type(current_list_items)}
        num_id_level_map = {'current_list': min(item['level'] for item in current_list_items)}
        
        list_blocks = process_docx_lists(
            all_list_items, 
            num_id_to_type,
            num_id_level_map
        )
        
        # Add list blocks to the main blocks
        blocks.extend(list_blocks)
        log_message("Added final list block", "full")
    
    # Final summary
    log_message(f"Processed {element_count} DOCX body elements, created {len(blocks)} content blocks", "basic")
    
    return blocks


def get_predominant_list_type(list_items):
    """
    Determine the predominant list type in a collection of list items.
    Prioritizes items at level 0 (root level).
    
    Parameters:
        list_items: List of list item dictionaries
        
    Returns:
        String: The predominant list type ('bullet-list' or 'numbered-list')
    """
    # First look at level 0 items
    root_items = [item for item in list_items if item['level'] == 0]
    
    if root_items:
        # Count types at root level
        type_counts = {}
        for item in root_items:
            item_type = item.get('type', 'bullet-list')
            type_counts[item_type] = type_counts.get(item_type, 0) + 1
        
        # Get the most common type
        predominant_type = max(type_counts.items(), key=lambda x: x[1])[0]
        log_message(f"Determined predominant list type from {len(root_items)} root items: {predominant_type}", "full")
        return predominant_type
    
    # If no root items, count all items
    type_counts = {}
    for item in list_items:
        item_type = item.get('type', 'bullet-list')
        type_counts[item_type] = type_counts.get(item_type, 0) + 1
    
    # Default to bullet-list if empty
    if not type_counts:
        return 'bullet-list'
    
    # Get the most common type
    predominant_type = max(type_counts.items(), key=lambda x: x[1])[0]
    log_message(f"Determined predominant list type from all {len(list_items)} items: {predominant_type}", "full")
    return predominant_type

def group_list_items(all_list_items, element_indices=None):
    """
    Group list items based on their adjacency in the document.
    Handles mixed list types by normalizing to the type of the root items.
    
    Parameters:
        all_list_items (dict): Dictionary mapping num_id to list of items
        element_indices (dict, optional): Dictionary mapping num_id to original document positions
                                        If None, positions are taken from items themselves
    
    Returns:
        dict: New grouped list structure with assigned group IDs
    """
    # Initialize result
    grouped_lists = {}
    group_counter = 0
    
    # Flatten all list items into a single list with their num_id and sort by element_index
    flat_items = []
    total_items = 0
    
    for num_id, items in all_list_items.items():
        total_items += len(items)
        for item in items:
            flat_item = item.copy()
            flat_item['num_id'] = num_id
            flat_items.append(flat_item)
    
    # Sort by element_index to ensure document order
    flat_items.sort(key=lambda x: x.get('element_index', 0))
    
    # If no items, return empty result
    if not flat_items:
        return grouped_lists
    
    log_message(f"Processing {len(flat_items)} list items from {len(all_list_items)} lists", "full")
    
    # Initialize first group
    current_group = {
        'type': None,
        'document_position': flat_items[0].get('element_index', 0),
        'items': []
    }
    
    # Process items
    current_num_id = None
    last_level = None
    
    for item in flat_items:
        num_id = item.get('num_id')
        level = item.get('level', 0)
        item_type = item.get('type', 'bullet-list')
        item_text = item.get('text', '')[:30] + "..." if len(item.get('text', '')) > 30 else item.get('text', '')
        
        log_message(f"Processing item: level={level}, type={item_type}, text='{item_text}'", "full")
        
        # Check if we should start a new group
        start_new_group = False
        
        # Start a new group if:
        # 1. This is a level 0 item and previous group has items and it's not part of the current logical group
        # 2. There's a significant gap in document position indicating a new list structure
        if level == 0 and current_group['items']:
            # If we're at a root level and the previous item was also at root level,
            # check if this is a continuation of the same list or a new one
            if last_level == 0:
                # For root items, consider document position and num_id
                if num_id != current_num_id:
                    # Different num_id indicates a different list
                    start_new_group = True
                    log_message(f"Starting new group: different num_id at root level", "full")
            else:
                # If previous item was a nested level, don't start a new group just because 
                # we returned to the root level - this is normal list structure
                pass
        
        if start_new_group:
            # Finalize current group
            if current_group['items']:
                group_id = f'group_{group_counter}'
                grouped_lists[group_id] = current_group
                group_counter += 1
                log_message(f"Finalized group with {len(current_group['items'])} items of type {current_group['type']}", "full")
            
            # Start a new group
            current_group = {
                'type': None,
                'document_position': item.get('element_index', 0),
                'items': []
            }
        
        # Update tracking variables
        current_num_id = num_id
        last_level = level
        
        # Add item to current group with a group-specific index
        new_item = {
            'text': item.get('text', ''),
            'level': level,
            'original_type': item_type,  # Store original type for conversion logging
            'group_index': len(current_group['items'])
        }
        current_group['items'].append(new_item)
        
        # Set group type based on first level 0 item if not already set
        if current_group['type'] is None and level == 0:
            current_group['type'] = item_type
            log_message(f"Set group type to {item_type} based on root item", "full")
    
    # Add the last group if it has items
    if current_group['items']:
        group_id = f'group_{group_counter}'
        grouped_lists[group_id] = current_group
        log_message(f"Finalized last group with {len(current_group['items'])} items of type {current_group['type']}", "full")
    
    # Count how many items have a different type than their group
    conversion_count = 0
    for group_id, group in grouped_lists.items():
        group_type = group['type']
        for item in group['items']:
            if item.get('original_type') != group_type:
                conversion_count += 1
                log_message(
                    f"Converting item in group {group_id} from {item.get('original_type')} to {group_type}: "
                    f"level={item['level']}, text='{item['text'][:30]}...'", 
                    "basic"
                )
    
    # Summarize what happened
    total_items_in_groups = sum(len(group['items']) for group in grouped_lists.values())
    log_message(f"Grouped {total_items_in_groups} of {total_items} list items into {len(grouped_lists)} groups", "basic")
    if conversion_count > 0:
        log_message(f"Converted {conversion_count} items to match their group's type", "basic")
    
    # Sanity check - make sure we didn't lose any items
    if total_items_in_groups != total_items:
        log_message(f"WARNING: Item count mismatch! Original={total_items}, Grouped={total_items_in_groups}", "basic", is_warning=True)
    
    return grouped_lists

def flatten_list_hierarchy(grouped_lists):
    """
    Flatten lists with nesting deeper than 2 levels.
    
    Parameters:
        grouped_lists (dict): Grouped list structure
    
    Returns:
        dict: Flattened list structure preserving visual hierarchy
    """
    flattened_lists = {}
    flattened_count = 0
    
    for group_id, group in grouped_lists.items():
        # Find the minimum level in the group (root level)
        min_level = min(item['level'] for item in group['items']) if group['items'] else 0
        
        # Create a new group with the same metadata
        flattened_group = {
            'type': group['type'],
            'document_position': group['document_position'],
            'items': []
        }
        
        # Process each item in the group
        for item in group['items']:
            # Make a copy to avoid modifying the original
            new_item = item.copy()
            
            # Calculate the actual depth relative to the minimum level
            relative_level = item['level'] - min_level
            
            if relative_level > 1:
                # This is a deep level (3+) that needs to be flattened
                excess_depth = relative_level - 1  # How many levels deeper than level 2
                prefix = "--- " * excess_depth
                
                # Add prefix to text content
                original_text = new_item['text']
                new_item['text'] = prefix + original_text
                
                # Set level to second level
                new_item['level'] = min_level + 1
                
                flattened_count += 1
                log_message(
                    f"Flattened list item from level {item['level']} to level {new_item['level']} "
                    f"with prefix: '{prefix}', Text: '{original_text[:30]}...'", 
                    "full"
                )
            
            flattened_group['items'].append(new_item)
        
        flattened_lists[group_id] = flattened_group
    
    if flattened_count > 0:
        log_message(
            f"Flattened {flattened_count} list items from level 3+ to level 2 with '---' prefixes", 
            "basic"
        )
    
    return flattened_lists

def generate_list_html(flattened_lists):
    """
    Generate StoryMap-compatible HTML for each list group.
    
    Parameters:
        flattened_lists (dict): Flattened list structure
    
    Returns:
        dict: Dictionary mapping group IDs to HTML content
    """
    html_content = {}
    
    for group_id, group in flattened_lists.items():
        # Get the minimum level (root level)
        min_level = min(item['level'] for item in group['items']) if group['items'] else 0
        
        # Group items by their parent
        parent_items = []
        current_parent = None
        
        # First pass: identify root items and build structure
        for item in group['items']:
            level = item['level']
            text = item.get('text', '')
            
            if level == min_level:
                # This is a root level item
                current_parent = {'text': text, 'children': []}
                parent_items.append(current_parent)
            elif current_parent is not None:
                # This is a child item of the current parent
                current_parent['children'].append(text)
        
        # Generate the HTML
        html = ""
        
        for parent in parent_items:
            # Add parent list item
            html += f"<li>{parent['text']}"
            
            # Add children if any
            if parent['children']:
                # Use appropriate tag based on list type
                tag = 'ol' if group['type'] == 'numbered-list' else 'ul'
                html += f"<{tag}>"
                
                # Add each child
                for child in parent['children']:
                    html += f"<li>{child}</li>"
                    
                html += f"</{tag}>"
            
            # Close parent list item
            html += "</li>"
        
        # Store the HTML content for this group
        html_content[group_id] = html
        log_message(f"Generated HTML for {group['type']} with {len(group['items'])} items", "full")
    
    return html_content

def process_docx_lists(all_list_items, num_id_to_type, num_id_level_map, element_indices=None):
    """
    Process lists end-to-end from original items to StoryMap blocks.
    
    Parameters:
        all_list_items (dict): Dictionary mapping num_id to list of items
        num_id_to_type (dict): Dictionary mapping num_id to list type
        num_id_level_map (dict): Dictionary mapping num_id to minimum level
        element_indices (dict, optional): Dictionary mapping num_id to original document positions
    
    Returns:
        list: List of blocks ready for StoryMap
    """
    # Step 1: Group list items
    grouped_lists = group_list_items(all_list_items, element_indices)
    
    # Step 2: Flatten list hierarchy
    flattened_lists = flatten_list_hierarchy(grouped_lists)
    
    # Step 3: Generate HTML content
    html_content = generate_list_html(flattened_lists)
    
    # Step 4: Create blocks for each list
    blocks = []
    
    for group_id, group in flattened_lists.items():
        # Get the HTML content for this group
        html = html_content.get(group_id, "")
        
        # Create a text block with the appropriate list type
        block = create_text_block(group['type'], html)
        
        # Add document position for proper ordering
        block['document_position'] = group['document_position']
        
        blocks.append(block)
        log_message(f"Created {group['type']} block for group {group_id}", "full")
    
    # Sort blocks by document position
    blocks.sort(key=lambda x: x.get('document_position', 0))
    
    return blocks

def integrate_list_blocks(content_blocks, list_blocks):
    """
    Integrate list blocks at the correct positions in the content flow.
    
    Parameters:
        content_blocks (list): Existing content blocks
        list_blocks (list): List blocks to be integrated
    
    Returns:
        list: Combined and properly ordered content blocks
    """
    # If no list blocks, return original content
    if not list_blocks:
        return content_blocks
    
    # Sort list blocks by document position
    sorted_list_blocks = sorted(list_blocks, key=lambda x: x.get('document_position', 0))
    
    # Initialize result list
    result_blocks = []
    list_index = 0
    
    # Iterate through content blocks
    for i, block in enumerate(content_blocks):
        # Get current position (assuming content_blocks are in order)
        current_position = i
        
        # Check if we need to insert list blocks before this content block
        while (list_index < len(sorted_list_blocks) and 
               sorted_list_blocks[list_index].get('document_position', float('inf')) <= current_position):
            # Add the list block
            result_blocks.append(sorted_list_blocks[list_index])
            list_index += 1
        
        # Add the current content block
        result_blocks.append(block)
    
    # Add any remaining list blocks
    while list_index < len(sorted_list_blocks):
        result_blocks.append(sorted_list_blocks[list_index])
        list_index += 1
    
    log_message(f"Integrated {len(list_blocks)} list blocks into {len(content_blocks)} content blocks", "basic")
    
    return result_blocks


def is_caption_paragraph(paragraph, namespaces, style_name=None, previous_element_type=None):
    """
    Determine if a paragraph is likely a caption for the previous element.
    Takes into account the type of the previous element to determine appropriate caption styles.
    """
    if not style_name:
        # Try to get style from paragraph
        style_elem = paragraph.find('.//{%s}pStyle' % namespaces['w'])
        if style_elem is not None:
            style_name = style_elem.get('{%s}val' % namespaces['w'])
        
    if not style_name:
        return False
        
    # Get text content for analysis
    text_elements = paragraph.findall('.//{%s}t' % namespaces['w'])
    text_content = "".join([t.text or "" for t in text_elements])
    
    # Log all potential caption paragraphs for debugging
    log_message(f"DEBUG: Checking potential caption: style={style_name}, content={text_content[:30]}, prev_type={previous_element_type}", "basic")
    
    # Style must explicitly be a caption style to be considered
    style_lower = style_name.lower()
    is_caption_style = 'caption' in style_lower or 'popisek' in style_lower
    
    # If it's not explicitly a caption style, it's not a caption
    if not is_caption_style:
        log_message(f"DEBUG: Not a caption style: {style_name}", "basic")
        return False
    
    # If this is a caption style and follows an image or table, consider it a caption
    if previous_element_type in ['image', 'table'] and is_caption_style:
        log_message(f"DEBUG: Found caption with style '{style_name}' for {previous_element_type}", "basic")
        return True
        
    # For all other cases, not a caption
    log_message(f"DEBUG: Caption check result: caption_style={is_caption_style}, final=False", "basic")
    return False


def get_paragraph_list_info(paragraph, namespaces):
    """
    Enhanced function to determine if a paragraph is a list item and its properties.
    More aggressively detects ordered vs. unordered lists.
    
    Parameters:
        paragraph: A paragraph element from DOCX
        namespaces: Dictionary of XML namespaces
    
    Returns:
        Dictionary with list information
    """
    result = {'is_list_item': False, 'level': 0, 'list_type': None, 'num_id': None}
    
    # Find numPr element (indicates a list item)
    num_pr = paragraph.find('.//{%s}numPr' % namespaces['w'])
    if num_pr is None:
        return result
    
    # Get list level (ilvl)
    ilvl = num_pr.find('.//{%s}ilvl' % namespaces['w'])
    level = 0
    if ilvl is not None:
        try:
            level = int(ilvl.get('{%s}val' % namespaces['w'], '0'))
        except ValueError:
            level = 0
    
    # Get list ID (numId)
    num_id = num_pr.find('.//{%s}numId' % namespaces['w'])
    if num_id is None:
        return result
        
    num_id_val = num_id.get('{%s}val' % namespaces['w'])
    if not num_id_val:
        return result
    
    # Extract text content to help determine list type
    text_elements = paragraph.findall('.//{%s}t' % namespaces['w'])
    text_content = "".join([t.text or "" for t in text_elements])
    text_clean = text_content.strip()
    
    # Stronger detection for ordered lists
    is_ordered = False
    
    # Check for explicit ordered list indicators
    if 'order' in text_clean.lower() or 'number' in text_clean.lower():
        is_ordered = True
        log_message(f"Detected ordered list from keywords: '{text_clean[:30]}...'", "full")
    
    # Check for number patterns at start of text
    elif re.match(r'^(\d+|[a-zA-Z]|[ivxIVX]+)[\.\)\:]', text_clean):
        is_ordered = True
        log_message(f"Detected ordered list from pattern: '{text_clean[:30]}...'", "full")
    
    # Check for numbering properties in the XML
    else:
        # Look for numFmt element in related numbering definition
        # This is a simplification - in a real implementation, you'd need to trace
        # through the numbering definitions more thoroughly
        if 'decimal' in str(paragraph) or 'number' in str(paragraph):
            is_ordered = True
            log_message(f"Detected ordered list from paragraph properties", "full")
    
    list_type = 'numbered-list' if is_ordered else 'bullet-list'
    
    return {
        'is_list_item': True, 
        'level': level, 
        'list_type': list_type,
        'num_id': num_id_val
    }

def convert_list_for_storymap(list_data):
    """
    Convert complex nested lists to StoryMap-compatible format.
    Ensures:
    1. Lists have at most 2 levels (main and sub)
    2. All items within a list are the same type
    
    Returns the modified list_data with appropriate warnings logged.
    """
    if not list_data['items'] or list_data['type'] is None:
        return list_data
    
    # Find the minimum level (root level)
    min_level = min(item['level'] for item in list_data['items'])
    
    # Create a proper tree structure to track relationships
    tree = []
    current_node = None
    current_level = None
    prev_level = None
    
    # First create a tree structure preserving parent-child relationships
    for item in list_data['items']:
        level = item['level']
        
        if level == min_level:
            # Root level item
            node = {'text': item['text'], 'children': [], 'level': level}
            tree.append(node)
            current_node = node
            prev_level = level
        elif level == min_level + 1 and current_node:
            # Direct child of current root
            child = {'text': item['text'], 'children': [], 'level': level}
            current_node['children'].append(child)
            current_level = child
            prev_level = level
        elif level > min_level + 1 and current_level:
            # This is a deeper level that needs to be flattened
            # Process the text to indicate its deeper nesting
            indent = "—" * (level - (min_level + 1)) + " "
            flattened_text = indent + item['text']
            
            # Add as a direct child of the current level 2 item
            if prev_level == min_level + 1:
                child = {'text': flattened_text, 'children': [], 'level': min_level + 1}
                current_node['children'].append(child)
            else:
                # If prev_level is higher, this is a continuation of flattened items
                # Add to the same parent as the previous item
                child = {'text': flattened_text, 'children': [], 'level': min_level + 1}
                current_node['children'].append(child)
            
            log_message(
                f"Flattened deeply nested list item at level {level}: '{item['text'][:30]}...'", 
                "basic", is_warning=True
            )
            
            prev_level = level
    
    # If we have more than 2 distinct levels, log a warning
    all_levels = set(item['level'] for item in list_data['items'])
    if len(all_levels) > 2:
        log_message(
            f"List has {len(all_levels)} distinct nesting levels. StoryMap only supports 2 levels. "
            "Deeper levels have been flattened.", 
            "basic", is_warning=True
        )
    
    # Check if there are items of a different type than the main list
    diff_type_items = [item for item in list_data['items'] 
                      if item.get('type') and item.get('type') != list_data['type']]
    if diff_type_items:
        log_message(
            f"List contains mixed types (both ordered and unordered). "
            f"StoryMap requires consistent list types. All items will be set to '{list_data['type']}'.", 
            "basic", is_warning=True
        )
    
    # Convert the tree back to a flattened structure with proper relationships
    new_items = []
    for node in tree:
        new_items.append({'text': node['text'], 'level': node['level']})
        for child in node['children']:
            new_items.append({'text': child['text'], 'level': child['level']})
    
    list_data['items'] = new_items
    return list_data


def process_list_structure(list_data):
    """
    Process a list structure to create proper HTML for StoryMap lists.
    
    Ensures:
    1. Lists are maximum 2 levels deep
    2. All items within a list are the same type
    3. HTML structure matches StoryMap format
    """
    if not list_data['items'] or list_data['type'] is None:
        return None
    
    # First apply StoryMap compatibility conversion
    list_data = convert_list_for_storymap(list_data)
    
    # Determine the list type
    list_type = list_data['type']
    log_message(f"Processing list structure of type: {list_type} with {len(list_data['items'])} items", "full")
    
    # Find the minimum level (root level)
    min_level = min(item['level'] for item in list_data['items'])
    
    # Build the hierarchical structure
    hierarchy = []
    current_parent = None
    
    # First pass: build parent-child relationships
    for item in list_data['items']:
        level = item['level']
        
        if level == min_level:
            # This is a root level item
            current_parent = {'text': item['text'], 'children': []}
            hierarchy.append(current_parent)
        elif current_parent and level == min_level + 1:
            # This is a direct child of the current parent
            current_parent['children'].append(item['text'])
    
    # Generate the HTML according to StoryMap format
    html = ""
    for node in hierarchy:
        # Start list item (no ul/ol container at root level)
        html += f"<li>{node['text']}"
        
        # Add sublist if there are children
        if node['children']:
            # Use appropriate tag based on parent list type
            tag = 'ol' if list_type == 'numbered-list' else 'ul'
            html += f"<{tag}>"
            
            # Add child items
            for child_text in node['children']:
                html += f"<li>{child_text}</li>"
            
            # Close sublist
            html += f"</{tag}>"
        
        # Close main list item
        html += "</li>"
    
    log_message(f"Generated HTML for {list_type}: {html[:100]}...", "full")
    
    # Create the text block with the appropriate list type
    return create_text_block(list_type, html)


    """
    Process a list structure to create proper HTML for StoryMap lists.
    
    Ensures:
    1. Lists are maximum 2 levels deep
    2. All items within a list are the same type
    3. HTML structure matches StoryMap format
    """
    if not list_data['items'] or list_data['type'] is None:
        return None
    
    # First apply StoryMap compatibility conversion
    list_data = convert_list_for_storymap(list_data)
    
    # Determine the list type
    list_type = list_data['type']
    log_message(f"Processing list structure of type: {list_type} with {len(list_data['items'])} items", "full")
    
    # Find the minimum level (root level)
    min_level = min(item['level'] for item in list_data['items'])
    
    # Build the hierarchical structure
    hierarchy = []
    current_parent = None
    
    # First pass: build parent-child relationships
    for item in list_data['items']:
        level = item['level']
        
        if level == min_level:
            # This is a root level item
            current_parent = {'text': item['text'], 'children': []}
            hierarchy.append(current_parent)
        elif current_parent and level == min_level + 1:
            # This is a direct child of the current parent
            current_parent['children'].append(item['text'])
    
    # Generate the HTML according to StoryMap format
    html = ""
    for node in hierarchy:
        # Start list item (no ul/ol container at root level)
        html += f"<li>{node['text']}"
        
        # Add sublist if there are children
        if node['children']:
            # Use appropriate tag based on parent list type
            tag = 'ol' if list_type == 'numbered-list' else 'ul'
            html += f"<{tag}>"
            
            # Add child items
            for child_text in node['children']:
                html += f"<li>{child_text}</li>"
            
            # Close sublist
            html += f"</{tag}>"
        
        # Close main list item
        html += "</li>"
    
    log_message(f"Generated HTML for {list_type}: {html[:100]}...", "full")
    
    # Create the text block with the appropriate list type
    return create_text_block(list_type, html)
    
def process_docx_paragraph(paragraph, namespaces, image_rels, media_files, hyperlink_rels=None):
    """Process a DOCX paragraph element with language-independent heading detection."""
    if hyperlink_rels is None:
        hyperlink_rels = {}
    
    # Check for paragraph style
    style_elem = paragraph.find('.//{%s}pStyle' % namespaces['w'])
    style = None
    style_id = None
    if style_elem is not None:
        style = style_elem.get('{%s}val' % namespaces['w'])
        style_id = style  # Save the original style ID
        log_message(f"Processing DOCX paragraph with style ID: {style}", "full")
    
    # Check for images
    drawing = paragraph.find('.//{%s}drawing' % namespaces['w'])
    if drawing is not None:
        log_message("Paragraph contains drawing/image", "full")
        return process_docx_image(drawing, namespaces, image_rels, media_files)
    
    # Process text content with rich formatting
    text_tuple = extract_formatted_text(paragraph, namespaces, hyperlink_rels)
    text = text_tuple[0]
    alignment = text_tuple[1]
    
    # Skip empty paragraphs
    if not text or not text.strip():
        if alignment:
            log_message(f"Empty paragraph with alignment: {alignment}", "full")
            return create_text_block('paragraph', '', alignment)
        log_message("Skipping empty paragraph", "full")
        return None
    
    # Check outline level - this is language-independent
    outline_level = get_paragraph_outline_level(paragraph, namespaces)
    if outline_level is not None:
        log_message(f"Found outline level: {outline_level}", "full")
        if outline_level == 0:
            log_message(f"Created heading (h2) from outline level: {text[:30]}...", "full")
            return create_text_block('h2', text, alignment)
        elif outline_level == 1:
            log_message(f"Created subheading (h3) from outline level: {text[:30]}...", "full")
            return create_text_block('h3', text, alignment)
        elif outline_level == 2:
            log_message(f"Created subheading (h4) from outline level: {text[:30]}...", "full")
            return create_text_block('h4', text, alignment)
    
    # Check for numeric pattern in style ID (language-independent)
    if style_id:
        # Extract any digits from the style ID
        digits = re.findall(r'\d+', style_id)
        if digits:
            heading_level = int(digits[0])
            if heading_level == 1:
                log_message(f"Created heading (h2) from style number: {text[:30]}...", "full")
                return create_text_block('h2', text, alignment)
            elif heading_level == 2:
                log_message(f"Created subheading (h3) from style number: {text[:30]}...", "full")
                return create_text_block('h3', text, alignment)
            elif heading_level == 3:
                log_message(f"Created subheading (h4) from style number: {text[:30]}...", "full")
                return create_text_block('h4', text, alignment)
    
    # Check paragraph formatting attributes that suggest a heading
    is_heading = check_heading_formatting(paragraph, namespaces)
    if is_heading:
        # Determine heading level based on formatting
        font_size = get_font_size(paragraph, namespaces)
        if font_size:
            log_message(f"Found font size: {font_size}", "full")
            if font_size >= 20:
                log_message(f"Created heading (h2) from font size: {text[:30]}...", "full")
                return create_text_block('h2', text, alignment)
            elif font_size >= 16:
                log_message(f"Created subheading (h3) from font size: {text[:30]}...", "full")
                return create_text_block('h3', text, alignment)
            elif font_size >= 14:
                log_message(f"Created subheading (h4) from font size: {text[:30]}...", "full")
                return create_text_block('h4', text, alignment)
    
    # Check other style attributes often used for headings
    if style:
        style_lower = style.lower()
        # Check for common heading keywords in any language
        if 'title' in style_lower or 'heading' in style_lower or 'nadpis' in style_lower:
            log_message(f"Created heading (h2) from style keywords: {text[:30]}...", "full")
            return create_text_block('h2', text, alignment)
        
        # Check for quote styles
        if 'quote' in style_lower or 'citation' in style_lower:
            log_message(f"Created quote: {text[:30]}...", "full")
            return create_text_block('quote', text, alignment)
        
        # Check for code styles
        if 'code' in style_lower or 'source' in style_lower:
            log_message(f"Created code block: {text[:30]}...", "full")
            return create_code_block(text)
    
    # Check content characteristics
    if len(text.strip()) < 50:
        # Very short paragraphs might be headings
        if text.strip().isupper():
            log_message(f"Created heading (h2) from all-caps text: {text[:30]}...", "full")
            return create_text_block('h2', text, alignment)
        # Short paragraphs that end without punctuation might be headings
        if not re.search(r'[.!?]$', text.strip()):
            is_bold = check_is_bold(paragraph, namespaces)
            if is_bold:
                log_message(f"Created heading (h3) from short bold text: {text[:30]}...", "full")
                return create_text_block('h3', text, alignment)
    
    # Default to paragraph
    log_message(f"Created paragraph: {text[:30]}...", "full")
    return create_text_block('paragraph', text, alignment)

def get_paragraph_outline_level(paragraph, namespaces):
    """Get the outline level of a paragraph (language-independent)."""
    outline_lvl = paragraph.find('.//{%s}outlineLvl' % namespaces['w'])
    if outline_lvl is not None:
        try:
            level = int(outline_lvl.get('{%s}val' % namespaces['w'], '0'))
            return level
        except (ValueError, TypeError):
            pass
    return None

def check_heading_formatting(paragraph, namespaces):
    """Check if paragraph has formatting typically used for headings."""
    # Headings are often bold
    is_bold = check_is_bold(paragraph, namespaces)
    
    # Headings often have larger font size
    font_size = get_font_size(paragraph, namespaces)
    has_large_font = font_size and font_size >= 14
    
    # Headings often have special spacing
    spacing = paragraph.find('.//{%s}spacing' % namespaces['w'])
    has_special_spacing = spacing is not None
    
    # Return True if the paragraph has at least some heading-like formatting
    return is_bold or has_large_font or has_special_spacing

def check_is_bold(paragraph, namespaces):
    """Check if paragraph is bold."""
    # Check if all runs in the paragraph are bold
    runs = paragraph.findall('.//{%s}r' % namespaces['w'])
    if not runs:
        return False
    
    bold_runs = 0
    for run in runs:
        if run.find('.//{%s}b' % namespaces['w']) is not None:
            bold_runs += 1
    
    # If more than half the runs are bold, consider it a bold paragraph
    return bold_runs > len(runs) / 2

def get_font_size(paragraph, namespaces):
    """Get the font size of a paragraph (returns None if mixed or not specified)."""
    sizes = set()
    runs = paragraph.findall('.//{%s}r' % namespaces['w'])
    
    for run in runs:
        sz = run.find('.//{%s}sz' % namespaces['w'])
        if sz is not None:
            try:
                # Font size in Word is in half-points (so 24 = 12pt)
                size_half_points = int(sz.get('{%s}val' % namespaces['w'], '0'))
                sizes.add(size_half_points / 2)  # Convert to points
            except (ValueError, TypeError):
                pass
    
    # If all runs have the same size, return it
    if len(sizes) == 1:
        return next(iter(sizes))
    return None

def process_docx_image(drawing, namespaces, image_rels, media_files):
    """Process a DOCX image element with enhanced caption detection."""
    try:
        log_message(f"DEBUG: Processing drawing with tag: {drawing.tag}", "basic")
        
        # Initialize variables
        blip = None
        rel_id = None
        caption = None
        textbox_caption = None
        
        # First try standard blip approach (this worked for image 1 before)
        blip = drawing.find('.//{%s}blip' % namespaces['a'])
        if blip is not None:
            # Get relationship ID directly from blip
            if '{%s}embed' % namespaces['r'] in blip.attrib:
                rel_id = blip.get('{%s}embed' % namespaces['r'])
                log_message(f"DEBUG: Found standard blip with embed ID: {rel_id}", "basic")
            elif '{%s}link' % namespaces['r'] in blip.attrib:
                rel_id = blip.get('{%s}link' % namespaces['r'])
                log_message(f"DEBUG: Found standard blip with link ID: {rel_id}", "basic")
        
        # Get caption from docPr element (standard method)
        doc_pr = drawing.find('.//{%s}docPr' % namespaces['wp'])
        if doc_pr is not None and 'descr' in doc_pr.attrib:
            caption = doc_pr.get('descr')
            log_message(f"DEBUG: Found caption in image description: {caption[:50] if caption else 'None'}...", "basic")
            # Filter out AI-generated captions
            if caption and ('generated' in caption.lower() and ('ai' in caption.lower() or 'intelligence' in caption.lower())):
                log_message("DEBUG: This appears to be AI-generated alt text, ignoring", "basic")
                caption = None
        
        # Look for textbox caption (for image 2)
        txbx_elements = []
        for elem in drawing.iter():
            if elem.tag.endswith('txbx') or elem.tag.endswith('txbxContent'):
                txbx_elements.append(elem)
                
        log_message(f"DEBUG: Found {len(txbx_elements)} potential textbox elements", "basic")
        
        for textbox in txbx_elements:
            # Find paragraphs within the textbox
            paras = []
            for child in textbox.iter():
                if child.tag.endswith('}p'):
                    paras.append(child)
            
            log_message(f"DEBUG: Found {len(paras)} paragraphs in textbox", "basic")
            
            for para in paras:
                # Check if it's a caption style
                style = para.find('.//{%s}pStyle' % namespaces['w'])
                is_caption_style = False
                if style is not None:
                    style_val = style.get('{%s}val' % namespaces['w'], '')
                    log_message(f"DEBUG: Paragraph in textbox has style: {style_val}", "basic")
                    if 'caption' in style_val.lower() or 'titulek' in style_val.lower():
                        is_caption_style = True
                
                # Extract the text
                text_elements = para.findall('.//{%s}t' % namespaces['w'])
                text_content = "".join([t.text or "" for t in text_elements])
                if text_content:
                    log_message(f"DEBUG: Textbox paragraph content: {text_content[:50]}...", "basic")
                    
                    # If it has caption style or starts with Figure/Image
                    if is_caption_style or text_content.strip().lower().startswith(('figure', 'image', 'obr')):
                        textbox_caption = text_content
                        log_message(f"DEBUG: Found caption in textbox: {text_content[:50]}...", "basic")
                        break
            
            # Break out if we found a caption
            if textbox_caption:
                break
        
        # For image 2 (with textbox caption but missing blip)
        # If we have a textbox caption but no relationship ID yet, search more aggressively
        if textbox_caption and not rel_id:
            log_message("DEBUG: Found textbox caption but no relationship ID yet, searching in alternate content", "basic")
            
            # Try to find the relationship ID in the alternate content
            alt_content = drawing.find('.//mc:AlternateContent', namespaces)
            if alt_content is not None:
                log_message("DEBUG: Found mc:AlternateContent element", "basic")
                
                # Try to find any relationship ID attribute in the entire structure
                for elem in alt_content.iter():
                    for attr_name, attr_val in elem.attrib.items():
                        if attr_name.endswith('}id') or attr_name.endswith('}embed'):
                            if attr_val in image_rels:
                                rel_id = attr_val
                                log_message(f"DEBUG: Found relationship ID {rel_id} in alternate content", "basic")
                                break
                    if rel_id:
                        break
            
            # If still no relationship ID found, search in the entire drawing
            if not rel_id:
                log_message("DEBUG: Searching entire drawing XML for relationship IDs", "basic")
                drawing_xml = ET.tostring(drawing, encoding='unicode')
                
                # Check for any relationship IDs in the drawing text
                for potential_id in image_rels.keys():
                    id_pattern = f'"{potential_id}"'  # Look for "rId11" pattern
                    if id_pattern in drawing_xml:
                        rel_id = potential_id
                        log_message(f"DEBUG: Found relationship ID {rel_id} in XML text", "basic")
                        break
        
        # If we have a valid relationship ID but no blip, we'll create our own
        if rel_id in image_rels and not blip:
            log_message(f"DEBUG: Creating custom blip handler for relationship {rel_id}", "basic")
            # We'll handle this later in the code
        
        # If we still don't have a relationship ID but have a textbox caption,
        # try to match with any unused image relationship
        if not rel_id and textbox_caption:
            # Find all image relationships that have been used so far
            used_ids = []
            for block in blocks if 'blocks' in locals() else []:
                if block.get('type') == 'image' and 'original_path' in block:
                    for rel_id, path in image_rels.items():
                        if path in block['original_path']:
                            used_ids.append(rel_id)
            
            # Find any unused image relationship
            for potential_id in image_rels:
                if potential_id not in used_ids:
                    rel_id = potential_id
                    log_message(f"DEBUG: Using unused image relationship ID {rel_id}", "basic")
                    break
        
        # If we still don't have a relationship ID, we can't process this image
        if not rel_id or rel_id not in image_rels:
            log_message(f"DEBUG: No valid relationship ID found. Available IDs: {list(image_rels.keys())}", "basic", is_warning=True)
            return None
        
        # Get the image path from the relationship
        image_path = image_rels[rel_id]
        log_message(f"DEBUG: Found image path from relationship ID {rel_id}: {image_path}", "basic")
        
        if not image_path.startswith('media/'):
            log_message(f"DEBUG: Image path does not start with 'media/': {image_path}", "basic", is_warning=True)
            return None
        
        # Get the actual file path
        file_path = media_files.get(image_path)
        if not file_path or not os.path.exists(file_path):
            log_message(f"DEBUG: Image file not found: {file_path}", "basic", is_warning=True)
            return None
        
        # Use the original image name
        original_filename = os.path.basename(image_path)
        temp_img_path = os.path.join(tempfile.gettempdir(), original_filename)
        
        try:
            shutil.copy2(file_path, temp_img_path)
            log_message(f"DEBUG: Copied image to temporary location: {original_filename}", "basic")
        except Exception as copy_err:
            log_message(f"DEBUG: Error copying image: {str(copy_err)}", "basic", is_warning=True)
            return None
        
        # Determine display properties
        display, float_alignment = determine_image_display(temp_img_path, drawing, namespaces)

        
        # Get original dimensions
        dimensions = None
        try:
            from PIL import Image as PILImage
            with PILImage.open(temp_img_path) as img:
                width, height = img.size
                dimensions = (width, height)
                log_message(f"DEBUG: Image dimensions: {width}x{height}", "basic")
        except Exception as img_err:
            log_message(f"DEBUG: Error getting image dimensions: {str(img_err)}", "basic", is_warning=True)
        
        # Create image block with dimensions
        image_block = create_image_block(temp_img_path, caption=textbox_caption or caption, display=display, float_alignment=float_alignment)
        if dimensions:
            image_block['dimensions'] = dimensions
        
        # Add the original path for reference
        image_block['original_path'] = image_path
        
        # Use textbox caption if available, otherwise use description caption
        if textbox_caption:
            image_block['caption'] = textbox_caption
            log_message(f"DEBUG: Added textbox caption to image block: {textbox_caption[:50]}...", "basic")
        elif caption:
            image_block['caption'] = caption
            log_message(f"DEBUG: Added description caption to image block: {caption[:50]}...", "basic")
        
        log_message(f"DEBUG: Created image block from: {original_filename} with display: {display}", "basic")
        return image_block
        
    except Exception as e:
        log_message(f"DEBUG: Error processing image: {str(e)}", "basic", is_warning=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_warning=True)
        return None

def determine_image_display(image_path=None, drawing=None, namespaces=None):
    """
    Determine the optimal display setting for an image based on its dimensions and context.
    
    Args:
        image_path (str, optional): Path to the image file to analyze
        drawing (Element, optional): Drawing element from DOCX XML
        namespaces (dict, optional): XML namespaces for DOCX
        
    Returns:
        tuple: (display_type, float_alignment) - float_alignment will be "start" or "end" if applicable
    """
    # Default to "standard" if we can't determine
    display_type = "standard"
    float_alignment = None
    
    try:
        # Check for text wrapping in the DOCX XML (indicates "float")
        is_wrapped = False
        
        if drawing is not None and namespaces is not None:
            # Check for wrapSquare, wrapTight, or similar elements indicating wrapping
            wrap_elems = []
            for wrap_type in ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopBottom']:
                wrap_elem = drawing.find('.//{%s}%s' % (namespaces['wp'], wrap_type))
                if wrap_elem is not None:
                    wrap_elems.append(wrap_elem)
                    
            if wrap_elems:
                is_wrapped = True
                log_message(f"DEBUG: Image has text wrapping: {[elem.tag.split('}')[-1] for elem in wrap_elems]}", "basic")
                
                # Default to right alignment (StoryMap's "end")
                float_alignment = "end"
                
                # Try to find explicit alignment information
                pos_h = drawing.find('.//{%s}positionH' % namespaces['wp'])
                if pos_h is not None:
                    # Check for alignment value
                    align_elem = pos_h.find('.//{%s}align' % namespaces['wp'])
                    if align_elem is not None and align_elem.text:
                        align_val = align_elem.text.lower()
                        log_message(f"DEBUG: Found explicit alignment: {align_val}", "basic")
                        if align_val == "left":
                            float_alignment = "start"
                        elif align_val == "right":
                            float_alignment = "end"
                    
                    # Check relativeFrom attribute
                    if 'relativeFrom' in pos_h.attrib:
                        rel_from = pos_h.get('relativeFrom')
                        log_message(f"DEBUG: Position relative from: {rel_from}", "basic")
                        # Some relative positions imply left/right alignment
                        if rel_from in ['right', 'rightMargin']:
                            float_alignment = "end"
                        elif rel_from in ['left', 'leftMargin']:
                            float_alignment = "start"
                
                log_message(f"DEBUG: Determined float alignment: {float_alignment}", "basic")
        
        # Get dimensions from the image file
        width = height = 0
        aspect_ratio = 0
        
        from PIL import Image as PILImage
        if image_path and os.path.exists(image_path):
            with PILImage.open(image_path) as img:
                width, height = img.size
                if height > 0:
                    aspect_ratio = width / height
                log_message(f"DEBUG: Image dimensions: {width}x{height}, aspect ratio: {aspect_ratio:.2f}", "basic")
        
        # Apply the display rules
        if width > 1200 and aspect_ratio >= (16/9):
            display_type = "wide"
            log_message(f"DEBUG: Setting display to 'wide' (width > 1200 and aspect ratio ≥ 16:9)", "basic")
        elif width < 800 or is_wrapped:
            display_type = "float"
            log_message(f"DEBUG: Setting display to 'float' with alignment '{float_alignment}'", "basic")
        else:
            display_type = "standard"
            log_message(f"DEBUG: Setting display to 'standard' (default case)", "basic")
            
        return display_type, float_alignment
        
    except Exception as e:
        log_message(f"DEBUG: Error determining image display: {str(e)}", "basic", is_warning=True)
        return "standard", None  # Default to standard on error 
      
def process_docx_table(element, namespaces, hyperlink_rels=None):
    """Process a DOCX table element with proper hyperlink support."""
    if hyperlink_rels is None:
        hyperlink_rels = {}
        
    rows = []
    
    try:
        # Process each row
        for row in element.findall('.//{%s}tr' % namespaces['w']):
            cells = []
            
            # Process each cell
            for cell in row.findall('.//{%s}tc' % namespaces['w']):
                cell_content = ""
                
                # Extract text from paragraphs in the cell with hyperlinks
                for paragraph in cell.findall('.//{%s}p' % namespaces['w']):
                    # Use the full extraction function with hyperlink processing
                    para_text, _ = extract_formatted_text(paragraph, namespaces, hyperlink_rels)
                    
                    if para_text:
                        if cell_content:
                            cell_content += "\n"
                        cell_content += para_text
                
                cells.append(cell_content)
            
            if cells:
                rows.append(cells)
        
        if rows:
            log_message(f"Created table with {len(rows)} rows and {len(rows[0])} columns", "full")
            return create_table_block(rows)
        
        log_message("No rows found in table", "full", is_warning=True)
        return None
    except Exception as e:
        log_message(f"Error processing table: {str(e)}", "basic", is_warning=True)
        import traceback
        log_message(traceback.format_exc(), "full", is_warning=True)
        return None

def extract_formatted_text(element, namespaces, hyperlink_rels=None):
    """Extract text with rich formatting from a DOCX element."""
    if hyperlink_rels is None:
        hyperlink_rels = {}
        
    text_parts = []
    paragraph_alignment = None
    
    # Check paragraph alignment
    jc_elem = element.find('.//{%s}jc' % namespaces['w'])
    if jc_elem is not None:
        alignment_val = jc_elem.get('{%s}val' % namespaces['w'])
        if alignment_val:
            paragraph_alignment = alignment_val  # center, right, left, justify
            log_message(f"Found paragraph alignment: {alignment_val}", "full")
    
    # Handle hyperlinks differently - find them first at paragraph level
    hyperlinks = element.findall('.//{%s}hyperlink' % namespaces['w'])
    hyperlink_runs = {}
    
    # Map hyperlink relationship IDs to their runs
    for hyperlink in hyperlinks:
        rel_id = hyperlink.get('{%s}id' % namespaces['r'])
        if rel_id and rel_id in hyperlink_rels:
            # Get the actual URL from relationships
            target_url = hyperlink_rels[rel_id]
            
            # Find all runs within this hyperlink
            for run in hyperlink.findall('.//{%s}r' % namespaces['w']):
                run_text = "".join([t.text or "" for t in run.findall('.//{%s}t' % namespaces['w'])])
                hyperlink_runs[id(run)] = (target_url, run_text)
                log_message(f"Found hyperlink: {run_text} -> {target_url}", "full")
    
    # Process each run
    has_text = False
    run_count = 0
    
    for run in element.findall('.//{%s}r' % namespaces['w']):
        run_count += 1
        # Check for different formatting properties
        formatting = {
            'bold': run.find('.//{%s}b' % namespaces['w']) is not None,
            'italic': run.find('.//{%s}i' % namespaces['w']) is not None,
            'underline': run.find('.//{%s}u' % namespaces['w']) is not None,
            'strike': run.find('.//{%s}strike' % namespaces['w']) is not None,
        }
        
        # Check for vertical alignment (subscript/superscript)
        vert_align = run.find('.//{%s}vertAlign' % namespaces['w'])
        if vert_align is not None:
            val = vert_align.get('{%s}val' % namespaces['w'])
            if val == 'subscript':
                formatting['sub'] = True
            elif val == 'superscript':
                formatting['sup'] = True
        
        # Check for color
        color_elem = run.find('.//{%s}color' % namespaces['w'])
        text_color = None
        if color_elem is not None:
            color_val = color_elem.get('{%s}val' % namespaces['w'])
            if color_val and color_val.lower() != 'auto':
                text_color = color_val
                log_message(f"Found text color: {text_color}", "full")
        
        # Get text
        run_text = ""
        text_elements = run.findall('.//{%s}t' % namespaces['w'])
        for text_elem in text_elements:
            if text_elem.text is not None:
                run_text += text_elem.text
                has_text = True
        
        # Check for line breaks and special characters
        br_elem = run.find('.//{%s}br' % namespaces['w'])
        if br_elem is not None:
            run_text += "\n"
            log_message("Found line break", "full")
        
        # Apply formatting
        if run_text:
            formatted_text = run_text
            
            # Check if this run is part of a hyperlink
            run_id = id(run)
            if run_id in hyperlink_runs:
                url, link_text = hyperlink_runs[run_id]
                # Make sure URL is properly formed
                if url and not url.startswith(('http://', 'https://', '#')):
                    url = 'https://' + url
                
                formatted_text = f'<a href="{url}" rel="noopener noreferrer" target="_blank">{formatted_text}</a>'
                log_message(f"Applied hyperlink formatting with URL: {url}", "full")
            
            # Apply formatting in specific order to handle nesting correctly
            if formatting.get('sub', False):
                formatted_text = f"<sub>{formatted_text}</sub>"
            if formatting.get('sup', False):
                formatted_text = f"<sup>{formatted_text}</sup>"
            if formatting.get('italic'):
                formatted_text = f"<em>{formatted_text}</em>"
            if formatting.get('bold'):
                formatted_text = f"<strong>{formatted_text}</strong>"
            if formatting.get('underline'):
                formatted_text = f"<u>{formatted_text}</u>"
            if formatting.get('strike'):
                formatted_text = f"<s>{formatted_text}</s>"
            
            # Apply color if present
            if text_color:
                formatted_text = f'<span class="sm-text-color-{text_color}">{formatted_text}</span>'
            
            text_parts.append(formatted_text)
    
    # Return a non-empty string even for empty paragraphs with alignment
    if not has_text and paragraph_alignment:
        log_message("Empty paragraph with alignment", "full")
        return " ", paragraph_alignment
    
    log_message(f"Extracted text with {run_count} runs", "full")
    return "".join(text_parts), paragraph_alignment


def create_storymap(gis, title, tags, summary, description, cover_image, content_blocks):
    """Create a StoryMap with placeholders."""
    try:
        # Unpack content_blocks if it's a tuple containing the blocks and temp_dir
        if isinstance(content_blocks, tuple) and len(content_blocks) == 2:
            content_blocks, temp_dir = content_blocks
            log_message(f"Using temporary directory for content: {temp_dir}", "full")
        
        # Create new StoryMap
        log_message("Creating new StoryMap instance", "basic")
        story = StoryMap(gis=gis)
        
        # Set up placeholders and mapping
        placeholder_ids = {}
        parsed_blocks = {}
        image_dimensions = {}  # Track image dimensions for updating resources later
        
        # Process cover image
        process_cover_image(story, title, summary, cover_image, image_dimensions)
        
        # Add description as first text block if provided
        if description:
            add_description_block(story, description, placeholder_ids, parsed_blocks)

        # Add content blocks with placeholders
        log_message(f"Adding {len(content_blocks)} content blocks to StoryMap", "basic")
        blocks_added = 0
        
        for i, block in enumerate(content_blocks):
            block_type = block.get('type')
            placeholder_key = f"{block_type}_{i}"
            
            log_message(f"Adding block {i+1}/{len(content_blocks)}: {block_type}", "full")
            
            if block_type == 'separator':
                add_separator_block(story, block, placeholder_key, placeholder_ids, parsed_blocks)
                blocks_added += 1
                
            elif block_type == 'image':
                add_image_block(story, block, placeholder_key, placeholder_ids, parsed_blocks, image_dimensions)
                blocks_added += 1
                
            elif block_type == 'text':
                # Skip empty text blocks
                text_content = block.get('text', '')
                if not text_content or text_content.strip() == "" or text_content == "...":
                    log_message(f"Skipping empty text block of type '{block.get('text_type', 'paragraph')}'", "basic")
                    continue
                
                add_text_block(story, block, i, placeholder_key, placeholder_ids, parsed_blocks)
                blocks_added += 1
                
            elif block_type == 'table':
                add_table_block(story, block, placeholder_key, placeholder_ids, parsed_blocks)
                blocks_added += 1
                
            elif block_type == 'code':
                add_code_block(story, block, i, placeholder_key, placeholder_ids, parsed_blocks)
                blocks_added += 1
        
        # Save StoryMap
        log_message("Saving StoryMap...", "basic")
        storymap_item = story.save(title=title, tags=tags)
        log_message(f"StoryMap saved, item ID: {storymap_item.id}", "basic")
        
        # Log statistics
        log_message(f"Added {blocks_added} of {len(content_blocks)} content blocks to StoryMap", "basic")
        blocks_by_type = {}
        for node_id, block in parsed_blocks.items():
            block_type = block.get('type')
            if block_type in blocks_by_type:
                blocks_by_type[block_type] += 1
            else:
                blocks_by_type[block_type] = 1
                
        log_message(f"Content blocks by type: {', '.join([f'{count} {type}(s)' for type, count in blocks_by_type.items()])}", "basic")
        
        return storymap_item, placeholder_ids, parsed_blocks, image_dimensions
        
    except Exception as e:
        log_message(f"Error creating StoryMap: {str(e)}", "none", is_error=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_error=True)
        raise

def process_cover_image(story, title, summary, cover_image, image_dimensions):
    """Process and add cover image to the StoryMap with improved error handling and format conversion."""
    if cover_image and os.path.exists(cover_image):
        log_message(f"Processing cover image: {cover_image}", "basic")
        
        # Use a generic safe name for cover image
        _, extension = os.path.splitext(cover_image)
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        safe_cover_filename = f"coverimage_{timestamp}{extension}"
        temp_cover_path = os.path.join(tempfile.gettempdir(), safe_cover_filename)
        
        try:
            # Try to open and resave the image using PIL to resolve potential format issues
            try:
                from PIL import Image as PILImage
                with PILImage.open(cover_image) as img:
                    # Log image format details for debugging
                    log_message(f"Cover image format: {img.format}, mode: {img.mode}, size: {img.size}", "basic")
                    
                    # Create a new RGB image if needed
                    if img.mode != 'RGB':
                        log_message(f"Converting image from {img.mode} to RGB mode", "basic")
                        img = img.convert('RGB')
                    
                    # Save to a temp location with a more standard format
                    reprocessed_path = os.path.join(tempfile.gettempdir(), f"reprocessed_{safe_cover_filename}")
                    img.save(reprocessed_path, format='JPEG', quality=95)
                    log_message(f"Image reprocessed and saved to {reprocessed_path}", "basic")
                    
                    # Use the reprocessed image instead
                    temp_cover_path = reprocessed_path
                    
                    # Get dimensions from the reprocessed image
                    dimensions = img.size
                    if dimensions:
                        width, height = dimensions
                        image_dimensions['cover_image'] = (width, height)
                        log_message(f"Cover image dimensions: {width}x{height}", "basic")
            except ImportError:
                log_message("PIL/Pillow not available - copying image as-is", "basic", is_warning=True)
                shutil.copy2(cover_image, temp_cover_path)
            except Exception as pil_err:
                log_message(f"Error reprocessing image with PIL: {str(pil_err)}", "basic", is_warning=True)
                # Fall back to direct copy
                shutil.copy2(cover_image, temp_cover_path)
                log_message(f"Copied cover image to: {safe_cover_filename}", "full")
                
                # Get dimensions without resizing
                dimensions = extract_image_dimensions(temp_cover_path)
                if dimensions:
                    width, height = dimensions
                    image_dimensions['cover_image'] = (width, height)
                    log_message(f"Cover image dimensions: {width}x{height}", "basic")
            
            # Create StoryImage object for the cover
            try:
                cover_img = StoryImage(temp_cover_path)
                log_message("Setting cover with image", "basic")
                story.cover(title=title, summary=summary, image=cover_img)
                log_message("Successfully set cover with image", "basic")
            except Exception as e:
                log_message(f"Error setting cover with image: {str(e)}", "basic", is_warning=True)
                # Fall back to cover without image
                story.cover(title=title, summary=summary)
                log_message("Falling back to cover without image", "basic")
        except Exception as e:
            log_message(f"Error processing cover image: {str(e)}", "basic", is_warning=True)
            log_message("Setting cover without image", "basic")
            story.cover(title=title, summary=summary)
    else:
        log_message("Setting cover without image", "basic")
        story.cover(title=title, summary=summary)
        
        
def extract_image_dimensions(image_path):
    """
    Extract dimensions from an image without resizing it.
    
    Args:
        image_path (str): Path to the image file
        
    Returns:
        tuple: (width, height) or None if dimensions couldn't be extracted
    """
    try:
        from PIL import Image as PILImage
        
        # Open the image and get its dimensions
        with PILImage.open(image_path) as img:
            dimensions = img.size
            log_message(f"Image dimensions: {dimensions[0]}x{dimensions[1]}", "full")
            return dimensions
            
    except ImportError:
        log_message("PIL (Pillow) library not found. Image dimension extraction skipped.", "basic", is_warning=True)
        return None
    except Exception as e:
        log_message(f"Error extracting image dimensions: {str(e)}. Using original image.", "basic", is_warning=True)
        return None
        
        
def add_description_block(story, description, placeholder_ids, parsed_blocks):
    """Add description as first text block."""
    log_message("Adding description as first text block", "basic")
    text = Text(description, TextStyles.PARAGRAPH)
    node_id = story.add(text)
    placeholder_ids["description"] = node_id
    parsed_blocks[node_id] = {"type": "text", "text_type": "paragraph", "text": description}
    log_message(f"Added description text, node ID: {node_id}", "full")


def add_separator_block(story, block, placeholder_key, placeholder_ids, parsed_blocks):
    """Add a separator block to the StoryMap."""
    node_id = story.add()
    placeholder_ids[placeholder_key] = node_id
    parsed_blocks[node_id] = block
    log_message(f"Added separator, node ID: {node_id}", "full")

def add_image_block(story, block, placeholder_key, placeholder_ids, parsed_blocks, image_dimensions):
    """Add an image block to the StoryMap with proper float alignment."""
    # Get image path
    image_path = block.get('path')
    log_message(f"Processing image: {image_path}", "full")
    
    # Get dimensions without resizing
    dimensions = extract_image_dimensions(image_path)
    if dimensions:
        width, height = dimensions
        image_dimensions[image_path] = (width, height)
    
    # Get properties
    display = block.get('display', 'standard')
    float_alignment = block.get('float_alignment')
    caption = block.get('caption', '')
    alt_text = block.get('alt_text', '')
    position = block.get('position')
    
    # Create StoryImage object
    img = StoryImage(image_path)
    
    # For floating images, need to set alignment property on the image object if possible
    # Otherwise, we'll store it to apply later in the JSON update
    
    # Add the image to the story - position is always passed as the 5th parameter, not float_alignment
    node_id = story.add(img, caption, alt_text, display, position)
    
    # Store all properties for later JSON update
    placeholder_ids[placeholder_key] = node_id
    parsed_blocks[node_id] = block
    parsed_blocks[node_id]['original_path'] = image_path
    
    # Store float alignment for JSON update later
    if display == 'float' and float_alignment:
        parsed_blocks[node_id]['float_alignment'] = float_alignment
        log_message(f"Stored float alignment '{float_alignment}' for later JSON update", "full")
    
    if dimensions:
        parsed_blocks[node_id]['dimensions'] = dimensions
    
    log_message(f"Added image with display '{display}', node ID: {node_id}", "full")
    return node_id
    


def add_text_block(story, block, i, placeholder_key, placeholder_ids, parsed_blocks):
    """Add a text block to the StoryMap, skipping empty blocks."""
    text_type = block.get('text_type', 'paragraph')
    text_content = block.get('text', '')  # Get actual text content
    
    # Skip empty text blocks
    if not text_content or text_content.strip() == "" or text_content == "...":
        log_message(f"Skipping empty text block of type '{text_type}'", "basic")
        return None  # Return None for skipped blocks
    
    placeholder_text = f"PLACEHOLDER_TEXT_{i}"
    
    # Convert text_type to valid TextStyles enum value
    if text_type == 'bullet-list':
        text_style = TextStyles.BULLETLIST
    elif text_type == 'numbered-list':
        text_style = TextStyles.NUMBERLIST
    elif text_type == 'h2':
        text_style = TextStyles.HEADING
    elif text_type == 'h3':
        text_style = TextStyles.HEADING2
    elif text_type == 'h4':
        text_style = TextStyles.HEADING3
    elif text_type == 'quote':
        text_style = TextStyles.QUOTE
    elif text_type == 'large-paragraph':
        text_style = TextStyles.LARGEPARAGRAPH
    else:
        text_style = TextStyles.PARAGRAPH
    
    log_message(f"Adding text of type '{text_type}' with placeholder", "full")
    text = Text(placeholder_text, text_style)
    node_id = story.add(text)
    
    placeholder_ids[placeholder_key] = node_id
    # Store the actual text content in parsed_blocks
    parsed_blocks[node_id] = {
        "type": "text", 
        "text_type": text_type,
        "text": text_content,
        "alignment": block.get('alignment', None)
    }
    log_message(f"Added text placeholder, node ID: {node_id}", "full")
    return node_id

def add_table_block(story, block, placeholder_key, placeholder_ids, parsed_blocks):
    """Add a table block to the StoryMap."""
    rows = block.get('rows', [])
    num_rows = len(rows)
    num_cols = max(len(row) for row in rows) if rows else 2
    caption = block.get('caption', '')
    
    log_message(f"Adding table with {num_rows} rows, {num_cols} columns, caption: {caption[:30] if caption else 'None'}", "full")
    table = Table(num_rows, num_cols)
    node_id = story.add(table, caption)  # Pass caption here
    placeholder_ids[placeholder_key] = node_id
    parsed_blocks[node_id] = block
    log_message(f"Added table, node ID: {node_id}", "full")


def add_code_block(story, block, i, placeholder_key, placeholder_ids, parsed_blocks):
    """Add a code block to the StoryMap."""
    code_content = block.get('code', '')
    language = block.get('language', 'txt')
    
    # Create placeholder text - we'll replace this later
    placeholder_text = f"PLACEHOLDER_CODE_{i}"
    
    # Log what we're doing
    log_message(f"Adding code block with language {language} and placeholder", "full")
    log_message(f"Code content: {code_content[:50]}...", "full")
    
    # Initially add as a text node with a placeholder
    text = Text(placeholder_text, TextStyles.PARAGRAPH)
    node_id = story.add(text)
    
    placeholder_ids[placeholder_key] = node_id
    # Store the actual code content and language in parsed_blocks
    parsed_blocks[node_id] = {
        "type": "code", 
        "code": code_content,
        "language": language
    }
    log_message(f"Added code placeholder, node ID: {node_id}, to be transformed later", "full")
    return node_id

def update_storymap_json(storymap_item, placeholder_ids, parsed_blocks, image_dimensions=None):
    """Download StoryMap JSON, replace placeholders, and update both data and draft."""
    try:
        # Get StoryMap data
        main_data = get_storymap_data(storymap_item)
        if not main_data:
            return False
            
        # Update content in the data
        replacements = update_storymap_content(main_data, parsed_blocks)
        log_message(f"Made {replacements} content replacements in data", "basic")
        
        # Update image dimensions in resources
        if 'resources' in main_data and image_dimensions:
            images_updated = update_image_dimensions(main_data, parsed_blocks, image_dimensions)
            log_message(f"Updated dimensions for {images_updated} images in resources section", "basic")
        
        # Save updated data back to the StoryMap
        return save_storymap_updates(storymap_item, main_data)
        
    except Exception as e:
        log_message(f"Error updating StoryMap JSON: {str(e)}", "none", is_error=True)
        import traceback
        log_message(traceback.format_exc(), "basic", is_error=True)
        return False


def get_storymap_data(storymap_item):
    """Retrieve the StoryMap data, preferring the draft if available."""
    log_message("Getting StoryMap data...", "basic")
    
    # First try to get draft data
    resources = storymap_item.resources.list()
    draft_file_name = None
    
    for resource in resources:
        if isinstance(resource, dict) and 'resource' in resource:
            if resource['resource'].startswith('draft_'):
                draft_file_name = resource['resource']
                log_message(f"Found draft file: {draft_file_name}", "full")
                break
    
    main_data = None
    if draft_file_name:
        log_message(f"Retrieving draft file content", "basic")
        draft_content = storymap_item.resources.get(draft_file_name)
        
        if isinstance(draft_content, dict):
            main_data = draft_content
            log_message("Draft content is already a dictionary", "full")
        else:
            try:
                main_data = json.loads(draft_content)
                log_message("Parsed draft content as JSON", "full")
            except:
                log_message("Failed to parse draft content. Will try main data.", "basic", is_warning=True)
    
    # If draft data couldn't be obtained, fall back to main data
    if not main_data:
        log_message("Getting main item data", "basic")
        main_data = storymap_item.get_data()
        
    if not main_data or 'nodes' not in main_data:
        log_message("Unable to get valid StoryMap data structure. Cannot proceed.", "none", is_error=True)
        return None
        
    log_message(f"Retrieved StoryMap data with {len(main_data['nodes'])} nodes", "basic")
    return main_data
    

def update_storymap_content(main_data, parsed_blocks):
    """Update the StoryMap content with the actual content from parsed blocks."""
    # Language mapping for code blocks
    language_mapping = {
        'text': 'txt',
        'python': 'py',
        'javascript': 'js',
        'typescript': 'ts',
        'java': 'java',
        'csharp': 'cs',
        'html': 'html',
        'css': 'css',
        'ruby': 'rb',
        'php': 'php',
        'c': 'c',
        'cpp': 'cpp',
        'go': 'go',
        'sql': 'sql',
        'shell': 'sh',
        'xml': 'xml',
        'json': 'json',
        'yaml': 'yaml',
        'markdown': 'md',
        'r': 'r',
        'swift': 'swift',
        'kotlin': 'kt'
    }
    
    replacements = 0
    
    # Process each node
    log_message("Updating node content", "basic")
    for node_id, node in main_data['nodes'].items():
        log_message(f"Processing node: {node_id}, type: {node.get('type')}", "full")
        
        # Handle text nodes that need to be transformed to code nodes
        if node.get('type') == 'text' and node_id in parsed_blocks:
            block = parsed_blocks[node_id]
            
            # Check if this is a code block placeholder
            if block.get('type') == 'code':
                try:
                    log_message(f"Transforming text node {node_id} to code node", "full")
                    
                    # Get the code content and language
                    code_content = block.get('code', '')
                    language = block.get('language', 'text').lower()
                    mapped_language = language_mapping.get(language, 'txt')
                    
                    # Update the node type and data structure
                    node['type'] = 'code'
                    
                    # Create proper code node data structure
                    node['data'] = {
                        'content': code_content,
                        'lang': mapped_language,
                        'lineNumbers': True
                    }
                    
                    replacements += 1
                    log_message(f"Successfully transformed node {node_id} to code node with language {mapped_language}", "full")
                    continue
                except Exception as e:
                    log_message(f"Failed to transform text to code node {node_id}: {str(e)}", "basic", is_warning=True)
        
        # Process normal text nodes
        if node.get('type') == 'text' and 'data' in node and 'text' in node['data']:
            current_text = node['data']['text']
            if current_text.startswith("PLACEHOLDER_TEXT_"):
                # Try to find the matching content
                if node_id in parsed_blocks:
                    block = parsed_blocks[node_id]
                    if block.get('type') == 'text':
                        # Get the text content
                        text_content = block.get('text', '')
                        
                        # Skip completely empty text nodes
                        if not text_content or text_content.strip() == "" or text_content == "...":
                            log_message(f"Skipping empty text block for node {node_id}", "basic")
                            # Remove this node from the nodes collection
                            main_data['nodes'].pop(node_id, None)
                            log_message(f"Removed empty text node {node_id}", "basic")
                            continue
                        
                        # Debug information
                        log_message(f"For node {node_id}, replacing placeholder with text: '{text_content[:30]}...'", "full")
                        
                        # Ensure we have non-empty text to replace with
                        if not text_content or text_content.strip() == "":
                            # If the parsed text is empty, create a non-empty placeholder
                            text_content = " "  # Single space
                            log_message(f"Empty text content for node {node_id}, using space placeholder", "basic", is_warning=True)
                        
                        # Update the text content
                        node['data']['text'] = text_content
                        
                        # Update alignment if present
                        alignment = block.get('alignment')
                        if alignment:
                            alignment_map = {
                                'center': 'center',
                                'right': 'end',
                                'justify': 'justify',
                                'left': 'start'
                            }
                            if alignment in alignment_map:
                                node['data']['textAlignment'] = alignment_map[alignment]
                                log_message(f"Set text alignment to {alignment_map[alignment]}", "full")
                        
                        replacements += 1
                        log_message(f"Replaced placeholder in node {node_id}", "full")
        
        # Update table content
        elif node.get('type') == 'table' and 'data' in node:
            if node_id in parsed_blocks:
                block = parsed_blocks[node_id]
                if block.get('type') == 'table':
                    # Update table data with rows content
                    rows = block.get('rows', [])
                    num_rows = len(rows)
                    num_cols = max(len(row) for row in rows) if rows else 0
                    
                    log_message(f"Updating table node {node_id} with {num_rows}x{num_cols} cells", "full")
                    
                    # Update table data
                    node['data']['numRows'] = num_rows
                    node['data']['numColumns'] = num_cols
                    
                    # Create cells structure
                    if 'cells' not in node['data']:
                        node['data']['cells'] = {}
                    
                    # Populate cells
                    for row_idx, row in enumerate(rows):
                        if str(row_idx) not in node['data']['cells']:
                            node['data']['cells'][str(row_idx)] = {}
                        
                        for col_idx, cell_value in enumerate(row):
                            if col_idx < num_cols:  # Ensure we don't exceed defined columns
                                node['data']['cells'][str(row_idx)][str(col_idx)] = {
                                    "value": cell_value
                                }
                    
                    # Add caption if present
                    if 'caption' in block and block['caption']:
                        node['data']['caption'] = block['caption']
                        log_message(f"Added caption to table: {block['caption'][:30]}...", "full")
                    
                    replacements += 1
                    log_message(f"Updated table in node {node_id}", "full")
                    
        # Update image properties (caption and float alignment)
        elif node.get('type') == 'image' and 'data' in node:
            if node_id in parsed_blocks:
                block = parsed_blocks[node_id]
                
                # Check if there's a caption
                if 'caption' in block and block['caption']:
                    node['data']['caption'] = block['caption']
                    log_message(f"Added caption to image: {block['caption'][:30]}...", "full")
                    replacements += 1
                
                # Apply float alignment if specified
                if block.get('display') == 'float' and block.get('float_alignment'):
                    if 'config' not in node:
                        node['config'] = {}
                    node['config']['size'] = 'float'
                    node['config']['floatAlignment'] = block['float_alignment']
                    log_message(f"Applied float alignment '{block['float_alignment']}' to image node {node_id}", "full")
                    replacements += 1
    
    return replacements


def update_image_dimensions(main_data, parsed_blocks, image_dimensions):
    """Update image dimensions in the resources section."""
    log_message("Updating image dimensions in resources section", "basic")
    images_updated = 0
    
    for resource_id, resource in main_data['resources'].items():
        if resource.get('type') == 'image' and 'data' in resource:
            # Find the matching image in resource data by filename
            if 'resourceId' in resource['data']:
                filename = resource['data'].get('resourceId', '')
                log_message(f"Processing resource: {resource_id}, resourceId: {filename}", "full")
                
                # Look for image nodes with this resource
                for node_id, node in main_data['nodes'].items():
                    if node.get('type') == 'image' and 'data' in node and 'image' in node['data']:
                        image_ref = node['data']['image']
                        
                        # If this node references this resource
                        if image_ref == resource_id:
                            # Find corresponding image in parsed_blocks
                            if node_id in parsed_blocks and 'dimensions' in parsed_blocks[node_id]:
                                width, height = parsed_blocks[node_id]['dimensions']
                                resource['data']['width'] = width
                                resource['data']['height'] = height
                                images_updated += 1
                                log_message(f"Updated resource {resource_id} with dimensions {width}x{height}", "full")
                                break
    
    return images_updated


def save_storymap_updates(storymap_item, main_data):
    """Save the updated data back to the StoryMap."""
    # 1. Update the main item data
    log_message("Updating main StoryMap data...", "basic")
    storymap_item.update(data=main_data)
    log_message("Main data updated successfully", "basic")
    
    # 2. Update the draft file
    log_message("Retrieving draft file name...", "full")
    resources = storymap_item.resources.list()
    draft_file_name = None
    
    for resource in resources:
        if isinstance(resource, dict) and 'resource' in resource:
            if resource['resource'].startswith('draft_'):
                draft_file_name = resource['resource']
                break
    
    if not draft_file_name:
        log_message("Draft file not found. Only main data was updated.", "basic", is_warning=True)
        return True
    
    log_message(f"Found draft file: {draft_file_name}", "full")
    
    # Write the updated data to a temporary file
    import requests
    import tempfile
    
    draft_file_path = os.path.join(tempfile.gettempdir(), draft_file_name)
    
    with open(draft_file_path, "w", encoding="utf-8") as draft_file:
        json.dump(main_data, draft_file, ensure_ascii=False, indent=2)
    
    # Upload using direct REST API call
    log_message("Uploading modified draft file...", "basic")
    update_url = f"{storymap_item._gis._portal.resturl}content/users/{storymap_item.owner}/items/{storymap_item.id}/updateResources"
    
    params = {
        "f": "json",
        "resource": draft_file_name,
        "access": "inherit",
        "token": storymap_item._gis._con.token,
    }
    
    with open(draft_file_path, "rb") as draft_file:
        files = {"file": draft_file}
        response = requests.post(update_url, data=params, files=files)
    
    if response.status_code == 200 and response.json().get("success"):
        log_message("Modified draft file uploaded successfully", "basic")
        return True
    else:
        log_message(f"Failed to upload modified draft file. Response: {response.text}", "basic", is_error=True)
        return False

        
def create_text_block(text_type, text_content, alignment=None):
    """Create a text block with the specified type, content, and alignment."""
    # Make sure text content is not None
    if text_content is None:
        text_content = ""
        
    # Create the base block structure
    block = {
        'type': 'text',
        'text_type': text_type,
        'text': text_content,
        'alignment': alignment
    }
    
    # Handle alignment if provided
    if alignment:
        block['data'] = {
            'textAlignment': alignment
        }
    
    log_message(f"Created text block of type {text_type}", "full")
    return block

def create_image_block(image_path, caption=None, display=None, float_alignment=None):
    """Create an image block with the specified properties."""
    block = {
        'type': 'image',
        'path': image_path,
        'caption': caption,
        'display': display,
        'float_alignment': float_alignment
    }
    log_message(f"Created image block from {os.path.basename(image_path)} with display:{display}, alignment:{float_alignment}", "full")
    return block


def create_table_block(rows, caption=None):
    """Create a table block from the provided rows."""
    block = {
        'type': 'table',
        'rows': rows,
        'caption': caption
    }
    log_message(f"Created table block with {len(rows)} rows", "full")
    return block

def create_code_block(code_content, language=None):
    """
    Create a code block with the specified content and language.
    If language is None, attempts to auto-detect the language.
    """
    if language is None or language == 'text':
        language = detect_code_language(code_content)
        log_message(f"Auto-detected code language: {language}", "full")
    
    block = {
        'type': 'code',
        'code': code_content,
        'language': language
    }
    log_message(f"Created code block with language {language}", "full")
    return block


def create_separator_block():
    """Create a separator block."""
    block = {
        'type': 'separator'
    }
    log_message("Created separator block", "full")
    return block

def sanitize_html(html_text, allowed_tags=None):
    """Sanitize HTML to only include allowed tags."""
    if allowed_tags is None:
        # Expand allowed tags to include more formatting options
        allowed_tags = ['strong', 'em', 'a', 'span', 'u', 's', 'sub', 'sup', 'ul', 'ol', 'li']
    
    # Basic sanitization - preserve all allowed tags
    import re
    
    # First, replace all occurrences of <br> or <br/> with newlines
    html_text = re.sub(r'<br\s*/?>', '\n', html_text)
    
    # For each disallowed tag, replace the opening and closing tags with empty strings
    for tag in re.findall(r'</?([a-z][a-z0-9]*)', html_text, re.IGNORECASE):
        tag = tag.lower()
        if tag not in allowed_tags:
            html_text = re.sub(f'</?{tag}[^>]*>', '', html_text)
    
    log_message(f"Sanitized HTML, allowed tags: {', '.join(allowed_tags)}", "full")
    return html_text
    



def detect_code_language(code_content):
    """
    Detect the programming language of a code snippet using pattern recognition.
    
    Args:
        code_content (str): The code snippet to analyze
        
    Returns:
        str: The detected language code (compatible with StoryMap API)
    """
    import re
    import json
    
    # Debug header
    log_message(f"Detecting code language for snippet: '{code_content[:50]}...'", "full")
    
    # Default to plain text if we can't identify the language
    if not code_content or len(code_content.strip()) < 3:
        log_message("Code too short - returning txt", "full")
        return 'txt'
    
    # Prepare the code for analysis
    code = code_content.strip()
    code_lower = code.lower()
    
    # Check each language in priority order
    lang = check_sql(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (SQL check)", "full")
        return lang
    
    lang = check_arcade(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (Arcade check)", "full")
        return lang
    
    lang = check_json(code)
    if lang: 
        log_message(f"Detected language: {lang} (JSON check)", "full")
        return lang
    
    lang = check_python(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (Python check)", "full")
        return lang
    
    lang = check_csharp(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (C# check)", "full")
        return lang
    
    lang = check_javascript_typescript(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (JS/TS check)", "full")
        return lang
    
    lang = check_css(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (CSS check)", "full")
        return lang
    
    lang = check_html(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (HTML check)", "full")
        return lang
    
    # Fallback to keyword matching
    lang = check_keywords(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (keyword check)", "full")
        return lang
    
    # Special case: text about code
    lang = check_text_about_code(code, code_lower)
    if lang: 
        log_message(f"Detected language: {lang} (text about code check)", "full")
        return lang
    
    # Fallback to plain text
    log_message("No confident match found, using text as fallback", "full")
    return 'txt'

def check_sql(code, code_lower):
    """Check if the code is SQL."""
    sql_patterns = [
        (r'SELECT\s+.+?\s+FROM', "SELECT FROM"),
        (r'INSERT\s+INTO', "INSERT INTO"),
        (r'UPDATE\s+.+?\s+SET', "UPDATE SET"),
        (r'CREATE\s+TABLE', "CREATE TABLE"),
        (r'ALTER\s+TABLE', "ALTER TABLE"),
        (r'DROP\s+TABLE', "DROP TABLE"),
        (r'JOIN\s+\w+\s+ON', "JOIN ON"),
        (r'WHERE\s+\w+\s*[=<>]', "WHERE clause"),
        (r'ORDER\s+BY\s+\w+', "ORDER BY"),
        (r'GROUP\s+BY\s+\w+', "GROUP BY")
    ]
    
    matches = 0
    for pattern, description in sql_patterns:
        if re.search(pattern, code_lower, re.IGNORECASE | re.MULTILINE):
            matches += 1
            log_message(f"SQL match: {description}", "full")
    
    if matches >= 1:
        return 'sql'
    return None

def check_arcade(code, code_lower):
    """Check if the code is Arcade."""
    arcade_patterns = [
        (r'Geometry\(', "Arcade Geometry constructor"),
        (r'(Feature|FeatureSet)\(', "Arcade Feature constructor"),
        (r'When\(', "Arcade When function"),
        (r'(Text|Count|Concatenate|IIf|IsEmpty)\(', "Arcade functions"),
        (r'\$feature', "Arcade feature reference"),
        (r'\$map', "Arcade map reference"),
        (r'//.*$', "Arcade comment")
    ]
    
    matches = 0
    for pattern, description in arcade_patterns:
        if re.search(pattern, code, re.IGNORECASE | re.MULTILINE):
            matches += 1
            log_message(f"Arcade match: {description}", "full")
    
    if matches >= 1:
        return 'arcade'
    return None

def check_json(code):
    """Check if the code is valid JSON."""
    try:
        if (code.strip().startswith('{') and code.strip().endswith('}')) or \
           (code.strip().startswith('[') and code.strip().endswith(']')):
            json.loads(code)
            log_message("Valid JSON structure detected", "full")
            return 'json'
    except:
        pass
    return None

def check_python(code, code_lower):
    """Check if the code is Python."""
    python_patterns = [
        (r'\bdef\s+\w+\s*\(', "Function definition"),
        (r'\bclass\s+\w+\s*:', "Class definition"),
        (r'import\s+[\w.]+', "Import statement"),
        (r'from\s+[\w.]+\s+import', "From import"),
        (r'if\s+__name__\s*==\s*[\'"]__main__[\'"]', "Main block"),
        (r'arcpy\.\w+', "ArcPy call"),
        (r'#.*?$', "Python comment")
    ]
    
    matches = 0
    for pattern, description in python_patterns:
        if re.search(pattern, code, re.IGNORECASE | re.MULTILINE):
            matches += 1
            log_message(f"Python match: {description}", "full")
    
    if matches >= 1:
        return 'py'
    return None

def check_csharp(code, code_lower):
    """Check if the code is C#."""
    csharp_patterns = [
        (r'using\s+System;', "Using System"),
        (r'namespace\s+\w+', "Namespace declaration"),
        (r'(public|private|protected)\s+(class|interface)', "Class/interface declaration"),
        (r'(public|private|protected)\s+\w+\s+\w+\s*\(', "Method declaration"),
        (r'Console\.(Write|WriteLine)', "Console output")
    ]
    
    matches = 0
    for pattern, description in csharp_patterns:
        if re.search(pattern, code, re.IGNORECASE | re.MULTILINE):
            matches += 1
            log_message(f"C# match: {description}", "full")
    
    if matches >= 1:
        return 'cs'
    return None

def check_javascript_typescript(code, code_lower):
    """Check if the code is JavaScript or TypeScript."""
    js_patterns = [
        (r'function\s+\w+\s*\(', "Function declaration"),
        (r'(const|let|var)\s+\w+\s*=', "Variable declaration"),
        (r'=>', "Arrow function"),
        (r'console\.(log|warn|error)', "Console statement"),
        (r'document\.get(Element|ElementsByTagName)', "DOM manipulation"),
        (r'window\.\w+', "Window object"),
        (r'new\s+\w+\(', "Constructor call")
    ]
    
    js_matches = 0
    for pattern, description in js_patterns:
        if re.search(pattern, code, re.IGNORECASE | re.MULTILINE):
            js_matches += 1
            log_message(f"JavaScript match: {description}", "full")
    
    # Check TypeScript
    ts_patterns = [
        (r'interface\s+\w+', "Interface"),
        (r'type\s+\w+\s*=', "Type definition"),
        (r':\s*\w+[\[\]<>]*(\s*=|\))', "Type annotation"),
        (r'<\w+>[\(\[]', "Generic"),
        (r'as\s+\w+', "Type assertion")
    ]
    
    ts_matches = 0
    for pattern, description in ts_patterns:
        if re.search(pattern, code, re.IGNORECASE | re.MULTILINE):
            ts_matches += 1
            log_message(f"TypeScript match: {description}", "full")
    
    # Determine if it's JS, TS, JSX, or TSX
    if js_matches >= 1 or ts_matches >= 1:
        # Check if it contains HTML-like syntax
        has_html = bool(re.search(r'<\w+(\s+\w+=["\'].+?["\'])*>', code))
        
        if has_html:
            if ts_matches >= 1:
                return 'tsx'
            else:
                return 'jsx'
        elif ts_matches >= 1:
            return 'ts'
        else:
            return 'js'
    return None

def check_css(code, code_lower):
    """Check if the code is CSS."""
    if '{' in code and ('}' in code or ';' in code):
        css_patterns = [
            (r'[\w-]+\s*:\s*[^;]+;', "Property:value"),
            (r'\.\w+[\w-]*\s*\{', "Class selector"),
            (r'#\w+[\w-]*\s*\{', "ID selector"),
            (r'@(media|keyframes|import|font-face)', "CSS at-rule"),
            (r'(margin|padding|color|background|font|display):', "Common CSS property")
        ]
        
        matches = 0
        for pattern, description in css_patterns:
            if re.search(pattern, code, re.IGNORECASE):
                matches += 1
                log_message(f"CSS match: {description}", "full")
        
        if matches >= 2 or (matches >= 1 and len(code) < 100):
            return 'css'
    return None

def check_html(code, code_lower):
    """Check if the code is HTML."""
    if '<' in code and '>' in code:
        # Count HTML tags (both opening and closing)
        tag_pattern = r'</?[a-z][a-z0-9]*\b[^>]*>'
        tags = re.findall(tag_pattern, code, re.IGNORECASE)
        
        # Count angle brackets
        open_brackets = code.count('<')
        close_brackets = code.count('>')
        
        # Common HTML attribute pattern
        attr_pattern = r'\s+(href|src|alt|class|id|style)=["\'"][^\'"]*["\']'
        attrs = re.findall(attr_pattern, code, re.IGNORECASE)
        
        log_message(f"HTML tags found: {len(tags)}, attributes: {len(attrs)}", "full")
        
        # If we have at least 2 tags or 1 tag with attributes, it's likely HTML
        if len(tags) >= 2 or (len(tags) >= 1 and len(attrs) >= 1):
            # Check for balanced brackets as additional confidence
            if abs(open_brackets - close_brackets) <= 2:
                log_message(f"HTML detected with {len(tags)} tags and {len(attrs)} attributes", "full")
                return 'html'
                
        # Specific HTML pattern checks for additional confidence
        html_patterns = [
            (r'<!DOCTYPE\s+html', "HTML doctype"),
            (r'<html>|<html\s+', "HTML root tag"),
            (r'<(div|span|p|a|img|h[1-6])(\s+[^>]*)?>', "Common HTML tag")
        ]
        
        for pattern, description in html_patterns:
            if re.search(pattern, code, re.IGNORECASE):
                log_message(f"Strong HTML pattern match: {description}", "full")
                return 'html'
    return None

def check_keywords(code, code_lower):
    """Check for language-specific keywords."""
    keyword_sets = {
        'py': ['def', 'import', 'class', 'self', 'None', 'True', 'False', 'if', 'elif', 'else', 'for', 'in', 'try', 'except'],
        'js': ['function', 'const', 'let', 'var', 'return', 'true', 'false', 'null', 'undefined', 'this', 'new'],
        'sql': ['select', 'from', 'where', 'insert', 'update', 'delete', 'create', 'drop', 'alter', 'join'],
        'cs': ['using', 'namespace', 'public', 'private', 'class', 'void', 'string', 'int', 'bool'],
        'html': ['div', 'span', 'class', 'id', 'style', 'href', 'src'],
        'css': ['margin', 'padding', 'color', 'background', 'width', 'height', 'font'],
        'arcade': ['when', 'feature', 'geometry', 'text', 'count', 'iif']
    }
    
    # Count occurrences of keywords for each language
    language_scores = {}
    words = re.findall(r'\b(\w+)\b', code_lower)
    
    for lang, keywords in keyword_sets.items():
        matches = sum(1 for word in words if word in keywords)
        if matches > 0:
            language_scores[lang] = matches
            log_message(f"Keyword matches for {lang}: {matches}", "full")
    
    if language_scores:
        best_match = max(language_scores.items(), key=lambda x: x[1])
        lang, score = best_match
        
        # Only use keyword detection if we have enough matches
        if score >= 2:
            log_message(f"Best language match based on keywords: {lang} with {score} matches", "full")
            return lang
    return None

def check_text_about_code(code, code_lower):
    """Check if this is text about code rather than code itself."""
    if ('code' in code_lower and ('style' in code_lower or 'add' in code_lower)) or \
       ('syntax' in code_lower or 'example' in code_lower):
        log_message("Text appears to be about code, not code itself", "full")
        
        # Try to infer language from context
        if 'python' in code_lower:
            return 'py'
        elif 'javascript' in code_lower:
            return 'js'
        elif 'html' in code_lower:
            return 'html'
        elif 'css' in code_lower:
            return 'css'
        elif 'sql' in code_lower:
            return 'sql'
        elif 'c#' in code_lower or 'csharp' in code_lower:
            return 'cs'
    return None

def format_link_for_storymap(url, link_text):
    """
    Format links consistently for StoryMap.
    Ensures links open in a new tab and use the actual URL from Word.
    
    Args:
        url: The URL for the link
        link_text: The text to display for the link
        
    Returns:
        Properly formatted HTML link tag
    """
    # Make sure URL is properly formed
    if url and not url.startswith(('http://', 'https://', '#')):
        url = 'https://' + url
        
    # Create a properly formatted link that opens in a new tab
    return f'<a href="{url}" rel="noopener noreferrer" target="_blank">{link_text}</a>'
    

    
if __name__ == "__main__":
    main()