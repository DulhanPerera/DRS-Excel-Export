import configparser
import logging
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Color

logger = logging.getLogger('excel_data_writer')

STYLES = {}

def load_styles_from_ini(ini_file_path):
    """
    Load styles from the INI file and store them in the global STYLES dictionary.
    """
    global STYLES
    config = configparser.ConfigParser()
    try:
        config.read(ini_file_path)
        if not config.sections():
            logger.error(f"INI file '{ini_file_path}' is empty or not found.")
            sys.exit(1)

        for section in config.sections():
            STYLES[section] = {}
            for key, value in config[section].items():
                if key == 'font':
                    font_attrs = {}
                    for attr in value.split(','):
                        attr_name, attr_value = attr.strip().split('=')
                        if attr_name == "size":
                            attr_value = int(attr_value)
                        elif attr_name in ["bold", "italic"]:
                            attr_value = attr_value.lower() == "true"
                        elif attr_name == "color":
                            if not attr_value.startswith("FF"):
                                attr_value = "FF" + attr_value  # Ensure alpha channel
                            attr_value = Color(rgb=attr_value)
                        if attr_name in ["name", "size", "bold", "italic", "color"]:
                            font_attrs[attr_name] = attr_value
                    STYLES[section]['font'] = Font(**font_attrs)
                elif key == 'fill':
                    fill_attrs = {}
                    for attr in value.split(','):
                        attr_name, attr_value = attr.strip().split('=')
                        if attr_name in ["start_color", "end_color"]:
                            if not attr_value.startswith("FF"):
                                attr_value = "FF" + attr_value  # Ensure alpha channel
                            attr_value = Color(rgb=attr_value)
                        fill_attrs[attr_name] = attr_value
                    STYLES[section]['fill'] = PatternFill(**fill_attrs)
                elif key == 'alignment':
                    alignment_attrs = {}
                    for attr in value.split(','):
                        attr_name, attr_value = attr.strip().split('=')
                        alignment_attrs[attr_name] = attr_value
                    STYLES[section]['alignment'] = Alignment(**alignment_attrs)
                elif key == 'border':
                    border_attrs = {}
                    for attr in value.split(','):
                        attr_name, attr_value = attr.strip().split('=')
                        border_attrs[attr_name] = Side(style=attr_value)
                    STYLES[section]['border'] = Border(**border_attrs)
                elif key == 'width':
                    STYLES[section]['width'] = float(value)
    except Exception as e:
        logger.error(f"Failed to load styles from INI file: {e}")
        sys.exit(1)

# Load styles from the INI file
load_styles_from_ini('config/table_format.ini')