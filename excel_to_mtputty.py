#!/usr/bin/env python3
"""
Excel to MTPuTTY XML Converter - V3 with optional Script block support
Converts an Excel file with folder columns to MTPuTTY-compatible XML format

Script columns: Script_L0, Script_L1, Script_L2, ... (optional)
  These map to <Script><L0>...</L0><L1>...</L1>...</Script> in the XML.
"""

import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import sys
from pathlib import Path


def create_element_with_text(parent, tag, text):
    """Create an element with text content."""
    elem = ET.SubElement(parent, tag)
    elem.text = str(text)
    return elem


def create_folder_structure(root, path_parts):
    """
    Create nested folder structure in XML tree.
    Returns the deepest folder node.
    """
    current = root
    for part in path_parts:
        if part and str(part).strip() and str(part).lower() != 'nan':
            part_str = str(part).strip()

            # Check if folder already exists
            existing = None
            for child in current:
                if child.tag == 'Node' and child.get('Type') == '0':
                    display_name = child.find('DisplayName')
                    if display_name is not None and display_name.text == part_str:
                        existing = child
                        break

            if existing is not None:
                current = existing
            else:
                # Create new folder node
                folder = ET.SubElement(current, 'Node')
                folder.set('Type', '0')
                create_element_with_text(folder, 'DisplayName', part_str)
                current = folder

    return current


def excel_to_mtputty_xml(excel_file, output_file=None):
    """
    Convert Excel file to MTPuTTY XML format.

    Expected Excel columns:
    - Folder1, Folder2, Folder3, ... (as many as needed)
    - Server Name (required)
    - Host (required)
    - Port (optional, default 22)
    - Username (optional)
    - Password (optional)
    - Protocol (optional, default SSH)
    - Script_L0, Script_L1, Script_L2, ... (optional script lines)
      These become <Script><L0>...</L0><L1>...</L1>...</Script> in the XML.
    """

    # Read Excel file
    print(f"Reading Excel file: {excel_file}")
    df = pd.read_excel(excel_file)

    # Display available columns
    print(f"\nFound columns: {', '.join(df.columns.tolist())}")

    # Identify folder columns
    folder_columns = [col for col in df.columns if
                      'folder' in str(col).lower() or
                      str(col).lower().startswith('folder')]
    print(f"Identified folder columns: {', '.join(folder_columns)}")

    # Identify script line columns (Script_L0, Script_L1, ...)
    script_columns = sorted(
        [col for col in df.columns if str(col).lower().startswith('script_l')],
        key=lambda c: int(c.split('_L')[-1]) if c.split('_L')[-1].isdigit() else 0
    )
    if script_columns:
        print(f"Identified script columns: {', '.join(script_columns)}")
    else:
        print("No script columns found (Script_L0, Script_L1, ...) — Script block will be omitted.")

    # Create root XML structure
    servers = ET.Element('Servers')
    putty = ET.SubElement(servers, 'Putty')
    root_node = ET.SubElement(putty, 'Node')
    root_node.set('Type', '0')

    # Process each row
    print(f"\nProcessing {len(df)} rows...")

    for idx, row in df.iterrows():
        # Extract folder path
        folder_path = []
        for folder_col in folder_columns:
            if folder_col in row and pd.notna(row[folder_col]):
                folder_path.append(str(row[folder_col]).strip())

        # Create folder structure and get the parent folder
        parent_folder = create_folder_structure(root_node, folder_path)

        # Get server name
        server_name = None
        if 'Server Name' in row and pd.notna(row['Server Name']):
            server_name = str(row['Server Name']).strip()
        elif 'ServerName' in row and pd.notna(row['ServerName']):
            server_name = str(row['ServerName']).strip()
        elif 'Name' in row and pd.notna(row['Name']):
            server_name = str(row['Name']).strip()

        if not server_name:
            print(f"Warning: Row {idx + 2} missing server name, skipping...")
            continue

        # Get host
        host = None
        if 'Host' in row and pd.notna(row['Host']):
            host = str(row['Host']).strip()
        elif 'IP' in row and pd.notna(row['IP']):
            host = str(row['IP']).strip()
        elif 'IP Address' in row and pd.notna(row['IP Address']):
            host = str(row['IP Address']).strip()

        if not host:
            print(f"Warning: Row {idx + 2} missing host/IP, skipping...")
            continue

        # Get optional fields
        port = '22'
        if 'Port' in row and pd.notna(row['Port']):
            port = str(int(row['Port']))

        username = ''
        if 'Username' in row and pd.notna(row['Username']):
            username = str(row['Username']).strip()

        password = ''
        if 'Password' in row and pd.notna(row['Password']):
            password = str(row['Password']).strip()

        protocol = '0'  # Default SSH
        if 'Protocol' in row and pd.notna(row['Protocol']):
            protocol_val = str(row['Protocol']).strip().upper()
            protocol_map = {'SSH': '0', 'TELNET': '1', 'RLOGIN': '2', 'RAW': '3', 'SERIAL': '4'}
            protocol = protocol_map.get(protocol_val, '0')

        # Build CLParams
        cl_params = f"{host} -ssh -l {username}" if username else f"{host} -ssh -l root"

        # Collect script lines from Script_L0, Script_L1, ... columns
        script_lines = []
        for col in script_columns:
            if col in row and pd.notna(row[col]) and str(row[col]).strip():
                script_lines.append(str(row[col]).strip())

        # Create server node
        server = ET.SubElement(parent_folder, 'Node')
        server.set('Type', '1')

        # Add child elements
        create_element_with_text(server, 'SavedSession', 'Default Settings')
        create_element_with_text(server, 'DisplayName', server_name)
        create_element_with_text(server, 'ServerName', host)
        create_element_with_text(server, 'PuttyConType', protocol)
        create_element_with_text(server, 'Port', port)
        create_element_with_text(server, 'UserName', username)
        create_element_with_text(server, 'Password', password)
        create_element_with_text(server, 'PasswordDelay', '500')
        create_element_with_text(server, 'CLParams', cl_params)
        create_element_with_text(server, 'ScriptDelay', '1000')

        # Optional <Script> block — only added if at least one script line exists
        if script_lines:
            script_elem = ET.SubElement(server, 'Script')
            for line_idx, line_text in enumerate(script_lines):
                create_element_with_text(script_elem, f'L{line_idx}', line_text)

    # Pretty print XML
    xml_str = minidom.parseString(ET.tostring(servers)).toprettyxml(indent="  ")

    # Remove extra blank lines
    xml_str = '\n'.join([line for line in xml_str.split('\n') if line.strip()])

    # Determine output filename
    if output_file is None:
        input_path = Path(excel_file)
        output_file = input_path.stem + '_mtputty.xml'

    # Write to file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(xml_str)

    print(f"\n✓ Successfully created MTPuTTY XML: {output_file}")
    print(f"\nTo import into MTPuTTY:")
    print(f"1. Open MTPuTTY")
    print(f"2. Go to Servers > Import servers")
    print(f"3. Select the file: {output_file}")

    return output_file


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python excel_to_mtputty.py <excel_file> [output_file]")
        print("\nExample:")
        print("  python excel_to_mtputty.py servers.xlsx")
        print("  python excel_to_mtputty.py servers.xlsx my_servers.xml")
        print("\nOptional Script columns in Excel: Script_L0, Script_L1, Script_L2, ...")
        sys.exit(1)

    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        excel_to_mtputty_xml(excel_file, output_file)
    except FileNotFoundError:
        print(f"Error: File '{excel_file}' not found!")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
