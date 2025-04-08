import logging
import os
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Tuple, Union

import comtypes.client
import comtypes.gen.SAP2000v1 as SAP2000

from .area_methods import AreaMethods
from .frame_methods import FrameMethods
from .group_methods import GroupMethods

logger = logging.getLogger(__name__)

class CustomSAP2000Model(AreaMethods, FrameMethods, GroupMethods):
    """
    Custom SAP2000 model class that extends the base SAP2000 model with additional functionality.
    """
    def __init__(self, sap_model):
        self._model = sap_model
        self.section_names = self._load_section_names()

    def _load_section_names(self) -> Dict[str, List[str]]:
        """
        Loads and parses the AISC16.xml file to get all section names
        """
        xml_path = os.path.join(os.path.dirname(__file__), 'data', 'AISC16.xml')
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Handle XML namespace
        # Extract namespace from root tag
        ns = {'ns': root.tag.split('}')[0].strip('{')} if '}' in root.tag else ''
        
        # Dictionary to store section names by type
        section_names = {
            'W': [],  # Wide flange sections (from STEEL_I_SECTION)
            'L': [],  # Angle sections
            'C': [],  # Channel sections
            'WT': [], # Tee sections
            'HSS': [] # Hollow structural sections
        }
        
        # Parse each section type
        section_mappings = {
            'STEEL_I_SECTION': ['W'],  # W sections from I sections
            'STEEL_ANGLE': ['L'],  # Angle sections
            'STEEL_CHANNEL': ['C'],  # Channel sections
            'STEEL_TEE': ['WT'],  # Tee sections
            'STEEL_BOX': ['HSS'],  # Box sections
            'STEEL_PIPE': ['HSS']  # Pipe sections (also HSS)
        }
        
        for xml_type, designations in section_mappings.items():
            # Find all sections of this type, handling potential namespace
            xpath = f".//{xml_type}" if not ns else f".//ns:{xml_type}"
            for section in root.findall(xpath, namespaces=ns):
                # Find elements, handling potential namespace
                label_elem = section.find('LABEL' if not ns else 'ns:LABEL', namespaces=ns)
                designation_elem = section.find('DESIGNATION' if not ns else 'ns:DESIGNATION', namespaces=ns)
                
                if label_elem is not None and designation_elem is not None:
                    label = label_elem.text
                    designation = designation_elem.text
                    # Store in appropriate category if designation matches
                    if designation in designations:
                        section_names[designation].append(label)
        
        # Log the number of sections found for each type
        for section_type, sections in section_names.items():
            logger.info(f"Found {len(sections)} {section_type} sections")
            if sections:  # Log a sample of sections found
                logger.info(f"Sample {section_type} sections: {', '.join(sections[:5])}")
            
        return section_names

    def __getattr__(self, name):
        """
        Forward any unknown attribute access to the underlying model.
        """
        return getattr(self._model, name) 