import logging
import os
import traceback
from typing import Optional, Tuple, List, Dict, Union

import comtypes.client
import comtypes.gen.SAP2000v1 as SAP2000

from sap2000_model import CustomSAP2000Model

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SAPTest:
    def __init__(self):
        self.sap_object = None
        self.sap_model = None
        self.model_path = None
        self._connected = False
        self._try_connect()
    
    def _try_connect(self) -> bool:
        try:
            import comtypes.gen.SAP2000v1
            logger.info("Attempting to connect to SAP2000...")
            helper = comtypes.client.CreateObject('SAP2000v1.Helper')
            helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
            self.sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
            # Wrap the SAP model in our custom class
            self.sap_model = CustomSAP2000Model(self.sap_object.SapModel)
            APIPath = os.path.join(os.path.dirname(__file__), 'model')
            os.makedirs(APIPath, exist_ok=True)
            self.model_path = os.path.join(APIPath, 'custom_grid_model.sdb')
            info = self.sap_model.GetProgramInfo()
            logger.info(f"Successfully connected to SAP2000 (Version: {info[0]}, Build: {info[1]})")
            self._connected = True
            return True
        except Exception as e:
            logger.error(f"Failed to connect to SAP2000: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            self._connected = False
            return False

    def run_sap_test_script(self) -> bool:
        # to just run segment of scripts ... for adhoc testing
        
        self.sap_model.InitializeNewModel(4)
        self.sap_model.File.NewBlank()
        return True

    def run_sap_script(self) -> bool:
        # Pre-requisitite (you can assume its done):
        # 1. A model with defined frames and joints are already loaded into sap and connected to the script and available through self.sap_model!
        # Step 1: Add base restraints to all ground level columns.
        # This code identifies the ground level columns and restrains them with no translation, but free to rotate.
        restraints = [True, True, True, False, False, False]
        restrained_joints, restraint_status = self.sap_model.add_base_restraints(restraints)
        # Step 2: Create floor areas and add dead and live loads to them.
        # substep: add dead and live load patterns definitions  
        self.sap_model.LoadPatterns.Add("DEAD", 1, 1.0)  # 1 is eLoadPatternType_Dead
        self.sap_model.LoadPatterns.Add("LIVE", 3, 0.0)  # 3 is eLoadPatternType_Live

        # substep: identify all the floor levels.
        floor_levels, floor_status = self.sap_model.identify_floor_levels()
        for i, floor_level in enumerate(floor_levels):
            # Check if this is the roof level since it needs a different load value
            is_roof = (i == len(floor_levels) - 1)
            # substep: create floor areas at each floor level.
            areas, area_status = self.sap_model.add_floor_areas(floor_level)
            # substep: add dead and live loads to the floor areas.
            for area_name in areas:
                self.sap_model.AreaObj.SetLoadUniform(
                    area_name,    # Area name
                    "DEAD",       # Load pattern
                    75.0,         # Load value (psf)
                    6,            # Direction (6 = Global Z)
                    True,         # Replace existing load
                    "Global"      # Coordinate system
                )
                live_load_value = 20.0 if is_roof else 50.0
                self.sap_model.AreaObj.SetLoadUniform(
                    area_name,    # Area name
                    "LIVE",       # Load pattern
                    live_load_value,  # Load value (psf)
                    6,            # Direction (6 = Global Z)
                    True,         # Replace existing load
                    "Global"      # Coordinate system
                )
        
        # Step3: Create Beam section groups and assign sections to them.
        # substep: get beam information by length since we group beams based on the length
        beams_by_length = self.sap_model.get_beams_info()
        print(f"beams by length: {beams_by_length}")
        
        # Define auto-select lists for beams based on length
        beam_auto_select_lists = {}
        
        # Get section candidates for each beam length group
        for length, frames in beams_by_length.items():
            group_name = f"{int(length)}ft Beams"
            sections = self.sap_model.define_section_candidate(frames, section_type='w', member_type='beam')
            beam_auto_select_lists[group_name] = sections[:5]  # Take top 5 sections for each group
    
        # Step 1: Import all section properties first
        for group_name, section_list in beam_auto_select_lists.items():
            for section in section_list:
                ret = self.sap_model.PropFrame.ImportProp(
                    section,
                    "A992Fy50",
                    "AISC16.xml",
                    section
                )
                if ret != 0:
                    logger.warning(f"Failed to import section {section}")

        # Step 2 & 3: Create auto-select lists and assign to beam groups
        for length, frames in beams_by_length.items():
            group_name = f"{int(length)}ft Beams"
            if group_name in beam_auto_select_lists:
                # Create group and assign frames
                self.sap_model.create_assign_section_group(
                    group_name=group_name,
                    frames=frames
                )
                
                # Create the auto-select list
                section_list = beam_auto_select_lists[group_name]
                auto_list_name = f"AUTO_{group_name}"
                ret = self.sap_model.PropFrame.SetAutoSelectSteel(
                    auto_list_name,
                    len(section_list),
                    section_list,
                    section_list[0]  # Start with smallest section
                )
                
                # Assign the auto-select list to the group
                ret = self.sap_model.FrameObj.SetSection(group_name, auto_list_name, 1)  # 1 = apply to group
                logger.info(f"Assigned auto-select list {auto_list_name} to {group_name}")

        # Step4: Create Column section groups and assign sections to them.
        # substep: get column information by location since we group columns based on the location
        columns_by_location = self.sap_model.get_columns_info()
        print(f"columns by location: {columns_by_location}")

        # Define auto-select lists for columns based on location
        column_auto_select_lists = {}
        
        # Get section candidates for each column location group
        for location, frames in columns_by_location.items():
            group_name = f"{location.capitalize()} Columns"
            sections = self.sap_model.define_section_candidate(frames, section_type='w', member_type='column')
            column_auto_select_lists[group_name] = sections[:5]  # Take top 5 sections for each group
        
        # Step 1: Import all column section properties
        for group_name, section_list in column_auto_select_lists.items():
            for section in section_list:
                ret = self.sap_model.PropFrame.ImportProp(
                    section,
                    "A992Fy50",
                    "AISC16.xml",
                    section
                )
                if ret != 0:
                    logger.warning(f"Failed to import section {section}")
                    
        # Step 2 & 3: Create auto-select lists and assign to column groups
        for location, frames in columns_by_location.items():
            group_name = f"{location.capitalize()} Columns"
            # Create group and assign frames
            self.sap_model.create_assign_section_group(
                group_name=group_name,
                frames=frames
            )
            
            # Create the auto-select list
            section_list = column_auto_select_lists[group_name]
            auto_list_name = f"AUTO_{group_name}"
            ret = self.sap_model.PropFrame.SetAutoSelectSteel(
                auto_list_name,
                len(section_list),
                section_list,
                section_list[0]  # Start with smallest section
            )
            
            # Assign the auto-select list to the group
            ret = self.sap_model.FrameObj.SetSection(group_name, auto_list_name, 1)  # 1 = apply to group
            logger.info(f"Assigned auto-select list {auto_list_name} to {group_name}")
        
        # Step 5: Run the analysis.
        # Important: Always save the model before running the analysis.
        self.sap_model.File.Save(self.model_path)
        self.sap_model.Analyze.RunAnalysis()
        
        # Step 6: Run Steel Design to select optimal sections from auto-select lists
        # Set the design code explicitly before running design
        self.sap_model.DesignSteel.SetCode("AISC 360-16")
        ret = self.sap_model.DesignSteel.StartDesign() #"AISC 360-16"
        
        # Step 7: Check final selected sections and analyze utilization ratios
        # First, get basic section information for a sample of frames
        for length, frames in beams_by_length.items():
            if frames:
                group_name = f"{int(length)}ft Beams"
                sample_frame = frames[0]
                section_name, auto_list, ret = self.sap_model.FrameObj.GetSection(sample_frame)
                logger.info(f"Sample {group_name}: {sample_frame} - Selected section: {section_name} (from Auto List: {auto_list})")
                
        for location, frames in columns_by_location.items():
            if frames:
                sample_frame = frames[0]
                section_name, auto_list, ret = self.sap_model.FrameObj.GetSection(sample_frame)
                logger.info(f"Sample {location} column: {sample_frame} - Selected section: {section_name} (from Auto List: {auto_list})")
        
        # Now add detailed utilization ratio analysis for each group
        logger.info("==== UTILIZATION RATIO ANALYSIS ====")
        
        # Analyze beam groups
        for length, frames in beams_by_length.items():
            if frames:
                group_name = f"{int(length)}ft Beams"
                self.sap_model.analyze_group_utilization(group_name, "Beam")
        
        # Analyze column groups
        for location in columns_by_location.keys():
            if columns_by_location[location]:
                group_name = f"{location.capitalize()} Columns"
                self.sap_model.analyze_group_utilization(group_name, "Column")
        
        return True

if __name__ == "__main__":
    sap_test = SAPTest()
    sap_test.run_sap_script() 