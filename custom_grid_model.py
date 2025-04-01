import logging
import os
import traceback
from typing import Optional

import comtypes.client
import comtypes.gen.SAP2000v1 as SAP2000

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class e3DFrameType:
    OpenFrame = 0
    PerimeterFrame = 1
    BeamSlab = 2
    FlatPlate = 3

class Units:
    kip_ft_F = 4

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
            self.sap_model = self.sap_object.SapModel
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

    def customize_grid_spacing(self) -> bool:
        try:
            logger.info("Creating model with custom grid spacing...")

            # Define material properties
            try:
                # Define A992Fy50 steel material
                #logger.info(f"hi0 {SAP2000.eMatType_Steel}")
                #logger.info(f"hi1 {SAP2000.eMatType}")
                self.sap_model.PropMaterial.SetMaterial("A992Fy50", SAP2000.eMatType_Steel) # 1 also works
                self.sap_model.PropMaterial.SetMPIsotropic("A992Fy50", 4176000.0, 0.3, 0.00000650)
                #logger.info(f" SetMPIsotropic")
                # Elastic (isotropic) properties from dialog:
                # E = 4176000.0 (Modulus Of Elasticity)
                # U = 0.3 (Poisson's Ratio)
                # A = 6.500E-06 (Coefficient Of Thermal Expansion)
                self.sap_model.PropMaterial.SetMPIsotropic("A992Fy50", 4176000.0, 0.3, 6.5e-6)
                logger.info(f" SetMPIsotropic")
                # Mass per Unit Volume = 0.0152
               
                # Set additional steel properties
                # SSType =1 is for a simple bilinear stress-strain curve — appropriate unless you're defining a custom multilinear curve.
                # self.sap_model.PropMaterial.SetOSteel_1(str("Steel"), int(1), float(7200.0), float(9350.0), float(7920.0), float(10296.0))
                # logger.info(f" SetOSteel_1")
                # Fy = 7200.0 (Minimum Yield Stress)
                # Fu = 9350.0 (Minimum Tensile Stress)
                # Fye = 7920.0 (Expected Yield Stress)
                # Fue = 10296.0 (Expected Tensile Stress)
                
                logger.info("Material properties defined successfully")
            except Exception as e:
                logger.warning(f"Error defining material properties: {e}")
                # Continue anyway - SAP2000 should use defaults

            # Define load patterns
            try:
                self.sap_model.LoadPatterns.Delete("MODAL")
                logger.info(f" load {SAP2000.eLoadPatternType}")
                # for attr in dir(SAP2000.eLoadPatternType):
                #     if not attr.startswith("_"):
                #         value = getattr(SAP2000.eLoadPatternType, attr)
                #         if isinstance(value, int):
                #             logger.info(f"{attr} = {value}")
                logger.info(f" load {SAP2000.eLoadPatternType_Dead}")
                logger.info(f" load {SAP2000.eLoadPatternType.eLoadPatternType_Dead}")
                # Define DEAD load pattern with self-weight
                self.sap_model.LoadPatterns.Add("DEAD", SAP2000.eLoadPatternType_Dead, 1.0)  # 1 = Dead, 
                #SAP2000.eLoadPatternType.Dead resolves to 1, which is correct.
                # The third argument 1.0 means full self-weight is included.
                # The fourth argument True adds a linear static load case for this pattern.
                #      self-weight multiplier = 1.0
                logger.info(f" DEAD")
                # Define LIVE load pattern without self-weight
                self.sap_model.LoadPatterns.Add("LIVE", SAP2000.eLoadPatternType_Live, 0.0)  # 3 = Live, self-weight multiplier = 0.0
                logger.info(f" LIVE")
                # Try to remove MODAL load pattern if it exists. Modal analysis is typically used for dynamic analysis,

                logger.info("Load patterns configured successfully")
            except Exception as e:
                logger.warning(f"Failed to configure load patterns: {e}")

            # Define frame sections
            try:
                self.sap_model.PropFrame.SetRectangle("COLUMN", "A992Fy50", 2, 2)  # 24" x 24" column
                self.sap_model.PropFrame.SetRectangle("BEAM", "A992Fy50", 1.5, 1)  # 18" x 12" beam
                logger.info("Frame sections defined successfully")
            except Exception as e:
                logger.warning(f"Error defining frame sections: {e}")
                # Continue anyway - SAP2000 should use defaults

            # Define custom grid coordinates
            x_coords = [0, 24, 48, 72, 84, 96, 120, 144, 168]
            y_coords = [0, 22, 40, 54, 64]
            z_coords = [0, 18, 30, 42]
            logger.info(f"Using custom grid: {len(x_coords)} X-coords, {len(y_coords)} Y-coords, {len(z_coords)} Z-coords")

            # Create columns first
            logger.info("Creating columns...")
            columns_created = 0
            for x in x_coords:
                for y in y_coords:
                    # Create column segments for each floor
                    for i in range(len(z_coords)-1):
                        try:
                            # Create column from current floor to next floor
                            ret = self.sap_model.FrameObj.AddByCoord(
                                x, y, z_coords[i],      # Bottom point of this segment
                                x, y, z_coords[i+1],    # Top point of this segment
                                ""  # Auto-name
                            )
                            col_name = ret[0]
                            
                            # Set section property
                            try:
                                self.sap_model.FrameObj.SetSection(col_name, "COLUMN")
                            except Exception as col_sec_err:
                                logger.warning(f"Could not set section for column {col_name}: {col_sec_err}")
                            
                            columns_created += 1
                        except Exception as col_err:
                            logger.warning(f"Error creating column at ({x},{y}) between z={z_coords[i]} and z={z_coords[i+1]}: {col_err}")
            
            logger.info(f"Created {columns_created} column segments")

            # After creating all columns, add:
            self.apply_automatic_base_restraints()
            # it did not work: i could not find view options to make joints visible
            # self.make_joints_visible()

            # Create beams at each floor level
            logger.info("Creating beams...")
            beams_created = 0
            
            # Create X-direction beams
            for z in z_coords[1:]:  # Skip the base level
                for y in y_coords:
                    for i in range(len(x_coords)-1):
                        try:
                            ret = self.sap_model.FrameObj.AddByCoord(
                                x_coords[i], y, z,     # Start point
                                x_coords[i+1], y, z,   # End point
                                ""  # Auto-name
                            )
                            beam_name = ret[0]
                            try:
                                self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                            except Exception as beam_sec_err:
                                logger.warning(f"Could not set section for beam {beam_name}: {beam_sec_err}")
                            beams_created += 1
                        except Exception as beam_err:
                            logger.warning(f"Error creating X-beam at ({x_coords[i]},{y},{z}): {beam_err}")
                
                # Create Y-direction beams
                for x in x_coords:
                    for i in range(len(y_coords)-1):
                        try:
                            ret = self.sap_model.FrameObj.AddByCoord(
                                x, y_coords[i], z,     # Start point
                                x, y_coords[i+1], z,   # End point
                                ""  # Auto-name
                            )
                            beam_name = ret[0]
                            try:
                                self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                            except Exception as beam_sec_err:
                                logger.warning(f"Could not set section for beam {beam_name}: {beam_sec_err}")
                            beams_created += 1
                        except Exception as beam_err:
                            logger.warning(f"Error creating Y-beam at ({x},{y_coords[i]},{z}): {beam_err}")
            
            logger.info(f"Created {beams_created} beams")

            # Create floor areas with appropriate loads
            logger.info("Creating floor areas with loads...")
            areas_created = 0
            
            # Create shell areas at each floor level (skip the base level at z=0)
            for z_index, z in enumerate(z_coords[1:], 1):
                # Determine if this is the roof level (the highest level)
                is_roof = (z_index == len(z_coords) - 1)
                
                # Loop through each bay defined by grid lines
                for i in range(len(x_coords) - 1):
                    for j in range(len(y_coords) - 1):
                        try:
                            # Define the coordinates arrays for the 4 corners
                            x_array = [x_coords[i], x_coords[i+1], x_coords[i+1], x_coords[i]]
                            y_array = [y_coords[j], y_coords[j], y_coords[j+1], y_coords[j+1]]
                            z_array = [z, z, z, z]  # All points at the same elevation
                            
                            # Create the area
                            ret = self.sap_model.AreaObj.AddByCoord(
                                4,          # Number of points
                                x_array,    # X coordinates array
                                y_array,    # Y coordinates array
                                z_array,    # Z coordinates array
                                ""          # Auto-name
                            )
                            area_name = ret[3]
                            
                            # Apply dead load (75 psf)
                            self.sap_model.AreaObj.SetLoadUniform(
                                area_name,    # Area name
                                "DEAD",       # Load pattern
                                75.0,         # Load value (psf)
                                6,            # Direction (6 = Global Z)
                                True,         # Replace existing load
                                "Global"      # Coordinate system
                            )
                            
                            # Apply live load based on level
                            live_load = 20.0 if is_roof else 50.0  # 20 psf for roof, 50 psf for floors
                            self.sap_model.AreaObj.SetLoadUniform(
                                area_name,    # Area name
                                "LIVE",       # Load pattern
                                live_load,    # Load value (psf)
                                6,            # Direction (6 = Global Z)
                                True,         # Replace existing load
                                "Global"      # Coordinate system
                            )
                            
                            areas_created += 1
                        except Exception as area_err:
                            logger.warning(f"Error creating area at level {z}, bay ({i},{j}): {area_err}")
            
            logger.info(f"Created {areas_created} floor/roof areas with appropriate loads")

            # Refresh view - try different available methods
            try:
                self.sap_model.View.RefreshView(0, False)
                logger.info("View refreshed successfully")
            except Exception as view_err:
                logger.warning(f"Could not refresh view: {view_err}")
            
            # Try to set 3D view if available - handle this error gracefully
            try:
                # First attempt using Set3DView method
                self.sap_model.View.Set3DView(1, 0)
                logger.info("3D view set successfully")
            except AttributeError:
                # If Set3DView is not available, try alternative methods
                logger.info("Set3DView not available in this SAP2000 version, trying alternatives")
                try:
                    # Try using SetView method if available
                    self.sap_model.View.SetView(1)  # Try a standard view
                    logger.info("View set using alternative method")
                except:
                    # If all visualization attempts fail, just continue
                    logger.info("Could not set specific view - model is still created correctly")
            except Exception as e:
                # Handle any other errors with the view setting
                logger.warning(f"Error setting 3D view: {e}")
            
            # After creating all frames but before refreshing view:
            self.create_frame_groups()
            self.create_and_assign_sections()

            logger.info("Model creation with custom grid spacing completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error creating model with custom grid spacing: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return False

    def apply_automatic_base_restraints(self):
        """
        Automatically assign fixed supports at the base of all columns in the model.
        This restrains all translational and rotational degrees of freedom at the base nodes.
        """
        try:
            logger.info("Applying automatic base restraints...")
            
            # Get all frame objects (columns and beams)
            frame_names = self.sap_model.FrameObj.GetNameList()[1]
            
            # Track how many restraints were applied
            restraints_applied = 0
            
            # For each frame, check if it's a column by checking if its bottom point is at z=0
            for frame_name in frame_names:
                # Get points of the frame
                point_names = self.sap_model.FrameObj.GetPoints(frame_name, "", "")[0:2]
                
                # Get coordinates of each point
                for point_name in point_names:
                    xyz = self.sap_model.PointObj.GetCoordCartesian(point_name)[0:3]
                    
                    # If point is at z=0, it's a column base - apply restraint
                    if abs(xyz[2]) < 0.001:  # Check if z-coordinate is approximately 0
                        # Set fixed restraint (all DOFs)
                        #logger.info(f"Found column base at {point_name}")
                        restraint = [True, True, True, True, True, True]
                        self.sap_model.PointObj.SetRestraint(point_name, restraint)
                        restraints_applied += 1
                        break  # Only need to restrain one end of the column
            
            logger.info(f"Applied fixed restraints to {restraints_applied} column base points")
            return True
        except Exception as e:
            logger.error(f"Error applying automatic base restraints: {e}")
            return False

    def make_joints_visible(self):
        """
        Make all joints visible in the model by setting the joint display options.
        """
        try:
            logger.info("Making joints visible...")
            
            # Alternative approach using View Options method
            # First parameter is for setting options (1) rather than getting them (0)
            # Second parameter controls which options to set (2 for object options)
            # Then various option flags...
            
            # Get current view options first
            ret = self.sap_model.View.GetViewOptions(0, 2)
            if ret[0] == 0:  # Success
                # We have current view options, now set them with Invisible turned OFF
                # Keep other settings the same
                
                # Show joints, show restraints, don't make them invisible
                ret = self.sap_model.View.SetViewOptions(1, 2, True, ret[3], ret[4], ret[5], 
                                                       ret[6], ret[7], ret[8], ret[9], True, 
                                                       False, ret[12], ret[13])
                
                # Refresh view to apply changes
                self.sap_model.View.RefreshView(0, False)
                
                logger.info("Joints set to visible with restraints showing")
                return True
            else:
                logger.warning("Failed to get current view options")
                return False
        except Exception as e:
            logger.error(f"Error making joints visible: {e}")
            return False

    def create_frame_groups(self):
        """
        Automatically creates groups and assigns frames to them:
        1. Creates beam groups by length (10ft, 12ft, 14ft, 18ft, 22ft, 24ft)
        2. Creates column groups by location (Corner, Edge, Interior)
        3. Assigns each frame to appropriate group
        4. Reports assignment counts for verification
        """
        try:
            logger.info("Starting frame group creation and assignment process...")
            
            # Step 1: Define and create all groups in SAP2000
            groups = [
                "10ft Beams", "12ft Beams", "14ft Beams", 
                "18ft Beams", "22ft Beams", "24ft Beams",
                "Corner Columns", "Edge Columns", "Interior Columns"
            ]
            for group in groups:
                self.sap_model.GroupDef.SetGroup(group)
            logger.info("Groups created successfully")

            # Step 2: Initialize counters for verification
            beam_counts = {
                "10ft": 0, "12ft": 0, "14ft": 0, 
                "18ft": 0, "22ft": 0, "24ft": 0
            }
            column_counts = {"Corner": 0, "Edge": 0, "Interior": 0}
            
            # Step 3: Get all frame objects from model
            frames = self.sap_model.FrameObj.GetNameList()[1]
            
            # Step 4: Process each frame
            for frame in frames:
                # Get coordinates of frame endpoints
                points = self.sap_model.FrameObj.GetPoints(frame)[0:2]
                point1_coords = self.sap_model.PointObj.GetCoordCartesian(points[0])[0:3]
                point2_coords = self.sap_model.PointObj.GetCoordCartesian(points[1])[0:3]
                
                # Determine if frame is column (vertical) or beam (horizontal)
                is_column = abs(point1_coords[2] - point2_coords[2]) > 0.1
                
                if is_column:
                    # Process Columns
                    x, y = point1_coords[0], point1_coords[1]
                    
                    # Check column position (corner, edge, or interior)
                    is_corner = (x in [0, 168] and y in [0, 64])  # At building corners
                    is_edge = (x in [0, 168] or y in [0, 64])     # Along perimeter
                    
                    # Assign to appropriate column group
                    if is_corner:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "Corner Columns")
                        column_counts["Corner"] += 1
                    elif is_edge:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "Edge Columns")
                        column_counts["Edge"] += 1
                    else:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "Interior Columns")
                        column_counts["Interior"] += 1
                else:
                    # Process Beams
                    # Calculate beam length using Pythagorean theorem
                    dx = point2_coords[0] - point1_coords[0]
                    dy = point2_coords[1] - point1_coords[1]
                    length = (dx**2 + dy**2)**0.5
                    
                    # Assign to appropriate beam group based on length
                    # Using 1ft tolerance for length matching
                    if abs(length - 10) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "10ft Beams")
                        beam_counts["10ft"] += 1
                    elif abs(length - 12) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "12ft Beams")
                        beam_counts["12ft"] += 1
                    elif abs(length - 14) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "14ft Beams")
                        beam_counts["14ft"] += 1
                    elif abs(length - 18) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "18ft Beams")
                        beam_counts["18ft"] += 1
                    elif abs(length - 22) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "22ft Beams")
                        beam_counts["22ft"] += 1
                    elif abs(length - 24) < 1:
                        self.sap_model.FrameObj.SetGroupAssign(frame, "24ft Beams")
                        beam_counts["24ft"] += 1
            
            # Step 5: Report results
            logger.info("Frame group assignment completed. Member counts:")
            logger.info("Beam Groups:")
            for length, count in beam_counts.items():
                logger.info(f"  {length} Beams: {count} members")
            logger.info("Column Groups:")
            for type, count in column_counts.items():
                logger.info(f"  {type} Columns: {count} members")
            
        except Exception as e:
            logger.error(f"Error in frame group creation/assignment: {e}")
            logger.error(f"Stack trace: {traceback.format_exc()}")

    def create_and_assign_sections(self):
        """
        Creates frame section properties and assigns them to corresponding groups.
        Imports sections from SAP2000 library and assigns specific sections to each group.
        """
        try:
            logger.info("Starting frame section creation and assignment process...")

            # Step 1: Define section assignments for each group
            section_assignments = {
                "24ft Beams": "W24X76",
                "22ft Beams": "W21X44",
                "18ft Beams": "W18X40",
                "14ft Beams": "W14X34",
                "10ft Beams": "W10X33",
                "Corner Columns": "W10X12",
                "Edge Columns": "W12X190",
                "Interior Columns": "W14X193"
            }

            # Step 2: Import sections
            for group_name, section_name in section_assignments.items():
                try:
                    ret = self.sap_model.PropFrame.ImportProp(
                        section_name,          # Changed: Use section_name instead of group_name
                        "A992Fy50",
                        "AISC16.xml",
                        section_name
                    )
                    
                    if ret != 0:
                        logger.warning(f"Failed to import section {section_name} for {group_name}")
                    else:
                        logger.info(f"Successfully imported section {section_name} for {group_name}")

                except Exception as e:
                    logger.error(f"Error importing section for {group_name}: {str(e)}")

            # Step 3: Assign sections to groups
            for group_name, section_name in section_assignments.items():
                try:
                    # Select all frames in the group
                    self.sap_model.SelectObj.Group(group_name)
                    
                    # Assign the section to selected frames - using section_name instead of group_name
                    ret = self.sap_model.FrameObj.SetSection("", section_name)  # Changed this line
                    if ret != 0:
                        logger.warning(f"Failed to assign section {section_name} to group {group_name}")
                    else:
                        logger.info(f"Successfully assigned section {section_name} to group {group_name}")
                    
                    # Clear selection
                    self.sap_model.SelectObj.ClearSelection()
                    
                except Exception as e:
                    logger.error(f"Error assigning section to {group_name}: {str(e)}")

            logger.info("Frame section assignment completed")
            return True

        except Exception as e:
            logger.error(f"Error in frame section creation and assignment: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return False

    def run_sap_script(self) -> Optional[str]:
        if not self._connected:
            logger.error("Not connected to SAP2000.")
            return None
        try:
            # Initialize and create blank model
            self.sap_model.InitializeNewModel(Units.kip_ft_F)
            self.sap_model.File.NewBlank()
            logger.info("Created new blank model")
            
            # Create the custom grid model
            success = self.customize_grid_spacing()
            if not success:
                logger.warning("Failed to create model with custom grid spacing, but will still save whatever was created.")
            
            # Save the model
            self.sap_model.File.Save(self.model_path)
            logger.info(f"Model saved to: {self.model_path}")
            
            # Provide some feedback
            if success:
                return "Commands executed successfully - custom grid model created"
            else:
                return "Commands completed with warnings - model may be incomplete"
            
        except Exception as e:
            logger.error(f"Error executing commands: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return None

def main():
    sap_test = SAPTest()
    result = sap_test.run_sap_script()
    print(f"Execution result: {result}")

if __name__ == "__main__":
    main()
