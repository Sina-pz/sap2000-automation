import logging
import os
import traceback
from typing import Optional

import comtypes.client
import comtypes.gen.SAP2000v1 as SAP2000

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

        # Step 1: Create a new material as steel. 
        self.sap_model.PropMaterial.SetMaterial("Steel", int(SAP2000.eMatType_Steel))
        self.sap_model.PropMaterial.SetMPIsotropic("Steel", float(4176000.0), float(0.3), float(0.00000650))

        # Step 2: Delete the default load patterns and add new ones.
        self.sap_model.LoadPatterns.Delete("MODAL")
        self.sap_model.LoadPatterns.Add("DEAD", int(SAP2000.eLoadPatternType_Dead), 1.0)  # 1 = Dead, 
        self.sap_model.LoadPatterns.Add("LIVE", int(SAP2000.eLoadPatternType_Live), 0.0)  # 3 = Live, self-weight multiplier = 0.0


        # Step 3: Define the column and beam properties.
        self.sap_model.PropFrame.SetRectangle("COLUMN", "A992Fy50", 2, 2)  # 24" x 24" column
        self.sap_model.PropFrame.SetRectangle("BEAM", "A992Fy50", 1.5, 1)  # 18" x 12" beam


        # Step 4:  Define columns and beams
        x_coords = [0, 24, 48, 72, 84, 96, 120, 144, 168]
        y_coords = [0, 22, 40, 54, 64]
        z_coords = [0, 18, 30, 42]
        columns_created = 0
        for x in x_coords:
            for y in y_coords:
                ret = self.sap_model.FrameObj.AddByCoord(
                    x, y, z_coords[0],   # Bottom point
                    x, y, z_coords[-1],  # Top point
                    ""  # Auto-name
                )
                col_name = ret[0]
                self.sap_model.FrameObj.SetSection(col_name, "COLUMN")
                columns_created += 1

        frame_names = self.sap_model.FrameObj.GetNameList()[1]
        restraints_applied = 0
        for frame_name in frame_names:
            # Get points of the frame
            point_names = self.sap_model.FrameObj.GetPoints(frame_name, "", "")[0:2]
            # Get coordinates of each point
            for point_name in point_names:
                xyz = self.sap_model.PointObj.GetCoordCartesian(point_name)[0:3]
                
                # If point is at z=0, it's a column base - apply restraint
                if abs(xyz[2]) < 0.001:  # Check if z-coordinate is approximately 0
                    # Set fixed restraint (all DOFs)
                    restraint = [True, True, True, False, False, False]
                    self.sap_model.PointObj.SetRestraint(point_name, restraint)
                    restraints_applied += 1
                    break  # Only need to restrain one end of the column

        beams_created = 0
        # Create X-direction beams
        for z in z_coords[1:]:  # Skip the base level
            for y in y_coords:
                for i in range(len(x_coords)-1):
                    ret = self.sap_model.FrameObj.AddByCoord(
                        x_coords[i], y, z,     # Start point
                        x_coords[i+1], y, z,   # End point
                        ""  # Auto-name
                    )
                    beam_name = ret[0]
                    self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                    beams_created += 1
            # Create Y-direction beams
            for x in x_coords:
                for i in range(len(y_coords)-1):
                    ret = self.sap_model.FrameObj.AddByCoord(
                        x, y_coords[i], z,     # Start point
                        x, y_coords[i+1], z,   # End point
                        ""  # Auto-name
                    )
                    beam_name = ret[0]
                    self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                    beams_created += 1
        
        # Step 5: Create floor areas with appropriate loads
        x_coords = [0, 24, 48, 72, 84, 96, 120, 144, 168]
        y_coords = [0, 22, 40, 54, 64]
        z_coords = [0, 18, 30, 42]  # Floor levels
        
        # Track how many areas were created
        areas_created = 0
        # Create shell areas at each floor level (skip the base level at z=0)
        for z_index, z in enumerate(z_coords[1:], 1):
            # Determine if this is the roof level (the highest level)
            is_roof = (z_index == len(z_coords) - 1)
            # Loop through each bay defined by grid lines
            for i in range(len(x_coords) - 1):
                for j in range(len(y_coords) - 1):
                    # Define the coordinates arrays for the 4 corners
                    x_array = [x_coords[i], x_coords[i+1], x_coords[i+1], x_coords[i]]
                    y_array = [y_coords[j], y_coords[j], y_coords[j+1], y_coords[j+1]]
                    z_array = [z, z, z, z]  # All points at the same elevation
                    
                    # Create the area - using the proper API syntax
                    ret = self.sap_model.AreaObj.AddByCoord(
                        4,          # Number of points
                        x_array,    # X coordinates array
                        y_array,    # Y coordinates array
                        z_array,    # Z coordinates array
                        ""          # Auto-name
                    )
                    
                    # Extract the area name from the return value (4th element, index 3)
                    area_name = ret[3]
                    
                    # Apply dead load (75 psf) - Direction 6 is Global Z
                    self.sap_model.AreaObj.SetLoadUniform(
                        area_name,    # Area name
                        "DEAD",       # Load pattern
                        75.0,         # Load value (psf)
                        6,            # Direction (6 = Global Z)
                        True,         # Replace existing load
                        "Global"      # Coordinate system
                    )
                    
                    # Apply live load
                    if is_roof:
                        # Roof live load (20 psf)
                        self.sap_model.AreaObj.SetLoadUniform(
                            area_name,    # Area name
                            "LIVE",       # Load pattern  
                            20.0,         # Load value (psf)
                            6,            # Direction (6 = Global Z)
                            True,         # Replace existing load
                            "Global"      # Coordinate system
                        )
                    else:
                        # Floor live load (50 psf)
                        self.sap_model.AreaObj.SetLoadUniform(
                            area_name,    # Area name
                            "LIVE",       # Load pattern
                            50.0,         # Load value (psf)
                            6,            # Direction (6 = Global Z)
                            True,         # Replace existing load
                            "Global"      # Coordinate system
                        )
                    
                    areas_created += 1

        self.sap_model.View.RefreshView(0, False)


    def apply_automatic_base_restraints(self):
        """
        Automatically assign fixed supports at the base of all columns in the model.
        This restrains all translational and rotational degrees of freedom at the base nodes.
        """
        try:
            print("Applying automatic base restraints...")
            
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
                        restraint = [True, True, True, False, False, False]
                        self.sap_model.PointObj.SetRestraint(point_name, restraint)
                        restraints_applied += 1
                        break  # Only need to restrain one end of the column
            
            print(f"Applied fixed restraints to {restraints_applied} column base points")
            return True
        except Exception as e:
            print(f"Error applying automatic base restraints: {e}")
            return False

    def make_joints_visible(self):
        """
        Make all joints visible in the model by setting the joint display options.
        """
        try:
            print("Making joints visible...")
            
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
                
                print("Joints set to visible with restraints showing")
                return True
            else:
                print("Failed to get current view options")
                return False
        except Exception as e:
            print(f"Error making joints visible: {e}")
            return False

    def create_floor_areas(self):
        """
        Create floor areas at each story level and apply the specified loads:
        - Floor dead load = 75 psf
        - Floor live load = 50 psf  
        - Roof live load = 20 psf (for the top level only)
        """
        try:
            print("Creating floor areas with loads...")
            
            # Get your grid coordinates from the existing model
            x_coords = [0, 24, 48, 72, 84, 96, 120, 144, 168]
            y_coords = [0, 22, 40, 54, 64]
            z_coords = [0, 18, 30, 42]  # Floor levels
            
            # Track how many areas were created
            areas_created = 0
            
            # Create shell areas at each floor level (skip the base level at z=0)
            for z_index, z in enumerate(z_coords[1:], 1):
                # Determine if this is the roof level (the highest level)
                is_roof = (z_index == len(z_coords) - 1)
                
                # Loop through each bay defined by grid lines
                for i in range(len(x_coords) - 1):
                    for j in range(len(y_coords) - 1):
                        # Define the coordinates arrays for the 4 corners
                        x_array = [x_coords[i], x_coords[i+1], x_coords[i+1], x_coords[i]]
                        y_array = [y_coords[j], y_coords[j], y_coords[j+1], y_coords[j+1]]
                        z_array = [z, z, z, z]  # All points at the same elevation
                        
                        # Create the area - using the proper API syntax
                        ret = self.sap_model.AreaObj.AddByCoord(
                            4,          # Number of points
                            x_array,    # X coordinates array
                            y_array,    # Y coordinates array
                            z_array,    # Z coordinates array
                            ""          # Auto-name
                        )
                        
                        # Extract the area name from the return value (4th element, index 3)
                        area_name = ret[3]
                        
                        # Apply dead load (75 psf) - Direction 6 is Global Z
                        self.sap_model.AreaObj.SetLoadUniform(
                            area_name,    # Area name
                            "DEAD",       # Load pattern
                            75.0,         # Load value (psf)
                            6,            # Direction (6 = Global Z)
                            True,         # Replace existing load
                            "Global"      # Coordinate system
                        )
                        
                        # Apply live load
                        if is_roof:
                            # Roof live load (20 psf)
                            self.sap_model.AreaObj.SetLoadUniform(
                                area_name,    # Area name
                                "LIVE",       # Load pattern  
                                20.0,         # Load value (psf)
                                6,            # Direction (6 = Global Z)
                                True,         # Replace existing load
                                "Global"      # Coordinate system
                            )
                        else:
                            # Floor live load (50 psf)
                            self.sap_model.AreaObj.SetLoadUniform(
                                area_name,    # Area name
                                "LIVE",       # Load pattern
                                50.0,         # Load value (psf)
                                6,            # Direction (6 = Global Z)
                                True,         # Replace existing load
                                "Global"      # Coordinate system
                            )
                        
                        areas_created += 1
            
            print(f"Created {areas_created} floor/roof areas with appropriate loads")
            return True
        except Exception as e:
            print(f"Error creating floor areas: {str(e)}")
            print(f"Stack trace: {traceback.format_exc()}")
            return False

    def run_sap_script(self) -> Optional[str]:
        if not self._connected:
            print("Not connected to SAP2000.")
            return None
        try:
            # Initialize and create blank model
            self.sap_model.InitializeNewModel(4)
            self.sap_model.File.NewBlank()
            print("Created new blank model")
            
            # Create the custom grid model
            success = self.customize_grid_spacing()
            if not success:
                print("Failed to create model with custom grid spacing, but will still save whatever was created.")
            
            # Save the model
            self.sap_model.File.Save(self.model_path)
            print(f"Model saved to: {self.model_path}")
            
            # Provide some feedback
            if success:
                return "Commands executed successfully - custom grid model created"
            else:
                return "Commands completed with warnings - model may be incomplete"
            
        except Exception as e:
            print(f"Error executing commands: {str(e)}")
            print(f"Stack trace: {traceback.format_exc()}")
            return None

def main():
    sap_test = SAPTest()
    result = sap_test.run_sap_script()
    print(f"Execution result: {result}")

if __name__ == "__main__":
    main()
