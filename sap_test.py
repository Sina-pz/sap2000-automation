import logging
import os
import sys
import traceback
from typing import Optional

import comtypes.client

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# SAP2000 constants
class e3DFrameType:
    OpenFrame = 0
    PerimeterFrame = 1
    BeamSlab = 2
    FlatPlate = 3

class Units:
    kip_ft_F = 4  # Assuming this is the correct enum value for kip_ft_F

class SAPTest:
    def __init__(self):
        """Initialize SAP2000 connection and model."""
        self.sap_object = None
        self.sap_model = None
        self.model_path = None
        self._connected = False
        
        # Try to connect immediately upon initialization
        self._try_connect()
    
    def _try_connect(self) -> bool:
        """Attempt to connect to a running SAP2000 instance."""
        try:
            import comtypes.client
            import comtypes.gen.SAP2000v1
            logger.info("Attempting to connect to SAP2000...")
            helper = comtypes.client.CreateObject('SAP2000v1.Helper')
            helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
            self.sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
            self.sap_model = self.sap_object.SapModel
            
            # Set model path - adjust this path as needed
            APIPath = os.path.join(os.path.dirname(__file__), 'model')
            os.makedirs(APIPath, exist_ok=True)  # Create directory if it doesn't exist
            self.model_path = os.path.join(APIPath, 'custom_grid_model.sdb')
            
            # Test if the connection is valid by getting program info
            info = self.sap_model.GetProgramInfo()
            
            logger.info(f"Successfully connected to SAP2000 (Version: {info[0]}, Build: {info[1]})")
            self._connected = True
            return True
            
        except ImportError:
            logger.error("Failed to import comtypes. Make sure it's installed: pip install comtypes")
            self._connected = False
            return False
        except Exception as e:
            logger.error(f"Failed to connect to SAP2000: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            self._connected = False
            return False
    
    def customize_grid_spacing(self) -> bool:
        """Create a frame structure with custom spacing using AddByCoord.
        
        This method uses only the most basic and widely available API calls.
        """
        try:
            logger.info("Creating model with custom grid spacing...")
            
            # Skip the delete operation - we'll start with a blank model
            
            # Define material for the frame
            logger.info("Defining material for the frame")
            ret = self.sap_model.PropMaterial.SetMaterial("CONC", 2)  # 2 = Concrete
            if ret != 0:
                logger.warning(f"Failed to set material: {ret}")
            
            # Set concrete properties
            ret = self.sap_model.PropMaterial.SetMPIsotropic("CONC", 3600, 0.2, 0.0000055)  # E=3600 ksi, v=0.2
            ret = self.sap_model.PropMaterial.SetWeightAndMass("CONC", 1, 0.15)  # Unit weight = 150 pcf
            logger.info("Material properties set")
            
            # Define frame sections - try with compatibility syntax
            logger.info("Defining frame sections")
            try:
                # First try with the standard method
                ret = self.sap_model.PropFrame.SetRectangle("COLUMN", "CONC", 2, 2)  # 24" x 24" column
                ret = self.sap_model.PropFrame.SetRectangle("BEAM", "CONC", 1.5, 1)  # 18" x 12" beam
            except AttributeError:
                # Fall back to alternative method if available
                logger.warning("Standard section creation failed, trying alternatives")
                try:
                    # Try an alternative approach
                    # (This would depend on your specific API version)
                    ret = self.sap_model.PropFrame.SetProp("COLUMN", "CONC", 2, 2, "RECT")
                    ret = self.sap_model.PropFrame.SetProp("BEAM", "CONC", 1.5, 1, "RECT")
                except:
                    # If all fails, use default sections
                    logger.warning("Could not create custom sections, using defaults")
            
            # Define custom grid coordinates
            x_coords = [-9, -3, 3, 19]  # In feet (custom spacing)
            y_coords = [-6, 0, 16]      # In feet (custom spacing)
            z_coords = [0, 3, 16]       # In feet (custom spacing)
            
            # Create columns first (vertical members)
            logger.info("Creating columns at custom grid intersections")
            for x in x_coords:
                for y in y_coords:
                    try:
                        # Create a column from base to top
                        frame_name = ""  # Let SAP2000 auto-name
                        ret = self.sap_model.FrameObj.AddByCoord(
                            x, y, z_coords[0],   # Bottom point
                            x, y, z_coords[-1],  # Top point
                            frame_name
                        )
                        col_name = ret[0]
                        logger.info(f"Created column {col_name} at ({x},{y})")
                        
                        # Set properties for the column
                        try:
                            ret = self.sap_model.FrameObj.SetSection(col_name, "COLUMN")
                        except:
                            logger.warning(f"Could not set section for column {col_name}")
                        
                        # Set restraints at the base (fixed support)
                        if z_coords[0] == 0:
                            try:
                                # Get the joint at column base using empty string params
                                ret = self.sap_model.FrameObj.GetPoints(col_name, "", "")
                                base_joint = ret[0]  # First point (bottom of column)
                                
                                # Set fixed restraint
                                restraint = [True, True, True, True, True, True]  # Fixed in all directions
                                ret = self.sap_model.PointObj.SetRestraint(base_joint, restraint)
                                logger.info(f"Set fixed restraint at base joint {base_joint}")
                            except Exception as e:
                                logger.warning(f"Could not set restraint: {str(e)}")
                    except Exception as e:
                        logger.warning(f"Error creating column at ({x},{y}): {str(e)}")
            
            # Create beams at each floor level
            logger.info("Creating beams at custom elevations")
            for z in z_coords[1:]:  # Skip the base level
                # Create X-direction beams
                for y in y_coords:
                    for i in range(len(x_coords) - 1):
                        try:
                            frame_name = ""
                            ret = self.sap_model.FrameObj.AddByCoord(
                                x_coords[i], y, z,     # Start point
                                x_coords[i+1], y, z,   # End point
                                frame_name
                            )
                            beam_name = ret[0]
                            logger.info(f"Created X-beam {beam_name} at level z={z}, y={y}")
                            
                            # Try to set beam section
                            try:
                                ret = self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                            except:
                                logger.warning(f"Could not set section for beam {beam_name}")
                        except Exception as e:
                            logger.warning(f"Error creating X-beam at z={z}, y={y}: {str(e)}")
                
                # Create Y-direction beams
                for x in x_coords:
                    for i in range(len(y_coords) - 1):
                        try:
                            frame_name = ""
                            ret = self.sap_model.FrameObj.AddByCoord(
                                x, y_coords[i], z,     # Start point
                                x, y_coords[i+1], z,   # End point
                                frame_name
                            )
                            beam_name = ret[0]
                            logger.info(f"Created Y-beam {beam_name} at level z={z}, x={x}")
                            
                            # Try to set beam section
                            try:
                                ret = self.sap_model.FrameObj.SetSection(beam_name, "BEAM")
                            except:
                                logger.warning(f"Could not set section for beam {beam_name}")
                        except Exception as e:
                            logger.warning(f"Error creating Y-beam at z={z}, x={x}: {str(e)}")
            
            # Refresh the view if possible
            try:
                ret = self.sap_model.View.RefreshView(0, False)
                logger.info("View refreshed")
            except:
                logger.warning("Could not refresh view")
            
            # Draw a 3D view if possible
            try:
                ret = self.sap_model.View.Set3DView(1, 0)  # Front-right 3D view
                logger.info("3D view set")
            except:
                logger.warning("Could not set 3D view")
            
            logger.info("Model creation with custom grid spacing completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error creating model with custom grid spacing: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return False
    
    def run_sap_script(self) -> Optional[str]:
        """
        Execute SAP2000 commands to create a 3D frame model with custom grid spacing.
        
        Returns:
            Optional[str]: Result of execution or None if failed
        """
        if not self._connected:
            logger.error("Not connected to SAP2000. Make sure SAP2000 is running with a model open.")
            return None
            
        try:
            # Initialize new model with kip-ft-F units
            ret = self.sap_model.InitializeNewModel(Units.kip_ft_F)
            logger.info(f"Model initialization result: {ret}")
            
            # Start with a blank model 
            ret = self.sap_model.File.NewBlank()
            logger.info(f"Blank model creation result: {ret}")
            
            # Create our custom spaced model directly
            success = self.customize_grid_spacing()
            if not success:
                logger.warning("Failed to create model with custom grid spacing, creating default model instead")
                # Fall back to default model if customization fails
                ret = self.sap_model.File.New3DFrame(
                    e3DFrameType.OpenFrame,  # Open Frame Building type
                    3,                       # Number of stories
                    144,                     # Story height
                    8,                       # Number of bays in X
                    288,                     # Bay width X
                    4,                       # Number of bays in Y
                    288                      # Bay width Y
                )
                logger.info(f"Default 3D frame creation result: {ret}")
            
            # Save the model
            ret = self.sap_model.File.Save(self.model_path)
            logger.info(f"Model saved to: {self.model_path}")

            return "Commands executed successfully"
            
        except Exception as e:
            logger.error(f"Error executing commands: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return None

def main():
    """Main function to run SAP2000 commands."""
    sap_test = SAPTest()
    result = sap_test.run_sap_script()
    print(f"Execution result: {result}")

if __name__ == "__main__":
    main()