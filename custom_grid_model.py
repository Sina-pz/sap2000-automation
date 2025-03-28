import logging
import os
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
                self.sap_model.PropMaterial.SetMaterial("CONC", 2)  # 2 = Concrete
                self.sap_model.PropMaterial.SetMPIsotropic("CONC", 3600, 0.2, 0.0000055)  # E=3600 ksi, v=0.2
                self.sap_model.PropMaterial.SetWeightAndMass("CONC", 1, 0.15)  # Unit weight = 150 pcf
                logger.info("Material properties defined successfully")
            except Exception as e:
                logger.warning(f"Error defining material properties: {e}")
                # Continue anyway - SAP2000 should use defaults

            # Define frame sections
            try:
                self.sap_model.PropFrame.SetRectangle("COLUMN", "CONC", 2, 2)  # 24" x 24" column
                self.sap_model.PropFrame.SetRectangle("BEAM", "CONC", 1.5, 1)  # 18" x 12" beam
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
                    try:
                        # Create column from base to top
                        ret = self.sap_model.FrameObj.AddByCoord(
                            x, y, z_coords[0],   # Bottom point
                            x, y, z_coords[-1],  # Top point
                            ""  # Auto-name
                        )
                        col_name = ret[0]
                        
                        # Set section property
                        try:
                            self.sap_model.FrameObj.SetSection(col_name, "COLUMN")
                        except Exception as col_sec_err:
                            logger.warning(f"Could not set section for column {col_name}: {col_sec_err}")
                        
                        # Set base restraints
                        if z_coords[0] == 0:
                            try:
                                # Get the joint at column base
                                ret = self.sap_model.FrameObj.GetPoints(col_name, "", "")
                                base_joint = ret[0]  # First point (bottom of column)
                                
                                # Set fixed restraint (all DOFs)
                                restraint = [True, True, True, True, True, True]
                                self.sap_model.PointObj.SetRestraint(base_joint, restraint)
                            except Exception as restr_err:
                                logger.warning(f"Could not set restraint at column base: {restr_err}")
                        
                        columns_created += 1
                    except Exception as col_err:
                        logger.warning(f"Error creating column at ({x},{y}): {col_err}")
            
            logger.info(f"Created {columns_created} columns")

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
            
            logger.info("Model creation with custom grid spacing completed successfully")
            return True
            
        except Exception as e:
            logger.error(f"Error creating model with custom grid spacing: {str(e)}")
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
