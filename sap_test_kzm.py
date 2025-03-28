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
            self.model_path = os.path.join(APIPath, 'compass_model.sdb')
            
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
    
    def run_sap_script(self) -> Optional[str]:
        """
        Execute hardcoded SAP2000 commands to create a 3D frame model.
        
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
            
            # Create new 3D frame with specified parameters
            ret = self.sap_model.File.New3DFrame(
                e3DFrameType.OpenFrame,  # Open Frame Building type
                3,                       # Number of stories
                144,                     # Story height
                8,                       # Number of bays in X
                288,                     # Bay width X
                4,                       # Number of bays in Y
                288                      # Bay width Y
            )
            logger.info(f"3D frame creation result: {ret}")
            
            # TODO: Figure out how to update the frame sizes

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