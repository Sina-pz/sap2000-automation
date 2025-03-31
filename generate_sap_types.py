import os
import sys
import comtypes.client

def generate_sap_types():
    try:
        print("Generating SAP2000 COM type library...")
        # Create the SAP2000 COM object
        sap_object = comtypes.client.CreateObject('SAP2000v1.Helper')
        # This will automatically generate the type library
        print("Successfully generated SAP2000 COM type library!")
    except Exception as e:
        print(f"Error generating SAP2000 COM type library: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    generate_sap_types() 