from typing import Literal, Optional, Union

import comtypes.client


class LoadPatternType:
    """SAP2000 Load Pattern Types (eLoadPatternType enumeration)"""
    DEAD = 1                        # LTYPE_DEAD
    SUPERDEAD = 2                   # LTYPE_SUPERDEAD
    LIVE = 3                        # LTYPE_LIVE
    REDUCELIVE = 4                  # LTYPE_REDUCELIVE
    QUAKE = 5                       # LTYPE_QUAKE
    WIND = 6                        # LTYPE_WIND
    SNOW = 7                        # LTYPE_SNOW
    OTHER = 8                       # LTYPE_OTHER
    MOVE = 9                        # LTYPE_MOVE
    TEMPERATURE = 10                # LTYPE_TEMPERATURE
    ROOFLIVE = 11                   # LTYPE_ROOFLIVE
    NOTIONAL = 12                   # LTYPE_NOTIONAL
    PATTERNLIVE = 13                # LTYPE_PATTERNLIVE
    WAVE = 14                       # LTYPE_WAVE
    BRAKING = 15                    # LTYPE_BRAKING
    CENTRIFUGAL = 16                # LTYPE_CENTRIFUGAL
    FRICTION = 17                   # LTYPE_FRICTION
    ICE = 18                        # LTYPE_ICE
    WINDONLIVELOAD = 19             # LTYPE_WINDONLIVELOAD
    HORIZONTALEARTHPRESSURE = 20    # LTYPE_HORIZONTALEARTHPRESSURE
    VERTICALEARTHPRESSURE = 21      # LTYPE_VERTICALEARTHPRESSURE
    EARTHSURCHARGE = 22             # LTYPE_EARTHSURCHARGE
    DOWNDRAG = 23                   # LTYPE_DOWNDRAG
    VEHICLECOLLISION = 24           # LTYPE_VEHICLECOLLISION
    VESSELCOLLISION = 25            # LTYPE_VESSELCOLLISION
    TEMPERATUREGRADIENT = 26        # LTYPE_TEMPERATUREGRADIENT
    SETTLEMENT = 27                 # LTYPE_SETTLEMENT
    SHRINKAGE = 28                  # LTYPE_SHRINKAGE
    CREEP = 29                      # LTYPE_CREEP
    WATERLOADPRESSURE = 30          # LTYPE_WATERLOADPRESSURE
    LIVELOADSURCHARGE = 31          # LTYPE_LIVELOADSURCHARGE
    LOCKEDINFORCES = 32             # LTYPE_LOCKEDINFORCES
    PEDESTRIANLL = 33               # LTYPE_PEDESTRIANLL
    PRESTRESS = 34                  # LTYPE_PRESTRESS
    HYPERSTATIC = 35                # LTYPE_HYPERSTATIC
    BOUYANCY = 36                   # LTYPE_BOUYANCY
    STREAMFLOW = 37                 # LTYPE_STREAMFLOW
    IMPACT = 38                     # LTYPE_IMPACT
    CONSTRUCTION = 39               # LTYPE_CONSTRUCTION

class SAPWrapper:
    """Wrapper for SAP2000 API with documented methods"""
    
    def __init__(self):
        self.sap_object = None
        self.sap_model = None
        self._connect()
    
    def _connect(self):
        """Connect to SAP2000 instance"""
        helper = comtypes.client.CreateObject('SAP2000v1.Helper')
        helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)
        self.sap_object = helper.GetObject("CSI.SAP2000.API.SapObject")
        self.sap_model = self.sap_object.SapModel
    
    def add_load_pattern(self, 
                        name: str, 
                        pattern_type: Union[int, LoadPatternType],
                        self_weight_multiplier: float = 0.0,
                        add_load_case: bool = True) -> bool:
        """
        Adds a new load pattern to the model.
        
        Args:
            name: The name for the new load pattern
            pattern_type: Type of load pattern (use LoadPatternType enum)
            self_weight_multiplier: The self weight multiplier for the new load pattern (default: 0.0)
            add_load_case: If True, a linear static load case corresponding to the new load pattern is added (default: True)
            
        Returns:
            bool: True if successful, False if error occurred
            
        Example:
            >>> sap = SAPWrapper()
            >>> sap.add_load_pattern("DEAD", LoadPatternType.DEAD, 1.0)
            >>> sap.add_load_pattern("LIVE", LoadPatternType.LIVE, 0.0)
            
        Note:
            Returns False if the load pattern name already exists in the model
        """
        try:
            ret = self.sap_model.LoadPatterns.Add(name, pattern_type, 
                                                 self_weight_multiplier, add_load_case)
            return ret == 0  # Returns 0 if successful
        except Exception as e:
            print(f"Error adding load pattern: {e}")
            return False

    def set_material_properties(self,
                              name: str,
                              material_type: int,
                              elastic_modulus: float,
                              poisson_ratio: float,
                              thermal_coeff: float,
                              weight_per_vol: float,
                              mass_per_vol: float) -> bool:
        """
        Sets material properties for a material.
        
        Args:
            name: Material name
            material_type: Material type (1=Steel, 2=Concrete, etc)
            elastic_modulus: Modulus of elasticity (E)
            poisson_ratio: Poisson's ratio (U)
            thermal_coeff: Coefficient of thermal expansion (A)
            weight_per_vol: Weight per unit volume
            mass_per_vol: Mass per unit volume
            
        Returns:
            bool: True if successful
        """
        try:
            self.sap_model.PropMaterial.SetMaterial(name, material_type)
            self.sap_model.PropMaterial.SetMPIsotropic(name, elastic_modulus, 
                                                      poisson_ratio, thermal_coeff)
            self.sap_model.PropMaterial.SetWeightAndMass(name, weight_per_vol, 
                                                        mass_per_vol)
            return True
        except Exception as e:
            print(f"Error setting material properties: {e}")
            return False
            
    # Add more wrapped methods as needed... 