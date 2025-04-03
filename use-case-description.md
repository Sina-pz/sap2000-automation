# Structural Analysis and Design Optimization for Steel Frame Buildings

This use case provides automated structural analysis and design optimization for multi-story steel frame buildings with steel deck floor systems. Compass uses SAP2000. While fully autonomous, user has the power of reviewing and revising the progress (directly in SAP) or feedbacking agent (through the Compass texting bot) at any moment.

## Input Requirements

- **Initial Model**: An architectural layout with defined frames and joints already loaded in SAP2000 (this provides the frames and joints)
- **Material Specifications**:
  - Steel type (e.g., ASTM A992Fy50)
  - Material properties if using non-standard materials
- **Loading Requirements**:
  - Dead load values for floors and roof (psf)
  - Live load values for floors and roof (psf)
- **Design Parameters**:
  - Design code to follow (e.g., AISC 360-16)
  - Preferred section series (W-shapes, HSS, etc.)

## Output Deliverables

- **Fully Optimized Model**: Complete SAP2000 model with optimized sections assigned
- **Comprehensive Report**:
  - Summary of assigned section sizes by group
  - Utilization ratios for all structural members
  - Critical design checks and controlling load combinations
  - Deflection analysis results
  - Material quantity takeoffs for cost estimation

## Workflow Process

1. **Foundation Support Setup**

   - Automatically identifies all ground-level columns
   - Applies appropriate restraints to column bases

2. **Floor & Roof System Generation**

   - Detects distinct floor levels throughout the structure
   - Creates floor/roof areas by intelligently identifying enclosed spaces between beams
   - Applies specified dead and live loads to each area (with different values for typical floors vs. roof)

3. **Intelligent Member Grouping**

   - Groups beams by length for efficient section assignment
   - Categorizes columns by position (corner, edge, interior) to reflect different loading conditions
   - Creates logical section groups ready for design optimization

4. **Section Assignment & Optimization**

   - Assigns appropriate initial sections to each member group based on engineering requirements
   - Performs analysis to validate structural performance
   - Optimizes sections to achieve most efficient design while meeting code requirements
   - Iterate over steps 3 and 4 as needed.

5. **Analysis & Documentation**
   - Runs comprehensive structural analysis
   - Generates detailed reports and visualizations
   - Provides summary of results for engineering review

## Flexibility and Constraints

### User Experience (UX)

#### Flexibility

- **No Coding Required**: Engineers can use the tool without any programming knowledge
- **Stepwise Execution**: Process can be run in small, reviewable steps allowing users to verify each stage
- **Interactive Feedback**: Users can provide feedback or make manual adjustments between steps
- **Visual Confirmation**: Results are visually presented in the familiar SAP2000 interface
- **Partial Process Execution**: Users can execute only specific parts of the workflow (e.g., just creating floor areas or only performing optimization)
- **Parameter Customization**: Engineers can adjust load values, section preferences, and other parameters through a simple interface

#### Constraints

- **SAP2000 Dependency**: Requires an active SAP2000 installation and license
- **Initial Model Requirements**: Needs a pre-defined frame model with established joints before automation can begin
- **Limited Undo Capability**: Some operations may not be easily reversible once applied

### Building Geometry

#### Flexibility

- **Multi-Story Support**: Handles buildings with any number of stories and varying floor-to-floor heights
- **Non-Rectangular Layouts**: Processes irregular floor plans including L-shaped, U-shaped, and custom configurations
- **Floor Openings**: Accommodates openings for atriums, stairs, and elevator shafts
- **Mixed-Use Configurations**: Supports different grid patterns and loading requirements for different functional areas
- **Column Grid Variations**: Works with non-uniform column spacing and patterns across the structure
- **Custom Floor Boundaries**: Creates floor areas that follow the exact perimeter of the structural grid, not limited to rectangular shapes

#### Constraints

- **Flat Floor Requirement**: Each floor level must maintain a consistent elevation (no sloped floors within a single story)
- **Conventional Framing Approach**: Best suited for orthogonal framing systems with clearly defined beams and columns
- **Limited Curved Element Support**: May have difficulty with curved architectural elements or non-linear framing
- **Standardized Connection Assumption**: Assumes typical connections without special detailing requirements
- **Standard Column Alignment**: Performs best when columns are generally aligned vertically between floors
- **Base Restraint Uniformity**: Applies consistent restraint conditions to all ground-level columns

### Loading Conditions

#### Flexibility

- **Comprehensive Gravity Load Management**: The system handles both dead and live loads with customizable values for different floors and roof areas. Floor-specific loading can be defined based on intended usage, while the structural self-weight is automatically calculated and incorporated into the analysis.

- **Intelligent Load Application**: Loads are automatically detected and applied to all floor areas based on their location and function in the building. The system generates code-compliant load combinations for accurate design evaluation without requiring manual setup.

#### Constraints

- The system primarily focuses on gravity loads (dead and live), while specialized load types such as wind, seismic, snow, or thermal are not covered.

- Load distribution is limited to uniformly distributed loads across all areas on a floor.

- The system doesn't automatically identify and apply load combinations based on region, project requirements, safety factors, or specific code provisions.

### Analysis Capabilities

#### Flexibility

- The system automatically runs static structural analysis with standard load cases and combinations required by the selected design code. Results are presented directly in the familiar SAP2000 interface where engineers can review detailed member forces, deflections, and other critical values.

#### Constraints

- The system focuses on standard linear static analysis and does not automatically implement advanced analysis types such as P-Delta, non-linear, or dynamic analysis without additional manual configuration.

- Specialized analysis for seismic performance, blast resistance, progressive collapse, or other advanced scenarios requires manual setup beyond the standard automation workflow.

### Design Optimization

#### Flexibility

- The system intelligently groups similar structural members based on engineering principles (beams by length, columns by position) to create practical design groups that balance efficiency with constructability. This automated grouping can be reviewed and modified by the engineer as needed.

- Section selection follows the specified design code (e.g., AISC 360-16) considering strength, serviceability, and stability requirements. The optimization process can target either minimum weight, cost efficiency, or balanced utilization across the structure based on project priorities.

#### Constraints

- The optimization focuses on member sizing rather than overall system optimization, meaning lateral force resisting systems and stability considerations may require additional engineering judgment.

- Connection design and detailing are not automatically generated as part of the optimization process; these aspects must be addressed separately through traditional engineering workflows.

- Special design considerations such as fire protection requirements, vibration performance, or fatigue analysis must be evaluated separately from the standard optimization process.

Contraints:
