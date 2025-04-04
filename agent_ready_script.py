# import logging
# import os
# import traceback
# from typing import Optional, Tuple, List, Dict, Union

# import comtypes.client
# import comtypes.gen.SAP2000v1 as SAP2000

import logging
import math
import os
import traceback
from typing import Dict, List, Optional, Tuple, Union

import comtypes.client
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
import comtypes.gen.SAP2000v1 as SAP2000
# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class CustomSAP2000Model:
    """
    Custom SAP2000 model class that extends the base SAP2000 model with additional functionality.
    """
    def __init__(self, sap_model):
        self._model = sap_model

    def __getattr__(self, name):
        """
        Forward any unknown attribute access to the underlying model.
        """
        return getattr(self._model, name)

    def add_base_restraints(self, restraints=[True, True, True, False, False, False]) -> tuple:
        """
        Add restraints to ground-level column bases by checking frame connectivity.
        Only restrains points that are at the bottom of columns with no frames connecting below them.
        
        Args:
            restraints: List of 6 boolean values for [Ux, Uy, Uz, Rx, Ry, Rz] restraints
                        Default is [True, True, True, False, False, False] (fixed translations, free rotations)
        
        Returns:
            tuple: (list of restrained point names, status code) where status code is 0 for success, 1 for failure
        """
        try:
            # Get all frame names
            number_frames, frame_names, ret = self._model.FrameObj.GetNameList()
            if ret != 0 or not frame_names:
                logger.error("Failed to get frame names from model")
                return ([], 1)

            # Get all point names for later use
            number_points, point_names, ret = self._model.PointObj.GetNameList()
            if ret != 0 or not point_names:
                logger.error("Failed to get point names from model")
                return ([], 1)

            # Create a dictionary to store points and their connected frames
            point_connections = {}
            for point_name in point_names:
                point_connections[point_name] = []

            # Map frames to their points and build connectivity
            for frame_name in frame_names:
                point_i, point_j, ret = self._model.FrameObj.GetPoints(frame_name)
                if ret == 0:
                    point_connections[point_i].append(frame_name)
                    point_connections[point_j].append(frame_name)

            # Find and restrain ground-level column bases
            restrained_points = []
            for frame_name in frame_names:
                point_i, point_j, ret = self._model.FrameObj.GetPoints(frame_name)
                if ret != 0:
                    continue

                # Get coordinates of both points
                x_i, y_i, z_i, ret_i = self._model.PointObj.GetCoordCartesian(point_i)
                x_j, y_j, z_j, ret_j = self._model.PointObj.GetCoordCartesian(point_j)
                
                if ret_i != 0 or ret_j != 0:
                    continue

                # Determine which point is the bottom point
                bottom_point = point_i if z_i < z_j else point_j
                top_point = point_j if z_i < z_j else point_i

                # Check if this is a ground-level column by verifying no frames connect to the bottom point
                if len(point_connections[bottom_point]) == 1:  # Only connected to current frame
                    _, ret = self._model.PointObj.SetRestraint(bottom_point, restraints)
                    
                    if ret == 0:
                        restrained_points.append(bottom_point)
                        logger.info(f"Applied restraints to ground-level column base {bottom_point} "
                                  f"at coordinates ({x_i if bottom_point == point_i else x_j}, "
                                  f"{y_i if bottom_point == point_i else y_j}, "
                                  f"{z_i if bottom_point == point_i else z_j})")
                    else:
                        logger.warning(f"Failed to apply restraints to point {bottom_point}")

            logger.info(f"Successfully restrained {len(restrained_points)} ground-level column bases")
            return (restrained_points, 0)

        except Exception as e:
            logger.error(f"Error in add_base_restraints: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return ([], 1)

    def identify_floor_levels(self, tolerance=0.01) -> tuple:
        """
        Identify distinct floor elevations from points in the model.
        
        Args:
            tolerance: Coordinate comparison tolerance
            
        Returns:
            tuple: (list of floor elevations sorted from lowest to highest, status code)
                  where status code is 0 for success, 1 for failure
        """
        try:
            z_coords = set()
            
            # Get all points
            num_points, point_names, ret = self._model.PointObj.GetNameList()
            if ret != 0:
                logger.error("Failed to get point names")
                return ([], 1)
                
            # Extract z-coordinates
            for point_name in point_names:
                x, y, z, ret = self._model.PointObj.GetCoordCartesian(point_name)
                if ret == 0 and z > tolerance:  # Ignore ground level (z=0)
                    # Round to tolerance to group similar elevations
                    rounded_z = round(z / tolerance) * tolerance
                    z_coords.add(rounded_z)
            
            return (sorted(list(z_coords)), 0)
        
        except Exception as e:
            logger.error(f"Error in identify_floor_levels: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return ([], 1)

    def add_floor_areas(self, floor_z, tolerance=0.01) -> tuple:
        """
        Add floor areas at the specified elevation by detecting enclosed polygons in the floor structural grid.
        Uses a graph-based approach with face traversal algorithm. Does not add loads.
        
        Args:
            floor_z: Floor elevation to create areas for
            tolerance: Coordinate comparison tolerance
            
        Returns:
            tuple: (list of created area names, status code)
                  where status code is 0 for success, 1 for failure
        """
        try:
            # Get all beams at this floor level
            horizontal_beams = self._get_horizontal_beams_at_elevation(floor_z, tolerance)
            
            if not horizontal_beams:
                logger.warning(f"No horizontal beams found at elevation z={floor_z}")
                return ([], 0)  # Not an error, just no areas created
                
            # Build graph representation of the floor
            vertex_map, adjacency_list = self._build_graph(horizontal_beams, tolerance)
            
            # Sort edges by angle around each vertex
            sorted_neighbors = self._compute_angles_around_vertices(vertex_map, adjacency_list)
            
            # Find all closed polygons (faces)
            faces = self._find_all_faces(vertex_map, sorted_neighbors)
            
            if not faces:
                logger.warning(f"No closed areas detected at elevation z={floor_z}")
                return ([], 0)  # Not an error, just no areas created
                
            # Create area objects in SAP2000 (without loads)
            created_areas = self._create_areas_without_loads(faces, vertex_map, floor_z)
            
            # For each created area, adjust its local axis orientation to align with the shortest edge
            for area_name in created_areas:
                # Get the area's corner points
                ret = self._model.AreaObj.GetPoints(area_name)
                if ret[0] > 0:  # If points were successfully retrieved
                    point_names = ret[1]  # List of point names defining the area
                    
                    # Get coordinates of all area points
                    point_coords = []
                    
                    for point in point_names:
                        x, y, z, ret = self._model.PointObj.GetCoordCartesian(point)
                        if ret == 0:
                            point_coords.append((x, y))
                    
                    # Step 1: Determine shortest edge in the face
                    min_length = float('inf')
                    best_angle = 0.0
                    
                    for j in range(len(point_coords)):
                        x1, y1 = point_coords[j]
                        x2, y2 = point_coords[(j + 1) % len(point_coords)]
                        
                        dx = x2 - x1
                        dy = y2 - y1
                        length = math.sqrt(dx**2 + dy**2)
                        
                        if length < min_length:
                            min_length = length
                            best_angle = math.degrees(math.atan2(dy, dx))
                    
                    # Step 2: Set local axis to align with that shortest edge
                    # Normalize the angle between 0 and 360
                    best_angle = best_angle % 360
                    ret = self._model.AreaObj.SetLocalAxes(area_name, best_angle)
                    if ret != 0:
                        logger.warning(f"Failed to set local axes for area {area_name}")
            
            logger.info(f"Created {len(created_areas)} floor areas at elevation z={floor_z}")
            return (created_areas, 0)
            
        except Exception as e:
            logger.error(f"Error in add_floor_areas: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return ([], 1)
            
    def _get_horizontal_beams_at_elevation(self, elevation, tolerance):
        """
        Get all horizontal beams at the specified elevation.
        
        Args:
            elevation: The Z coordinate to look for beams at
            tolerance: Coordinate comparison tolerance
            
        Returns:
            List of dictionaries containing beam information
        """
        horizontal_beams = []
        
        # Get all frames
        num_frames, frame_names, ret = self._model.FrameObj.GetNameList()
        if ret != 0:
            logger.error("Failed to get frame names")
            return []
        
        # Find horizontal beams at this elevation
        for frame_name in frame_names:
            point_i, point_j, ret = self._model.FrameObj.GetPoints(frame_name)
            if ret != 0:
                continue
                
            # Get coordinates
            x_i, y_i, z_i, ret_i = self._model.PointObj.GetCoordCartesian(point_i)
            x_j, y_j, z_j, ret_j = self._model.PointObj.GetCoordCartesian(point_j)
            
            if ret_i != 0 or ret_j != 0:
                continue
                
            # Check if both points are at the target elevation
            if (abs(z_i - elevation) < tolerance and 
                abs(z_j - elevation) < tolerance):
                horizontal_beams.append({
                    'name': frame_name,
                    'point_i': point_i,
                    'point_j': point_j,
                    'coords_i': (x_i, y_i, z_i),
                    'coords_j': (x_j, y_j, z_j)
                })
        
        return horizontal_beams
        
    def _build_graph(self, beam_list, tolerance):
        """
        Build graph representation of the floor structure.
        
        Args:
            beam_list: List of beam dictionaries with coordinates
            tolerance: Coordinate comparison tolerance
            
        Returns:
            vertex_map: Dictionary mapping coordinates to vertex indices
            adjacency_list: Dictionary mapping vertex indices to neighbors
        """
        vertex_map = {}     # Maps (x, y, z) -> vertex_index
        adjacency_list = {} # Maps vertex_index -> list of (neighbor_index, beam_id)
        vertex_count = 0
        
        for beam in beam_list:
            beam_id = beam['name']
            coords_i = beam['coords_i']
            coords_j = beam['coords_j']
            
            # Convert coordinates to consistent format for comparison
            coords_i_key = tuple(round(v / tolerance) * tolerance for v in coords_i)
            coords_j_key = tuple(round(v / tolerance) * tolerance for v in coords_j)
            
            # Get or create vertex indices
            if coords_i_key not in vertex_map:
                vertex_map[coords_i_key] = vertex_count
                adjacency_list[vertex_count] = []
                vertex_count += 1
                
            if coords_j_key not in vertex_map:
                vertex_map[coords_j_key] = vertex_count
                adjacency_list[vertex_count] = []
                vertex_count += 1
                
            v1 = vertex_map[coords_i_key]
            v2 = vertex_map[coords_j_key]
            
            # Add bidirectional edges
            adjacency_list[v1].append((v2, beam_id))
            adjacency_list[v2].append((v1, beam_id))
        
        return vertex_map, adjacency_list
        
    def _compute_angles_around_vertices(self, vertex_map, adjacency_list):
        """
        Compute and sort edges by angle around each vertex.
        
        Args:
            vertex_map: Dictionary mapping coordinates to vertex indices
            adjacency_list: Dictionary mapping vertex indices to neighbors
            
        Returns:
            Dictionary mapping vertex indices to sorted neighbor lists
        """
        sorted_neighbors = {}
        
        # Invert vertex map for lookup
        vertex_coords = {v: coords for coords, v in vertex_map.items()}
        
        for vertex_index, neighbors in adjacency_list.items():
            v_coord = vertex_coords[vertex_index]
            angle_list = []
            
            for neighbor_index, beam_id in neighbors:
                n_coord = vertex_coords[neighbor_index]
                
                # Compute 2D angle (ignore Z)
                dx = n_coord[0] - v_coord[0]
                dy = n_coord[1] - v_coord[1]
                angle = math.atan2(dy, dx)
                
                angle_list.append((neighbor_index, beam_id, angle))
            
            # Sort by angle
            angle_list.sort(key=lambda x: x[2])
            sorted_neighbors[vertex_index] = angle_list
        
        return sorted_neighbors
        
    def _find_all_faces(self, vertex_map, sorted_neighbors):
        """
        Find all closed polygons using face traversal algorithm.
        Filters out duplicate faces and large perimeter faces.
        
        Args:
            vertex_map: Dictionary mapping coordinates to vertex indices
            sorted_neighbors: Dictionary mapping vertex indices to sorted neighbor lists
            
        Returns:
            List of faces, where each face is a list of vertex indices
        """
        visited_half_edges = set()  # Set of (v_current, v_next) pairs
        all_faces = []
        
        # For each vertex
        for v in sorted_neighbors.keys():
            # Try starting a face from each edge
            for neighbor_data in sorted_neighbors[v]:
                v_next = neighbor_data[0]  # neighbor vertex
                
                # If we haven't visited this half-edge
                if (v, v_next) not in visited_half_edges:
                    face_vertices = self._trace_face(v, v_next, sorted_neighbors, visited_half_edges)
                    
                    if face_vertices and len(face_vertices) >= 3:  # Valid polygon needs at least 3 vertices
                        all_faces.append(face_vertices)
        
        # Filter out duplicate faces (same vertices in different order)
        unique_faces = []
        unique_face_sets = set()
        
        for face in all_faces:
            # Convert face to a frozenset for set comparison
            face_set = frozenset(face)
            
            # Only add if we haven't seen this exact set of vertices
            if face_set not in unique_face_sets:
                unique_face_sets.add(face_set)
                
                # For grid structures, we're primarily interested in quadrilaterals
                # This also filters out large perimeter faces
                if len(face) == 4:
                    unique_faces.append(face)
        
        logger.info(f"Found {len(all_faces)} total faces, filtered to {len(unique_faces)} unique quadrilateral faces")
        return unique_faces
        
    def _trace_face(self, start_v, next_v, sorted_neighbors, visited_half_edges):
        """
        Trace a single face by following half-edges.
        
        Args:
            start_v: Starting vertex index
            next_v: Next vertex index
            sorted_neighbors: Dictionary mapping vertex indices to sorted neighbor lists
            visited_half_edges: Set of visited half-edges
            
        Returns:
            List of vertex indices forming a face, or empty list if no face found
        """
        face_vertices = [start_v]
        current_v = start_v
        next_v_candidate = next_v
        
        # Maximum iterations as a safety measure
        max_iterations = 1000
        iterations = 0
        
        while iterations < max_iterations:
            iterations += 1
            
            # Mark half-edge as visited
            visited_half_edges.add((current_v, next_v_candidate))
            
            # Move to next vertex
            current_v = next_v_candidate
            face_vertices.append(current_v)
            
            # Find the next half-edge - get the index of the edge that points back to previous vertex
            neighbors = sorted_neighbors[current_v]
            back_index = -1
            
            for i, (neighbor, _, _) in enumerate(neighbors):
                if neighbor == face_vertices[-2]:  # Previous vertex
                    back_index = i
                    break
                    
            if back_index == -1:
                # Error: can't find the return edge
                return []
                
            # Next edge in counter-clockwise order
            next_index = (back_index + 1) % len(neighbors)
            next_v_candidate = neighbors[next_index][0]
            
            # If we've returned to the start, we've completed the face
            if next_v_candidate == start_v:
                return face_vertices
        
        # If we get here, we didn't complete the face within max iterations
        logger.warning(f"Face trace exceeded {max_iterations} iterations without completion")
        return []
        
    def _create_areas_without_loads(self, faces, vertex_map, floor_z):
        """
        Create area objects in SAP2000 (without loads).
        
        Args:
            faces: List of faces, where each face is a list of vertex indices
            vertex_map: Dictionary mapping coordinates to vertex indices
            floor_z: Floor elevation
            
        Returns:
            List of created area names
        """
        created_areas = []
        
        # Invert vertex map for lookup
        vertex_coords = {v: coords for coords, v in vertex_map.items()}
        
        for i, face in enumerate(faces):
            # Extract coordinates for each vertex (retaining only X and Y, using floor_z for Z)
            x_array = []
            y_array = []
            z_array = []
            
            for vertex in face:
                coords = vertex_coords[vertex]
                x_array.append(coords[0])
                y_array.append(coords[1])
                z_array.append(floor_z)  # Use consistent floor Z value
            
            # Skip areas that are too small (likely noise or error)
            if len(x_array) < 3:
                continue
                
            # Create the area
            ret = self._model.AreaObj.AddByCoord(
                len(x_array),  # Number of points
                x_array,       # X coordinates 
                y_array,       # Y coordinates
                z_array,       # Z coordinates (all equal to floor_z)
                ""             # Auto-name
            )
            
            # Get the generated area name
            area_name = ret[3] if len(ret) > 3 else f"Area_{i+1}"
            created_areas.append(area_name)
            
            # Determine the shorter span direction and set local axis accordingly
            # Calculate the dimensions of the area in X and Y directions
            min_x, max_x = min(x_array), max(x_array)
            min_y, max_y = min(y_array), max(y_array)
            x_span = max_x - min_x
            y_span = max_y - min_y
            
            # If the X span is shorter, rotate the local axis
            if x_span < y_span:
                # Set local axis angle to 90 degrees (align with shorter span)
                self._model.AreaObj.SetLocalAxes(area_name, 90)
            else:
                # Set local axis angle to 0 degrees (default, align with shorter span)
                self._model.AreaObj.SetLocalAxes(area_name, 0)
            
        return created_areas

    def get_beams_info(self, tolerance: float = 1.0) -> Dict[float, List[str]]:
        """
        Groups all beams in the model by their approximate length.

        Args:
            tolerance: Tolerance for length matching (default 1.0 ft)
            
        Returns:
            Dictionary mapping lengths to lists of beam frame names
            e.g. {24.0: ['beam1', 'beam2', ...], 10.0: ['beam3', ...]}
        """
        try:
            # Get all frame objects
            number_frames, frame_names, ret = self._model.FrameObj.GetNameList()
            if ret != 0 or not frame_names:
                logger.error("Failed to get frame names from model")
                return {}
                
            # Dictionary to store beams grouped by length
            beams_by_length = {}
            
            # Process each frame
            for frame in frame_names:
                # Get frame endpoints
                point_i, point_j, ret = self._model.FrameObj.GetPoints(frame)
                if ret != 0:
                    continue
                    
                # Get coordinates
                x_i, y_i, z_i, ret_i = self._model.PointObj.GetCoordCartesian(point_i)
                x_j, y_j, z_j, ret_j = self._model.PointObj.GetCoordCartesian(point_j)
                
                if ret_i != 0 or ret_j != 0:
                    continue
                
                # Determine if frame is beam (horizontal)
                is_beam = abs(z_i - z_j) <= 0.1
                
                if is_beam:
                    # Calculate beam length
                    dx = x_j - x_i
                    dy = y_j - y_i
                    frame_length = (dx**2 + dy**2)**0.5
                    
                    # Round length to the nearest foot (or tolerance level)
                    rounded_length = round(frame_length / tolerance) * tolerance
                    
                    # Add to the dictionary
                    if rounded_length not in beams_by_length:
                        beams_by_length[rounded_length] = []
                        
                    beams_by_length[rounded_length].append(frame)
            
            # Log summary
            beam_counts = {length: len(beams) for length, beams in beams_by_length.items()}
            logger.info(f"Beam length distribution: {beam_counts}")
            
            return beams_by_length
            
        except Exception as e:
            logger.error(f"Error in get_beams_info: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return {}

    def get_columns_info(self, tolerance: float = 1.0) -> Dict[str, List[str]]:
        """
        Groups all columns in the model by their location (corner, edge, interior).

        Args:
            tolerance: Tolerance for coordinate comparison (default 1.0 ft)
            
        Returns:
            Dictionary mapping location types to lists of column frame names
            e.g. {'corner': ['column1', ...], 'edge': ['column2', ...], 'interior': ['column3', ...]}
        """
        try:
            # Get all frame objects
            number_frames, frame_names, ret = self._model.FrameObj.GetNameList()
            if ret != 0 or not frame_names:
                logger.error("Failed to get frame names from model")
                return {}
                
            # Dictionary to store columns grouped by location
            columns_by_location = {
                'corner': [],
                'edge': [],
                'interior': []
            }
            
            # Get all points in the model to determine building bounds
            number_points, point_names, ret = self._model.PointObj.GetNameList()
            if ret != 0 or not point_names:
                logger.error("Failed to get point names from model")
                return {}
            
            # Collect all X and Y coordinates
            x_coords = []
            y_coords = []
            for point in point_names:
                x, y, z, ret = self._model.PointObj.GetCoordCartesian(point)
                if ret == 0:
                    x_coords.append(x)
                    y_coords.append(y)
            
            # If no points found, return empty
            if not x_coords or not y_coords:
                logger.error("Failed to determine building bounds - no valid points found")
                return {}
                
            # Calculate actual bounds
            x_min = min(x_coords)
            x_max = max(x_coords)
            y_min = min(y_coords)
            y_max = max(y_coords)
            
            logger.info(f"Building bounds: X=[{x_min}, {x_max}], Y=[{y_min}, {y_max}]")
                
            # Process each frame
            for frame in frame_names:
                # Get frame endpoints
                point_i, point_j, ret = self._model.FrameObj.GetPoints(frame)
                if ret != 0:
                    continue
                    
                # Get coordinates
                x_i, y_i, z_i, ret_i = self._model.PointObj.GetCoordCartesian(point_i)
                x_j, y_j, z_j, ret_j = self._model.PointObj.GetCoordCartesian(point_j)
                
                if ret_i != 0 or ret_j != 0:
                    continue
                
                # Determine if frame is column (vertical)
                is_column = abs(z_i - z_j) > 0.1
                
                if is_column:
                    # Use lower Z point for column position check
                    x, y = (x_i, y_i) if z_i <= z_j else (x_j, y_j)
                    
                    # Check column position using calculated bounds
                    is_corner = (abs(x - x_min) < tolerance or abs(x - x_max) < tolerance) and \
                                (abs(y - y_min) < tolerance or abs(y - y_max) < tolerance)
                    
                    is_edge = (abs(x - x_min) < tolerance or abs(x - x_max) < tolerance or \
                              abs(y - y_min) < tolerance or abs(y - y_max) < tolerance)
                    
                    # Classify column based on position
                    if is_corner:
                        columns_by_location['corner'].append(frame)
                    elif is_edge:
                        columns_by_location['edge'].append(frame)
                    else:
                        columns_by_location['interior'].append(frame)
            
            # Log summary
            column_counts = {loc: len(cols) for loc, cols in columns_by_location.items()}
            logger.info(f"Column location distribution: {column_counts}")
            
            return columns_by_location
            
        except Exception as e:
            logger.error(f"Error in get_columns_info: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return {'corner': [], 'edge': [], 'interior': []}

    def create_assign_section_group(
        self, 
        group_name: str, 
        frames: List[str]
    ) -> Tuple[List[str], int]:
        """
        Creates a group and assigns the specified frames to it.
        
        Args:
            group_name: Name of the group to create
            frames: List of frame names to include in the group
            
        Returns:
            tuple: (list of frame names in the group, status code)
            where status code is 0 for success, 1 for failure
        """
        try:
            # Create the group if it doesn't exist
            ret = self._model.GroupDef.SetGroup(group_name)
            if ret != 0:
                logger.warning(f"Failed to create group {group_name}")
                return ([], 1)
                
            # Track successfully added frames
            added_frames = []
            
            # Add frames to the group
            for frame in frames:
                ret = self._model.FrameObj.SetGroupAssign(frame, group_name)
                if ret != 0:
                    logger.warning(f"Failed to assign frame {frame} to group {group_name}")
                else:
                    added_frames.append(frame)
            
            logger.info(f"Successfully created group {group_name} with {len(added_frames)} frames")
            return (added_frames, 0)
            
        except Exception as e:
            logger.error(f"Error in create_assign_section_group: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return ([], 1)

    def analyze_group_utilization(self, group_name: str, frame_type: str = "Beam") -> None:
        """
        Analyzes the utilization ratios for frames in a group and logs useful statistics.
        Identifies underutilized frames that could potentially be regrouped.
        
        Args:
            group_name: Name of the frame group to analyze
            frame_type: Type of frame ("Beam" or "Column") for logging purposes
            
        Returns:
            List of underutilized frames with their utilization ratios
        """
        try:
            # Get utilization ratios for the group
            # The GetSummaryResults API returns a tuple with the following structure:
            # (number_items, frame_names, ratios, ratio_types, locations, combo_names, error_summary, warning_summary, status)
            ret = self._model.DesignSteel.GetSummaryResults(
                group_name,
                0,  # number_items (will be populated in results)
                [],  # frame_names (will be populated in results)
                [],  # ratios (will be populated in results)
                [],  # ratio_types (will be populated in results)
                [],  # locations (will be populated in results)
                [],  # combo_names (will be populated in results)
                [],  # error_summary (will be populated in results)
                [],  # warning_summary (will be populated in results)
                1   # ItemType = Group
            )
            
            # Parse the return value - last element (index 8) is the status code
            status_code = ret[8]
            if status_code != 0:  # Status code (0 = success)
                logger.warning(f"No utilization results found for group: {group_name}, error code: {status_code}")
                return []
                
            # Extract data from the returned tuple
            number_items = ret[0]  # First element is the number of items
            frame_names = ret[1]
            ratios = ret[2]
            ratio_types = ret[3]
            locations = ret[4]
            combo_names = ret[5]
            error_summary = ret[6]
            warning_summary = ret[7]
            
            # Check if we have valid data
            if not frame_names or not ratios:
                logger.warning(f"No utilization results found for group: {group_name}")
                return []
                
            # Log any errors or warnings
            for i, error in enumerate(error_summary):
                if error and error.strip():
                    logger.error(f"Design error for {frame_names[i]}: {error}")
            
            for i, warning in enumerate(warning_summary):
                if warning and warning.strip():
                    logger.warning(f"Design warning for {frame_names[i]}: {warning}")
            
            # Calculate statistics
            avg_ratio = sum(ratios) / len(ratios) if ratios else 0
            max_ratio = max(ratios) if ratios else 0
            min_ratio = min(ratios) if ratios else 0
            
            # Get index of max utilized frame
            max_index = ratios.index(max_ratio) if ratios else -1
            
            # Log group summary
            logger.info(f"{group_name} Utilization Summary:")
            logger.info(f"  Total Members: {number_items}")
            logger.info(f"  Average Utilization: {avg_ratio:.3f}")
            logger.info(f"  Maximum Utilization: {max_ratio:.3f}")
            logger.info(f"  Minimum Utilization: {min_ratio:.3f}")
            
            # Log most utilized frame details
            if max_index >= 0:
                max_frame = frame_names[max_index]
                max_ratio_type = ratio_types[max_index]
                max_combo = combo_names[max_index]
                max_location = locations[max_index]
                
                # Get ratio type description
                ratio_type_desc = "Unknown"
                if max_ratio_type == 1:
                    ratio_type_desc = "PMM"
                elif max_ratio_type == 2:
                    ratio_type_desc = "Major shear"
                elif max_ratio_type == 3:
                    ratio_type_desc = "Minor shear"
                elif max_ratio_type == 4:
                    ratio_type_desc = "Major beam-column capacity"
                elif max_ratio_type == 5:
                    ratio_type_desc = "Minor beam-column capacity"
                elif max_ratio_type == 6:
                    ratio_type_desc = "Other"
                
                # Get section information
                section_name, auto_list, _ = self._model.FrameObj.GetSection(max_frame)
                
                logger.info(f"  Most Utilized {frame_type}:")
                logger.info(f"    Frame: {max_frame}")
                logger.info(f"    Section: {section_name}")
                logger.info(f"    Utilization: {max_ratio:.3f}")
                logger.info(f"    Controlling Type: {ratio_type_desc}")
                logger.info(f"    Location: {max_location} ft from start")
                logger.info(f"    Combo: {max_combo}")
            
            # Find and log underutilized frames (under 50% utilization)
            underutilized = [(frame_names[i], ratios[i]) for i in range(len(ratios)) if ratios[i] < 0.5]
            
            if underutilized:
                logger.info(f"  Underutilized {frame_type}s (below 50%):")
                for frame, ratio in sorted(underutilized, key=lambda x: x[1]):
                    section, _, _ = self._model.FrameObj.GetSection(frame)
                    logger.info(f"    {frame}: {ratio:.3f} (Section: {section})")
                    
                # Recommendation for potential regrouping
                if len(underutilized) > 2:
                    logger.info(f"  Recommendation: Consider regrouping {len(underutilized)} {frame_type.lower()}s with low utilization")
            
            return underutilized
        
        except Exception as e:
            logger.error(f"Error in analyze_group_utilization: {str(e)}")
            logger.error(f"Stack trace: {traceback.format_exc()}")
            return []

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
        beam_auto_select_lists = {
            "24ft Beams": ["W24X76", "W24X84", "W24X94", "W24X103", "W24X117"],
            "22ft Beams": ["W21X44", "W21X50", "W21X57", "W21X68", "W21X83"], 
            "18ft Beams": ["W18X40", "W18X46", "W18X50", "W18X55", "W18X60"],
            "14ft Beams": ["W14X34", "W14X38", "W14X43", "W14X48", "W14X53"],
            "10ft Beams": ["W10X33", "W10X39", "W10X45", "W10X49", "W10X54"]
        }
    
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
        column_auto_select_lists = {
            "corner": ["W10X12", "W10X15", "W10X19", "W10X22", "W10X26"],
            "edge": ["W12X190", "W12X210", "W12X230", "W12X252", "W12X279"],
            "interior": ["W14X193", "W14X211", "W14X233", "W14X257", "W14X283"]
        }
        
        # Step 1: Import all column section properties
        for location, section_list in column_auto_select_lists.items():
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
            section_list = column_auto_select_lists[location]
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
