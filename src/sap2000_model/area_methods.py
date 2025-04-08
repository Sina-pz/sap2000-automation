import logging
import math
from typing import List, Tuple, Dict
from collections import defaultdict

logger = logging.getLogger(__name__)

class AreaMethods:
    def identify_floor_levels(self, tolerance=0.01) -> Tuple[List[float], int]:
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
            return ([], 1)

    def add_floor_areas(self, floor_z, tolerance=0.01) -> Tuple[List[str], int]:
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
            return ([], 1) 

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