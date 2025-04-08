import logging
from typing import Dict, List
from collections import defaultdict

logger = logging.getLogger(__name__)

class FrameMethods:
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
            return {'corner': [], 'edge': [], 'interior': []}

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

    def define_section_candidate(self, frames, section_type='w', member_type='beam'):
        """
        Automatically defines section candidates based on frame lengths and member type.
        
        Args:
            frames (list): List of frame names to analyze
            section_type (str): Type of section ('w' for wide flange)
            member_type (str): Type of member ('beam' or 'column')
        
        Returns:
            list: List of section names
        """
        if member_type.lower() == 'column':
            # For columns, return all W sections sorted by depth
            w_sections = self.section_names['W']
            return sorted(w_sections, 
                        key=lambda x: float(x.split('X')[0][1:]),  # Sort by depth
                        reverse=True)  # Largest to smallest
        
        # For beams, calculate average length and find appropriate sections
        total_length = 0
        for frame in frames:
            # Get frame endpoints
            point_i, point_j, ret = self._model.FrameObj.GetPoints(frame)
            if ret != 0:
                continue
                
            # Get coordinates
            x_i, y_i, z_i, ret_i = self._model.PointObj.GetCoordCartesian(point_i)
            x_j, y_j, z_j, ret_j = self._model.PointObj.GetCoordCartesian(point_j)
            
            if ret_i != 0 or ret_j != 0:
                continue
            
            # Calculate length
            dx = x_j - x_i
            dy = y_j - y_i
            frame_length = (dx**2 + dy**2)**0.5
            total_length += frame_length
        
        if not frames:
            return []
            
        avg_length_ft = total_length / len(frames)  # Length is already in feet
        
        # Get all W sections
        w_sections = self.section_names['W']
        
        # Group sections by depth
        sections_by_depth = defaultdict(list)
        for section_name in w_sections:
            depth = int(section_name.split('X')[0][1:])  # Extract depth from name (e.g., W24X76 -> 24)
            sections_by_depth[depth].append(section_name)
        
        # Find the closest depth based on span length
        # Rule of thumb: depth (in inches) ≈ span (in feet)
        target_depth = round(avg_length_ft)
        
        # Find the closest available depth
        available_depths = sorted(sections_by_depth.keys())
        closest_depth = min(available_depths, key=lambda x: abs(x - target_depth))
        
        # Return sections of the closest depth, sorted by weight
        return sorted(sections_by_depth[closest_depth],
                    key=lambda x: float(x.split('X')[1]))  # Sort by weight 