"""
Revit Dynamo Script for Creating Exterior and Interior Walls

This script automates the process of creating exterior and interior walls in Revit based on user-selected lines or curves.
The user selects wall types and lines/curves in Dynamo, and the script generates the walls at specified heights.

Inputs:
IN[0] (str): Name of the wall type for exterior walls selected from a dropdown in Dynamo.
IN[1] (list): List of model elements (lines/curves) for exterior walls selected by the user.
IN[2] (str): Name of the wall type for interior walls selected from a dropdown in Dynamo.
IN[3] (list): List of model elements (lines/curves) for interior walls selected by the user.

Adjustable Parameters:
DEFAULT_EXTERIOR_WALL_HEIGHT (float): Default height for exterior walls in feet.
DEFAULT_INTERIOR_WALL_HEIGHT (float): Default height for interior walls in feet.

Outputs:
OUT (list): List of created wall elements.

Notes:
- Ensure that the wall types and model elements (lines/curves) are correctly selected in Dynamo before running the script.
- The script includes logging for debugging purposes. Check the 'revit_script_delete_this.log' file for detailed logs.
"""
import clr
import os
import logging
from datetime import datetime
from Autodesk.Revit.DB import (
    Level, WallType, Wall, FilteredElementCollector,
    CurveElement, Line, Arc, PolyLine, Ellipse, JoinGeometryUtils,
    UnitUtils, UnitTypeId, Transaction, XYZ, SetComparisonResult
)
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Revit.Elements import *
from Revit.GeometryConversion import *
import tempfile
from logging.handlers import RotatingFileHandler


clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')
clr.AddReference('RevitAPIUI')


# Setup logging
def setup_logging():
    try:
        log_directory = tempfile.gettempdir()
        log_filename = datetime.now().strftime('%I%p_%m.%d.%Y_revit_script_delete_this.log')
        log_path = os.path.join(log_directory, log_filename)
        handler = RotatingFileHandler(log_path, maxBytes=5*1024*1024, backupCount=5)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.basicConfig(level=logging.DEBUG, handlers=[handler])
        logger = logging.getLogger()
        logger.info("Logging setup complete.")
        return logger
    except Exception as e:
        print("Logging setup failed: {}".format(e))
        return None


logger = setup_logging()


def log_message(level, message):
    """Logs a message at the specified level."""
    if logger:
        if level == 'info':
            logger.info(message)
        elif level == 'debug':
            logger.debug(message)
        elif level == 'warning':
            logger.warning(message)
        elif level == 'error':
            logger.error(message)
    print("{}: {}".format(level.upper(), message))


doc = DocumentManager.Instance.CurrentDBDocument

DEFAULT_EXTERIOR_WALL_HEIGHT = 25.0  # Default height for exterior walls in feet
DEFAULT_INTERIOR_WALL_HEIGHT = 12.0  # Default height for interior walls in feet


def handle_exceptions(func):
    """Decorator function to handle exceptions and log errors."""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as ex:
            error_message = "Exception occurred in {}: {}".format(func.__name__, str(ex))
            log_message('error', error_message)
            raise
    return wrapper


@handle_exceptions
def validate_inputs():
    """Validate user inputs."""
    log_message('debug', "IN[0] (Exterior wall type name): {}".format(IN[0]))
    log_message('debug', "IN[1] (Exterior wall elements): {}".format(IN[1]))
    log_message('debug', "IN[2] (Interior wall type name): {}".format(IN[2]))
    log_message('debug', "IN[3] (Interior wall elements): {}".format(IN[3]))
    if not isinstance(IN[0], WallType):
        raise ValueError("Exterior wall type must be a WallType element.")
    if not isinstance(IN[1], list) or not all(isinstance(item, CurveElement) for item in IN[1]):
        raise ValueError("Exterior wall elements must be a list of CurveElements.")
    if not isinstance(IN[2], WallType):
        raise ValueError("Interior wall type must be a WallType element.")
    if not isinstance(IN[3], list) or not all(isinstance(item, CurveElement) for item in IN[3]):
        raise ValueError("Interior wall elements must be a list of CurveElements.")
    log_message('info', "Input validation passed.")


@handle_exceptions
def get_first_level():
    """Retrieve the first level based on elevation in the document."""
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    log_message('debug', "Found levels: {}".format([level.Name for level in levels]))
    if not levels:
        raise ValueError("No levels found in the document.")
    sorted_levels = sorted(levels, key=lambda x: x.Elevation)
    first_level = sorted_levels[0]
    log_message('info', "Using first level: {}, Elevation: {}".format(first_level.Name, first_level.Elevation))
    return first_level


@handle_exceptions
def convert_to_internal_units(input_length, unit_type_id=UnitTypeId.Feet):
    """Convert a given length to Revit internal units."""
    internal_units = UnitUtils.ConvertToInternalUnits(input_length, unit_type_id)
    log_message('debug', "Converted {} to internal units: {}".format(input_length, internal_units))
    return internal_units


@handle_exceptions
def validate_curve_elements(curve_elements):
    """Validate that the curve elements are of a supported type and are planar."""
    valid_curves = []
    for curve in curve_elements:
        if isinstance(curve, CurveElement):
            geometry_curve = curve.GeometryCurve
            if isinstance(geometry_curve, (Line, Arc, PolyLine, Ellipse)):
                valid_curves.append(curve)
                log_message('debug', "Valid curve: {}".format(curve.Id))
            else:
                log_message('warning', "Invalid or non-planar curve type: {}".format(type(geometry_curve)))
        else:
            log_message('warning', "Invalid curve element: {}".format(curve))
    log_message('debug', "Validated curves: {} valid curves found.".format(len(valid_curves)))
    return valid_curves


@handle_exceptions
def adjust_curve_elevation(curves, target_elevation=0):
    """Adjust the elevation of curves to the target elevation."""
    adjusted_curves = []
    for curve in curves:
        geometry_curve = curve.GeometryCurve
        start_point = geometry_curve.GetEndPoint(0)
        end_point = geometry_curve.GetEndPoint(1)
        new_start_point = XYZ(start_point.X, start_point.Y, target_elevation)
        new_end_point = XYZ(end_point.X, end_point.Y, target_elevation)
        if isinstance(geometry_curve, Line):
            new_curve = Line.CreateBound(new_start_point, new_end_point)
        elif isinstance(geometry_curve, Arc):
            mid_point = geometry_curve.Evaluate(0.5, True)
            new_mid_point = XYZ(mid_point.X, mid_point.Y, target_elevation)
            new_curve = Arc.Create(new_start_point, new_end_point, new_mid_point)
        elif isinstance(geometry_curve, PolyLine):
            points = geometry_curve.GetCoordinates()
            new_points = [XYZ(point.X, point.Y, target_elevation) for point in points]
            new_curve = PolyLine.Create(new_points)
        else:
            log_message('warning', "Unsupported curve type for elevation adjustment: {}".format(type(geometry_curve)))
            continue
        adjusted_curves.append(new_curve)
    log_message('debug', "Adjusted curves: {} curves adjusted to elevation {}.".format(len(adjusted_curves), target_elevation))
    return adjusted_curves


@handle_exceptions
def filter_overlapping_curves(curves):
    """Filter overlapping curves, keeping the larger one if overlaps are detected."""
    filtered_curves = []
    for curve in curves:
        overlap_found = False
        for other_curve in filtered_curves:
            if curve.GeometryCurve.Intersect(other_curve.GeometryCurve) != SetComparisonResult.Disjoint:
                overlap_found = True
                if curve.GeometryCurve.Length > other_curve.GeometryCurve.Length:
                    filtered_curves.remove(other_curve)
        if not overlap_found:
            filtered_curves.append(curve)
    log_message('debug', "Filtered curves: {} curves after removing overlaps.".format(len(filtered_curves)))
    return filtered_curves


@handle_exceptions
def create_wall(curve, wall_type_id, level_id, wall_height):
    """Create a wall in Revit based on the provided curve."""
    with Transaction(doc, "Create Wall") as trans:
        trans.Start()
        wall = Wall.Create(doc, curve, wall_type_id, level_id, wall_height, 0, False, False)
        trans.Commit()
        log_message('info', "Wall created with ID: {} from curve: {}".format(wall.Id, curve.Id))
        return wall


@handle_exceptions
def process_curves(curves, wall_type_id, level_id, wall_height, wall_type):
    """Process the provided curves to create walls in Revit."""
    created_walls = []

    with Transaction(doc, "Create Walls from Curves") as trans:
        trans.Start()
        for curve in curves:
            geometry_curve = curve.GeometryCurve
            wall = Wall.Create(doc, geometry_curve, wall_type_id, level_id, wall_height, 0, False, False)
            created_walls.append(wall)
            log_message('info', "Created {} wall with ID: {} from curve: {}".format(wall_type, wall.Id, curve.Id))
        trans.Commit()
        log_message('info', "Created {} {} walls.".format(len(created_walls), wall_type))

    return created_walls


@handle_exceptions
def clean_up_walls(walls):
    """Clean up walls by joining geometry and removing unnecessary walls."""
    with Transaction(doc, "Clean Up Walls") as trans:
        trans.Start()
        for i in range(len(walls)):
            for j in range(i + 1, len(walls)):
                try:
                    JoinGeometryUtils.JoinGeometry(doc, walls[i], walls[j])
                    log_message('debug', "Joined walls: {} and {}".format(walls[i].Id, walls[j].Id))
                except Exception as e:
                    log_message('warning', "Could not join walls {} and {}: {}".format(walls[i].Id, walls[j].Id, e))

        for wall in walls:
            if wall.Location.Curve.Length < 1.0:
                doc.Delete(wall.Id)
                log_message('info', "Removed short wall: {}".format(wall.Id))

        trans.Commit()
        log_message('info', "Wall cleanup completed successfully.")


@handle_exceptions
def perform_quality_assurance_checks(walls):
    """Perform quality assurance checks on created walls."""
    try:
        for wall in walls:
            if not isinstance(wall, Wall):
                log_message('error', "Invalid wall element found: {}".format(wall.Id))
    except Exception as e:
        log_message('error', "Error during quality assurance checks: {}".format(e))


@handle_exceptions
def main():
    """Main function to execute the script logic."""
    results = []
    log_message('info', "Starting wall creation process.")
    validate_inputs()
    first_level = get_first_level()
    first_level_id = first_level.Id
    log_message('info', "Using first level: {}".format(first_level.Name))

    wall_types = FilteredElementCollector(doc).OfClass(WallType).ToElements()

    selected_exterior_wall_type_name = IN[0].Name
    selected_exterior_wall_type = next((wt for wt in wall_types if wt.Name == selected_exterior_wall_type_name), None)
    if not selected_exterior_wall_type:
        raise ValueError("Selected exterior wall type '{}' not found.".format(selected_exterior_wall_type_name))
    exterior_wall_type_id = selected_exterior_wall_type.Id
    log_message('info', "Selected exterior wall type: {}".format(selected_exterior_wall_type_name))

    selected_interior_wall_type_name = IN[2].Name
    selected_interior_wall_type = next((wt for wt in wall_types if wt.Name == selected_interior_wall_type_name), None)
    if not selected_interior_wall_type:
        raise ValueError("Selected interior wall type '{}' not found.".format(selected_interior_wall_type_name))
    interior_wall_type_id = selected_interior_wall_type.Id
    log_message('info', "Selected interior wall type: {}".format(selected_interior_wall_type_name))

    exterior_lines = UnwrapElement(IN[1])
    interior_lines = UnwrapElement(IN[3])
    log_message('debug', "Selected exterior lines: {}".format(exterior_lines))
    log_message('debug', "Selected interior lines: {}".format(interior_lines))

    exterior_lines = validate_curve_elements(exterior_lines)
    interior_lines = validate_curve_elements(interior_lines)
    if not exterior_lines:
        raise ValueError("No valid exterior lines provided.")
    if not interior_lines:
        raise ValueError("No valid interior lines provided.")
    log_message('info', "Valid exterior lines: {}".format(len(exterior_lines)))
    log_message('info', "Valid interior lines: {}".format(len(interior_lines)))

    exterior_wall_height = convert_to_internal_units(DEFAULT_EXTERIOR_WALL_HEIGHT, UnitTypeId.Feet)
    interior_wall_height = convert_to_internal_units(DEFAULT_INTERIOR_WALL_HEIGHT, UnitTypeId.Feet)
    log_message('info', "Exterior wall height (internal units): {}".format(exterior_wall_height))
    log_message('info', "Interior wall height (internal units): {}".format(interior_wall_height))

    exterior_lines = adjust_curve_elevation(exterior_lines, first_level.Elevation)
    interior_lines = adjust_curve_elevation(interior_lines, first_level.Elevation)
    log_message('debug', "Adjusted curve elevations to match the first level.")

    exterior_lines = filter_overlapping_curves(exterior_lines)
    interior_lines = filter_overlapping_curves(interior_lines)
    log_message('debug', "Filtered overlapping curves.")

    TransactionManager.Instance.EnsureInTransaction(doc)

    exterior_results = process_curves(exterior_lines, exterior_wall_type_id, first_level_id, exterior_wall_height, 'exterior')
    results.extend(exterior_results)

    interior_results = process_curves(interior_lines, interior_wall_type_id, first_level_id, interior_wall_height, 'interior')
    results.extend(interior_results)

    log_message('info', "Total walls created: {}, Exterior walls: {}, Interior walls: {}".format(len(results), len(exterior_results), len(interior_results)))

    clean_up_walls(results)

    perform_quality_assurance_checks(results)

    TransactionManager.Instance.TransactionTaskDone()
    log_message('info', "Transaction completed successfully.")


    return results


# Execute the main function
OUT = main()
