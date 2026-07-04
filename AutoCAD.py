import logging
import win32com.client
import pythoncom
from enum import Enum

logger = logging.getLogger(__name__)

# Enum to represent common colors in AutoCAD
class Color(Enum):
    RED = 1
    YELLOW = 2
    GREEN = 3
    CYAN = 4
    BLUE = 5
    MAGENTA = 6
    WHITE = 7
    GRAY = 8
    ORANGE = 30
    PURPLE = 40
    BROWN = 41

    @staticmethod
    def from_name(name):
        try:
            return Color[name.upper()].value
        except KeyError:
            raise ValueError(f"Color '{name}' is not a valid color name")

# Enum to represent alignments in AutoCAD
class Alignment(Enum):
    LEFT = 'left'
    CENTER = 'center'
    RIGHT = 'right'

# Enum to represent dimension types in AutoCAD
class DimensionType(Enum):
    ALIGNED = 'aligned'
    LINEAR = 'linear'
    ANGULAR = 'angular'
    RADIAL = 'radial'
    DIAMETER = 'diameter'

# Enum to represent line styles in AutoCAD
class LineStyle(Enum):
    CONTINUOUS = 'Continuous'        # ------------
    DASHED = 'Dashed'                # - - - - - - 
    DOTTED = 'Dotted'                # . . . . . .
    CENTER = 'Center'                # - . - . - .
    HIDDEN = 'Hidden'                # - - - - - - 
    PHANTOM = 'Phantom'              # - . . - . .
    BREAK = 'Break'                  # -     -     
    BORDER = 'Border'                # - - - . - -
    DOT2 = 'Dot2'                    # .  .  .  . 
    DOTX2 = 'DotX2'                  # .   .   .  
    DIVIDE = 'Divide'                # -  .  -  .
    TRACKING = 'Tracking'            # - .  - .  
    DASHDOT = 'Dashdot'              # - . - . - 

# Class to represent a 3D point
class APoint:
    def __init__(self, x=0, y=0, z=0):
        self.x = x
        self.y = y
        self.z = z

    def to_variant(self):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [self.x, self.y, self.z])

    def to_tuple(self):
        return (self.x, self.y)

    def __repr__(self):
        return f"APoint({self.x}, {self.y}, {self.z})"

# Class to represent a layer
class Layer:
    def __init__(self, name, color=Color.WHITE, visible=True):
        self.name = name
        self.color = color
        self.visible = visible

    def __repr__(self):
        return f"Layer(name='{self.name}', color={self.color}, visible={self.visible})"

# Class to represent a block reference
class BlockReference:
    def __init__(self, name, insertion_point, scale=1.0, rotation=0.0):
        self.name = name
        self.insertion_point = insertion_point
        self.scale = scale
        self.rotation = rotation

    def __repr__(self):
        return f"BlockReference(name='{self.name}', insertion_point={self.insertion_point}, scale={self.scale}, rotation={self.rotation})"

# Class to represent a text object
class Text:
    def __init__(self, content, insertion_point, height, alignment=Alignment.LEFT):
        self.content = content
        self.insertion_point = insertion_point
        self.height = height
        self.alignment = alignment

    def __repr__(self):
        return f"Text(content='{self.content}', insertion_point={self.insertion_point}, height={self.height}, alignment='{self.alignment}')"

# Class to represent a dimension
class Dimension:
    def __init__(self, start_point, end_point, text_point=None, dimension_type=DimensionType.ALIGNED,
                 rotation_angle=0.0, vertex_point=None, leader_length=None):
        self.start_point = start_point
        self.end_point = end_point
        self.text_point = text_point
        self.dimension_type = dimension_type
        self.rotation_angle = rotation_angle
        # Angle vertex, required by ANGULAR dimensions
        self.vertex_point = vertex_point
        # Leader length, required by RADIAL and DIAMETER dimensions
        self.leader_length = leader_length

    def __repr__(self):
        return (f"Dimension(start_point={self.start_point}, end_point={self.end_point}, "
                f"text_point={self.text_point}, dimension_type='{self.dimension_type}', "
                f"rotation_angle={self.rotation_angle}, vertex_point={self.vertex_point}, "
                f"leader_length={self.leader_length})")

# Custom class for handling AutoCAD errors
class AutoCADError(Exception):
    def __init__(self, message):
        super().__init__(message)
        logger.error(message)

# Main class for interacting with AutoCAD
class AutoCAD:
    def __init__(self):
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.acad.Visible = True
            self.doc = self.acad.ActiveDocument
            self.modelspace = self.doc.ModelSpace
        except Exception as e:
            raise AutoCADError(f"Error initializing AutoCAD: {e}") from e

    # Iterate over objects in the model space, optionally filtering by object type
    def iter_objects(self, object_type=None):
        for obj in self.modelspace:
            if object_type is None or obj.EntityName == object_type:
                yield obj

    # Add a circle to the model space
    def add_circle(self, center, radius):
        try:
            circle = self.modelspace.AddCircle(center.to_variant(), radius)
            return circle
        except Exception as e:
            raise AutoCADError(f"Error adding circle: {e}") from e

    # Add a line to the model space
    def add_line(self, start_point, end_point):
        try:
            line = self.modelspace.AddLine(start_point.to_variant(), end_point.to_variant())
            return line
        except Exception as e:
            raise AutoCADError(f"Error adding line: {e}") from e

    # Add a rectangle to the model space
    def add_rectangle(self, lower_left, upper_right):
        try:
            x1, y1 = lower_left.to_tuple()
            x2, y2 = upper_right.to_tuple()
            points = [
                x1, y1,
                x2, y1,
                x2, y2,
                x1, y2,
                x1, y1
            ]
            points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points)
            polyline = self.modelspace.AddLightweightPolyline(points_variant)
            return polyline
        except Exception as e:
            raise AutoCADError(f"Error adding rectangle: {e}") from e

    # Add an ellipse to the model space
    def add_ellipse(self, center, major_axis, ratio):
        try:
            ellipse = self.modelspace.AddEllipse(center.to_variant(), major_axis.to_variant(), ratio)
            return ellipse
        except Exception as e:
            raise AutoCADError(f"Error adding ellipse: {e}") from e

    # Add a text object to the model space
    def add_text(self, text):
        try:
            text_obj = self.modelspace.AddText(text.content, text.insertion_point.to_variant(), text.height)
            return text_obj
        except Exception as e:
            raise AutoCADError(f"Error adding text: {e}") from e

    # Add a dimension to the model space
    def add_dimension(self, dimension):
        try:
            dimension_obj = None
            if dimension.dimension_type == DimensionType.ALIGNED:
                dimension_obj = self.modelspace.AddDimAligned(dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.text_point.to_variant())
            elif dimension.dimension_type == DimensionType.LINEAR:
                # AutoCAD's Automation API has no AddDimLinear method; a linear
                # dimension is created via AddDimRotated with a rotation angle.
                dimension_obj = self.modelspace.AddDimRotated(dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.text_point.to_variant(), dimension.rotation_angle)
            elif dimension.dimension_type == DimensionType.ANGULAR:
                # AddDimAngular needs the angle vertex plus the two points on
                # the lines being measured; the constructor's start/end points
                # alone are not enough.
                if dimension.vertex_point is None:
                    raise ValueError("ANGULAR dimensions require 'vertex_point' (the angle vertex)")
                dimension_obj = self.modelspace.AddDimAngular(dimension.vertex_point.to_variant(), dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.text_point.to_variant())
            elif dimension.dimension_type == DimensionType.RADIAL:
                # AddDimRadial takes a leader length (a number), not a text point.
                if dimension.leader_length is None:
                    raise ValueError("RADIAL dimensions require 'leader_length'")
                dimension_obj = self.modelspace.AddDimRadial(dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.leader_length)
            elif dimension.dimension_type == DimensionType.DIAMETER:
                # The real Automation method is AddDimDiametric (not
                # AddDimDiameter) and it also takes a leader length, not a
                # text point.
                if dimension.leader_length is None:
                    raise ValueError("DIAMETER dimensions require 'leader_length'")
                dimension_obj = self.modelspace.AddDimDiametric(dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.leader_length)
            return dimension_obj
        except Exception as e:
            raise AutoCADError(f"Error adding dimension: {e}") from e

    # Add a point to the model space
    def add_point(self, point):
        try:
            point_obj = self.modelspace.AddPoint(point.to_variant())
            return point_obj
        except Exception as e:
            raise AutoCADError(f"Error adding point: {e}") from e

    # Add a polyline to the model space
    def add_polyline(self, points):
        try:
            points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [coord for point in points for coord in point.to_tuple()])
            polyline = self.modelspace.AddLightweightPolyline(points_variant)
            return polyline
        except Exception as e:
            raise AutoCADError(f"Error adding polyline: {e}") from e

    # Add a spline to the model space
    def add_spline(self, points):
        try:
            points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [coord for point in points for coord in point.to_variant()])
            spline = self.modelspace.AddSpline(points_variant)
            return spline
        except Exception as e:
            raise AutoCADError(f"Error adding spline: {e}") from e

    # Add an arc to the model space
    def add_arc(self, center, radius, start_angle, end_angle):
        try:
            arc = self.modelspace.AddArc(center.to_variant(), radius, start_angle, end_angle)
            return arc
        except Exception as e:
            raise AutoCADError(f"Error adding arc: {e}") from e

    # Explode an object or a set of joined objects
    def explode_object(self, obj):
        try:
            exploded_items = obj.Explode()
            return exploded_items
        except Exception as e:
            raise AutoCADError(f"Error exploding object: {e}") from e

    # Get maximum extents of a block
    def get_block_extents(self, block_name):
        try:
            for entity in self.iter_objects("AcDbBlockReference"):
                if entity.Name == block_name:
                    min_point = APoint(*entity.GeometricExtents.MinPoint)
                    max_point = APoint(*entity.GeometricExtents.MaxPoint)
                    return min_point, max_point
        except Exception as e:
            raise AutoCADError(f"Error getting extents of block '{block_name}': {e}") from e

    # Add overall dimensions to an object or block
    def add_overall_dimensions(self, entity):
        try:
            min_point, max_point = APoint(*entity.GeometricExtents.MinPoint), APoint(*entity.GeometricExtents.MaxPoint)
            # Horizontal dimension
            self.add_dimension(Dimension(min_point, APoint(max_point.x, min_point.y, min_point.z), APoint(min_point.x, min_point.y - 5, min_point.z), DimensionType.ALIGNED))
            # Vertical dimension
            self.add_dimension(Dimension(min_point, APoint(min_point.x, max_point.y, min_point.z), APoint(min_point.x - 5, min_point.y, min_point.z), DimensionType.ALIGNED))
        except Exception as e:
            raise AutoCADError(f"Error adding overall dimensions: {e}") from e

    # Get user-defined blocks in the document
    def get_user_defined_blocks(self):
        try:
            blocks = self.doc.Blocks
            user_defined_blocks = [block.Name for block in blocks 
                                   if not block.IsLayout and not block.Name.startswith('*') and block.Name != 'GENAXEH']
            return user_defined_blocks
        except Exception as e:
            raise AutoCADError(f"Error getting user-defined blocks: {e}") from e

    # Create a new layer
    def create_layer(self, layer):
        try:
            layers = self.doc.Layers
            new_layer = layers.Add(layer.name)
            new_layer.Color = layer.color.value
            return new_layer
        except Exception as e:
            raise AutoCADError(f"Error creating layer '{layer.name}': {e}") from e

    # Set the active layer
    def set_active_layer(self, layer_name):
        try:
            self.doc.ActiveLayer = self.doc.Layers.Item(layer_name)
        except Exception as e:
            raise AutoCADError(f"Error setting active layer '{layer_name}': {e}") from e

    # Insert a block into the model space
    def insert_block(self, block):
        try:
            block_ref = self.modelspace.InsertBlock(block.insertion_point.to_variant(), block.name, block.scale, block.scale, block.scale, block.rotation)
            return block_ref
        except Exception as e:
            raise AutoCADError(f"Error inserting block '{block.name}': {e}") from e

    # Save the document with a new name
    def save_as(self, file_path):
        try:
            self.doc.SaveAs(file_path)
        except Exception as e:
            raise AutoCADError(f"Error saving document as '{file_path}': {e}") from e

    # Open an existing file
    def open_file(self, file_path):
        try:
            self.acad.Documents.Open(file_path)
        except Exception as e:
            raise AutoCADError(f"Error opening file '{file_path}': {e}") from e

    # Get the insertion coordinates of a specific block
    def get_block_coordinates(self, block_name):
        try:
            block_references = []
            for entity in self.iter_objects("AcDbBlockReference"):
                if entity.Name == block_name:
                    insertion_point = entity.InsertionPoint
                    block_references.append(APoint(insertion_point[0], insertion_point[1], insertion_point[2]))
            return block_references
        except Exception as e:
            raise AutoCADError(f"Error getting coordinates of block '{block_name}': {e}") from e

    # Delete an object
    def delete_object(self, obj):
        try:
            obj.Delete()
        except Exception as e:
            raise AutoCADError(f"Error deleting object: {e}") from e

    # Clone an object
    def clone_object(self, obj, new_insertion_point):
        try:
            cloned_obj = obj.Copy(new_insertion_point.to_variant())
            return cloned_obj
        except Exception as e:
            raise AutoCADError(f"Error cloning object: {e}") from e

    # Modify a property of an object
    def modify_object_property(self, obj, property_name, new_value):
        try:
            setattr(obj, property_name, new_value)
        except Exception as e:
            raise AutoCADError(f"Error modifying property '{property_name}' of object: {e}") from e

    # Repeat a block horizontally until a specified length is reached
    def repeat_block_horizontally(self, block_name, total_length, block_length, insertion_point):
        try:
            x, y, z = insertion_point.x, insertion_point.y, insertion_point.z
            num_blocks = total_length // block_length

            for i in range(int(num_blocks)):
                new_insertion_point = APoint(x + i * block_length, y, z)
                self.insert_block(BlockReference(block_name, new_insertion_point))
        except Exception as e:
            raise AutoCADError(f"Error repeating block '{block_name}' horizontally: {e}") from e

    # Set the visibility of a layer
    def set_layer_visibility(self, layer_name, visible=True):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.LayerOn = visible
        except Exception as e:
            raise AutoCADError(f"Error setting visibility of layer '{layer_name}': {e}") from e

    # Lock or unlock a layer
    def lock_layer(self, layer_name, lock=True):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.Lock = lock
        except Exception as e:
            raise AutoCADError(f"Error locking/unlocking layer '{layer_name}': {e}") from e

    # Delete a layer
    def delete_layer(self, layer_name):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.Delete()
        except Exception as e:
            raise AutoCADError(f"Error deleting layer '{layer_name}': {e}") from e

    # Change the color of a layer
    def change_layer_color(self, layer_name, color):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.color = color.value
        except Exception as e:
            raise AutoCADError(f"Error changing color of layer '{layer_name}': {e}") from e

    # Set the linetype of a layer
    def set_layer_linetype(self, layer_name, linetype_name):
        try:
            layer = self.doc.Layers.Item(layer_name)
            linetypes = self.doc.Linetypes
            if linetype_name not in linetypes:
                self.doc.Linetypes.Load(linetype_name, linetype_name)
            layer.Linetype = linetype_name
        except Exception as e:
            raise AutoCADError(f"Error setting linetype of layer '{layer_name}': {e}") from e

    # Get a stable reference point for an object, falling back to the center
    # of its geometric extents when it has no InsertionPoint (e.g. lines,
    # circles, polylines)
    def _get_reference_point(self, obj):
        try:
            point = obj.InsertionPoint
            return APoint(point[0], point[1], point[2])
        except AttributeError:
            extents = obj.GeometricExtents
            min_point, max_point = extents.MinPoint, extents.MaxPoint
            return APoint(
                (min_point[0] + max_point[0]) / 2,
                (min_point[1] + max_point[1]) / 2,
                (min_point[2] + max_point[2]) / 2,
            )

    # Move an object
    def move_object(self, obj, new_insertion_point):
        try:
            current_point = self._get_reference_point(obj)
            obj.Move(current_point.to_variant(), new_insertion_point.to_variant())
        except Exception as e:
            raise AutoCADError(f"Error moving object: {e}") from e

    # Scale an object
    def scale_object(self, obj, base_point, scale_factor):
        try:
            obj.ScaleEntity(base_point.to_variant(), scale_factor)
        except Exception as e:
            raise AutoCADError(f"Error scaling object: {e}") from e

    # Rotate an object
    def rotate_object(self, obj, base_point, rotation_angle):
        try:
            obj.Rotate(base_point.to_variant(), rotation_angle)
        except Exception as e:
            raise AutoCADError(f"Error rotating object: {e}") from e

    # Align objects
    def align_objects(self, objects, alignment=Alignment.LEFT):
        try:
            if not objects:
                return
            reference_points = [(obj, self._get_reference_point(obj)) for obj in objects]
            if alignment == Alignment.LEFT:
                min_x = min(point.x for _, point in reference_points)
                for obj, point in reference_points:
                    self.move_object(obj, APoint(min_x, point.y, point.z))
            elif alignment == Alignment.RIGHT:
                max_x = max(point.x for _, point in reference_points)
                for obj, point in reference_points:
                    self.move_object(obj, APoint(max_x, point.y, point.z))
            elif alignment == Alignment.CENTER:
                xs = [point.x for _, point in reference_points]
                center_x = (min(xs) + max(xs)) / 2
                for obj, point in reference_points:
                    self.move_object(obj, APoint(center_x, point.y, point.z))
        except Exception as e:
            raise AutoCADError(f"Error aligning objects: {e}") from e

    # Distribute objects with specified spacing
    def distribute_objects(self, objects, spacing):
        try:
            if not objects:
                return
            objects.sort(key=lambda obj: self._get_reference_point(obj).x)
            for i in range(1, len(objects)):
                previous_point = self._get_reference_point(objects[i - 1])
                current_point = self._get_reference_point(objects[i])
                new_x = previous_point.x + spacing
                self.move_object(objects[i], APoint(new_x, current_point.y, current_point.z))
        except Exception as e:
            raise AutoCADError(f"Error distributing objects: {e}") from e

    # Insert a block from a file
    def insert_block_from_file(self, file_path, insertion_point, scale=1.0, rotation=0.0):
        try:
            block_name = self.doc.Blocks.Import(file_path, file_path)
            block_ref = self.modelspace.InsertBlock(insertion_point.to_variant(), block_name, scale, scale, scale, rotation)
            return block_ref
        except Exception as e:
            raise AutoCADError(f"Error inserting block from file '{file_path}': {e}") from e

    # Export a block to a file
    def export_block_to_file(self, block_name, file_path):
        try:
            block = self.doc.Blocks.Item(block_name)
            block.Export(file_path)
        except Exception as e:
            raise AutoCADError(f"Error exporting block '{block_name}' to '{file_path}': {e}") from e

    # Modify a block attribute
    def modify_block_attribute(self, block_ref, tag, new_value):
        try:
            for attribute in block_ref.GetAttributes():
                if attribute.TagString == tag:
                    attribute.TextString = new_value
        except Exception as e:
            raise AutoCADError(f"Error modifying block attribute '{tag}': {e}") from e

    # Delete a block attribute
    def delete_block_attribute(self, block_ref, tag):
        try:
            for attribute in block_ref.GetAttributes():
                if attribute.TagString == tag:
                    attribute.Delete()
        except Exception as e:
            raise AutoCADError(f"Error deleting block attribute '{tag}': {e}") from e

    # Request point input from the user
    def get_user_input_point(self, prompt="Select a point"):
        try:
            point = self.doc.Utility.GetPoint(None, prompt)
            return APoint(point[0], point[1], point[2])
        except Exception as e:
            raise AutoCADError(f"Error getting point input from user: {e}") from e

    # Request string input from the user
    def get_user_input_string(self, prompt="Enter a string"):
        try:
            return self.doc.Utility.GetString(False, prompt)
        except Exception as e:
            raise AutoCADError(f"Error getting string input from user: {e}") from e

    # Request integer input from the user
    def get_user_input_integer(self, prompt="Enter an integer"):
        try:
            return self.doc.Utility.GetInteger(prompt)
        except Exception as e:
            raise AutoCADError(f"Error getting integer input from user: {e}") from e

    # Display a message to the user
    def show_message(self, message):
        try:
            self.doc.Utility.Prompt(message + "\n")
        except Exception as e:
            raise AutoCADError(f"Error displaying message: {e}") from e

    # Create a group of objects
    def create_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Add(group_name)
            
            # Ensure objects is a list or tuple
            if not isinstance(objects, (list, tuple)):
                objects = [objects]
            
            # Create SAFEARRAY of IDispatch pointers
            variant_array = win32com.client.VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,
                objects
            )
            group.AppendItems(variant_array)
            return group
        except Exception as e:
            raise AutoCADError(f"Error creating group '{group_name}': {e}") from e

    # Add objects to a group
    def add_to_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Item(group_name)
            for obj in objects:
                group.AppendItems([obj])
        except Exception as e:
            raise AutoCADError(f"Error adding objects to group '{group_name}': {e}") from e

    # Remove objects from a group
    def remove_from_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Item(group_name)
            for obj in objects:
                group.RemoveItems([obj])
        except Exception as e:
            raise AutoCADError(f"Error removing objects from group '{group_name}': {e}") from e

    # Select a group of objects
    def select_group(self, group_name):
        try:
            group = self.doc.Groups.Item(group_name)
            return [item for item in group.GetItems()]
        except Exception as e:
            raise AutoCADError(f"Error selecting group '{group_name}': {e}") from e
