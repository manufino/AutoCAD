
# AutoCAD Python Automation

This Python library provides a set of classes and methods to interact with AutoCAD using the COM API. The library allows automating many common operations in AutoCAD, such as creating and managing layers, objects, blocks, attributes, and groups of objects.

## Main Features

- **Layer Management**: Create, modify, set visibility, lock/unlock, change color, and manage layer linetype.
- **Object Management**: Create, select, move, scale, rotate, align, and distribute objects.
- **Block Management**: Insert, export, create, modify, and remove blocks.
- **Attribute Management**: Add, modify, and delete block attributes.
- **User Input and Output**: Request input from the user (points, strings, integers) and display messages.
- **Group Management**: Create, modify, add/remove objects, and select groups.

## Requirements

- AutoCAD installed on the system.
- Python 3.x.
- `pywin32` package installed (installable via pip).

## Installation

1. Clone this repository:
   ```sh
   git clone https://github.com/your-username/autocad-python-automation.git
   ```

2. Install the dependencies:
   ```sh
   pip install pywin32
   ```

## Usage Examples

Below are some examples of how to use the library to automate operations in AutoCAD.

## Create the AutoCAD object

```python
# Create the AutoCAD object
acad = AutoCAD()
```

## Repeat block horizontally

```python
# Repeat the "blockname" block horizontally
total_length = 100  # Total length X
block_length = 10  # Length of the block "blockname"
insertion_point = APoint(0, 0, 0)  # Initial insertion point

# Execute the block repetition
acad.repeat_block_horizontally("blockname", total_length, block_length, insertion_point)
```

## Set the visibility of a layer

```python
# Set the visibility of a layer
acad.set_layer_visibility("Linea di mezzeria", visible=False)
```

## Lock a layer

```python
# Lock a layer
acad.lock_layer("Quote", lock=True)
```

## Delete a layer

```python
# Delete a layer
acad.delete_layer("Simboli")
```

## Change the color of a layer

```python
# Change the color of a layer
acad.change_layer_color("Contorni", Color.YELLOW)
```

## Set the linetype of a layer

```python
# Set the linetype of a layer
acad.set_layer_linetype("Assi", "DASHED")
```

## Select objects

```python
# Select objects
selected_objects = acad.select_objects(object_type="AcDbLine", layer_name="Contorni")
print(f"Selected objects: {len(selected_objects)}")
```

## Move, scale, and rotate objects

```python
# Move, scale, and rotate objects
for obj in selected_objects:
    acad.move_object(obj, APoint(10, 10, 0))
    acad.scale_object(obj, APoint(0, 0, 0), 2)
    acad.rotate_object(obj, APoint(0, 0, 0), 45)
```

## Align and distribute objects

```python
# Align and distribute objects
acad.align_objects(selected_objects, alignment="left")
acad.distribute_objects(selected_objects, spacing=5)
```

## Insert a block from a file

```python
# Insert a block from a file
acad.insert_block_from_file("path_to_file.dwg", APoint(0, 0, 0))
```

## Export a block to a file

```python
# Export a block to a file
acad.export_block_to_file("piatto", "path_to_export.dwg")
```

## Modify block attributes

```python
# Modify block attributes
block_references = acad.get_block_coordinates("piatto")
if block_references:
    block_ref = block_references[0]  # Get the first found block
    acad.modify_block_attribute(block_ref, "Tag", "NewValue")
```

## Delete block attributes

```python
# Delete block attributes
acad.delete_block_attribute(block_ref, "Tag")
```

## Request user input

```python
# Request user input
point = acad.get_user_input_point("Select a point")
text = acad.get_user_input_string("Enter a string")
integer = acad.get_user_input_integer("Enter an integer")
```

## Display a message to the user

```python
# Display a message to the user
acad.show_message("Operation completed")
```

## Create a group of objects

```python
# Create a group of objects
group = acad.create_group("MyGroup", selected_objects)
```

## Add objects to a group

```python
# Add objects to a group
acad.add_to_group("MyGroup", selected_objects)
```

## Remove objects from a group

```python
# Remove objects from a group
acad.remove_from_group("MyGroup", selected_objects)
```

## Select a group of objects

```python
# Select a group of objects
group_items = acad.select_group("MyGroup")
print(f"Objects in group 'MyGroup': {len(group_items)}")
```

## Print the created layers for confirmation

```python
# Print the created layers for confirmation
for layer in acad.doc.Layers:
    print(f"Layer: {layer.Name}, Color: {layer.color}")
```
