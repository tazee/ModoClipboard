# External Clipboard for Modo plug-in

This kit provides the ability to transfer mesh data between Modo and Blender via an external clipboard. When you execute the Copy command, the mesh data to be transferred is converted to JSON format text data and saved to a temporary file or the OS clipboard. When you execute the Paste command, the saved clipboard data is created in the current mesh item or a new mesh item. The selected polygons of the primary mesh will be output.<br>

The external clipboard for Blender is provided as part of [YT-Tools for Blender](https://tazaki.gumroad.com/l/gamki).

This kit tested on macOS and Windows.

The following data is import and export:

**Vertex**<br>
Vertex position data is output. Modo's coordinate system is right-handed and Y-Up. The coordinate system will be converted when loading into Blender.<br>

**Polygon**<br>
The selected polygon data for faces and subdivision surfaces will be output. Other polygons such as curves are not supported. All surface polygons are output as Faces.

**Morph**<br>
Relative and absolute morph data is output as Shapekey data.

**Weight**<br>
Vertex weight map is output as Blender Vertex Group data.

**Subdivision Weight**<br>
Subdivision Edge Weight map data is output as Blender's Crease Edge.

**Material**<br>
Material and texture data are output. Only diffuse color and the texture file path are supported.

**UV Map**<br>
UV maps is output as Blender UV Set data.

**Freestyle edge**<br>
Import Freestyle edges from Blender into edge selection set vmap named "_Freestyle". If an edge selection set named "_Freestyle" is in the current mesh, it will be output as Freestyle data.

**RGB, RGBA Map**<br>
RGB, RGBA maps are output as Blender Color Attributes data with Face_Corner domain.

**Selection Sets**<br>
Selection Sets for vertex, edge and polygon are output to the selection set feature of YT-Tools for Blender

**UV Seam**<br>
The primary UV Seam map or the UV Seam with "_Seam" name is output to edge seam mark of Blender

## Installing

- Download lpk from releases. Drag and drop into your Modo viewport. If you're upgrading, delete previous version.

## How to use Modo Clipboard

The clipboard panel can be displayed by clicking the clipboard icon in the Kits icon in the upper right corner of the screen.

<div align="left">
<img src="images/UI.png" style='max-height: 620px; object-fit: contain'/>
</div>

### Cut, Copy<br>

**Copy** and **Cut** export the currently selected polygons or all mesh data to the clipboard. Cut exports the data and then deletes the currently selected polygon data.

### Paste<br>

Imports mesh data from the clipboard and constructs it onto the currently selected mesh.

### New Mesh from Clipboard

Builds the mesh imported from the clipboard into a newly created mesh item, replacing the current mesh.

### Type

Specifies the type of external clipboard. The default is **Temporary File**. If you specify **OS Clipboard**, the converted JSON text data will be used on the OS standard clipboard. This can be used to check or modify the copied data.
<br>

## History

### v1.0.1 Bug Fix

- Fixed several bugs
- Omited unnecessary edge data to reduce file size

### v1.0.2 Minor Changes

- Support Vertex Colors, Selection Sets and UV Seams
- Export key hole polygon as triangles
