#
# External Clipboard to exchange CPMF data between Modo and Blender
#
# CPMF: Custom Portable Mesh Format
#

from email.mime import image
import lx
import modo
import logging
import json
import sys
import subprocess
from datetime import datetime
import os
import tempfile


# ---------- Clipboard helpers ----------
try:
    import pyperclip
    def clipboard_copy(text): pyperclip.copy(text)
    def clipboard_paste(): return pyperclip.paste()
except Exception:
    def clipboard_copy(text):
        if sys.platform == 'darwin':
            p = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
            p.communicate(text.encode('utf-8'))
        elif sys.platform.startswith('linux'):
            p = subprocess.Popen(['xclip','-selection','clipboard'], stdin=subprocess.PIPE)
            p.communicate(text.encode('utf-8'))
        elif sys.platform.startswith('win'):
            cmd = ['powershell', '-NoProfile', '-Command', 'Set-Clipboard -Value ([Text.Encoding]::Utf8.GetString([Text.Encoding]::Utf8.GetBytes($input)))']
            p = subprocess.Popen(cmd, stdin=subprocess.PIPE)
            p.communicate(text.encode('utf-8'))
        else:
            raise RuntimeError('No clipboard method')

    def clipboard_paste():
        if sys.platform == 'darwin':
            return subprocess.check_output(['pbpaste']).decode('utf-8')
        elif sys.platform.startswith('linux'):
            return subprocess.check_output(['xclip','-selection','clipboard','-o']).decode('utf-8')
        elif sys.platform.startswith('win'):
            return subprocess.check_output(['powershell','-NoProfile','-Command','Get-Clipboard']).decode('utf-8')
        else:
            raise RuntimeError('No clipboard method')

# ---------- small helpers ----------
def get_cpmf_tempfile_path():
    temp_dir = tempfile.gettempdir()
    path = os.path.join(temp_dir, "cpmf_clipboard.json")
    return path

def write_tempfile(text, path=None):
    """Write text to path if provided; otherwise create in OS tempdir and return path."""
    if path:
        d = os.path.dirname(path)
        if d and not os.path.exists(d):
            try:
                os.makedirs(d, exist_ok=True)
            except Exception:
                pass
        with open(path, 'w', encoding='utf-8') as f:
            f.write(text)
        return os.path.abspath(path)
    else:
        fn = f"CPMF_{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}.json"
        tmp = os.path.join(tempfile.gettempdir(), fn)
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(text)
        return tmp

def read_tempfile(path):
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()

# ---------- (the other utility functions are the same as in the v1.5 code) ----------
# For brevity, core functions (coordinate conversion, material helpers, gatherers, etc.)
# are taken from the previous v1.5 implementation. We'll include them here verbatim
# so the addon is self-contained.
# ---------- Coordinate conversion helpers ----------
#
# Modo (Y-Up) -> Blender (Z-Up)
M2B = modo.Matrix3((
    (1.0, 0.0,  0.0),
    (0.0, 0.0, -1.0),
    (0.0, 1.0,  0.0),
))

# Blender (Z-Up) -> Modo (Y-Up)
B2M= modo.Matrix3((
    (1.0, 0.0,  0.0),
    (0.0, 0.0,  1.0),
    (0.0, -1.0, 0.0),
))

def mat3_mul_vec3(M, v):
    """
    M: Modo Matrix3 オブジェクト（3x3行列）
    v: (x, y, z) のタプル または Vector3

    戻り値: (x', y', z') のタプル
    """
    x = M[0][0] * v[0] + M[0][1] * v[1] + M[0][2] * v[2]
    y = M[1][0] * v[0] + M[1][1] * v[1] + M[1][2] * v[2]
    z = M[2][0] * v[0] + M[2][1] * v[1] + M[2][2] * v[2]

    return modo.Vector3((x, y, z))


# --- Vector conversion ---
def convert_vector_from_coord(vec, src_coord):
    if not src_coord:
        return modo.Vector3(vec)

    key = src_coord.lower()
    v = modo.Vector3(vec)

    # Blender → Modo
    if "z_up" in key:
        return mat3_mul_vec3(B2M, v)

    # Already Modo space
    return v

# --- Quaternion conversion ---
def convert_quaternion_from_coord(q_in, src_coord):
    if not src_coord or "z_up" not in src_coord.lower():
        return modo.Quaternion((q_in[0], q_in[1], q_in[2], q_in[3]))

    # Quaternion → Matrix
    Rb = modo.Quaternion((q_in[0], q_in[1], q_in[2], q_in[3])).toMatrix3()

    # Blender rotation → Modo rotation
    Rm = B2M * Rb * M2B

    # Matrix → Quaternion
    return modo.Quaternion.fromMatrix3(Rm)


# --- Transform conversion (Location / Rotation / Scale) ---
def convert_matrix_transform_from_coord(translation, rotation_quat, scale, src_coord):
    t = convert_vector_from_coord(translation, src_coord)
    q = convert_quaternion_from_coord(rotation_quat, src_coord)
    s = modo.Vector3(scale)
    return t, q, s


# Clipboard class
class ClipboardData:
    def __init__(self):
        self.mesh = None
        self.geom = None
        self.vertices = []
        self.edges = []
        self.polygons = []
        self.materials = []
        self.mark_select = None
        self.selType = None
        self.base_nvert = 0

    def selected(self, v):
        return v.accessor.TestMarks(self.mark_select)

    def index(self, v):
        return self.vertex_indices[v.index]
    
    def lookupMap(self, map_type, name):
        for vmap in self.geom.vmaps:
            if vmap.map_type == map_type and vmap.name == name:
                return vmap
        return None

    def setWeight(self, vmap, v, weight):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        storage.set([weight])
        v.accessor.SetMapValue(vmap.id, storage)

    def setMorph(self, vmap, v, pos):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(3)
        storage.set(pos)
        v.accessor.SetMapValue(vmap.id, storage)

    def setEdgePick(self, vmap):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        self.edge_accessor.SetMapValue(vmap.id, storage)

    def setSubdivWeight(self, vmap, weight):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        storage.set([weight])
        self.edge_accessor.SetMapValue(vmap.id, storage)

    def setup_mesh_elements(self):
        # store selected vertices
        self.vertex_indices = {}
        self.vertices.clear()
        index = 0
        for v in self.geom.vertices:
            if self.selected(v):
                self.vertex_indices[v.index] = index
                self.vertices.append(v)
                index += 1

        # store selected edges
        self.edges.clear()
        for e in self.geom.edges:
            v0, v1 = e.vertices
            if self.selected(v0) and self.selected(v1):
                self.edges.append(e)

        # store selected polygons
        self.polygons.clear()
        for p in self.geom.polygons:
            ptype = p.Type()
            if ptype != lx.symbol.iPTYP_FACE and \
               ptype != lx.symbol.iPTYP_SUBD and \
               ptype != lx.symbol.iPTYP_PSUB:
                continue
            if self.selected(p):
                self.polygons.append(p)

        if self.selType == lx.symbol.iSEL_VERTEX:
            return True if len(self.vertices) > 0 else False
        elif self.selType == lx.symbol.iSEL_EDGE:
            return True if len(self.edges) > 0 else False
        else:
            return True if len(self.polygons) > 0 else False

    def paste_vertices(self, positions):
        coord = self.data.get('metadata', {}).get('coordinate_system', '').lower()
        unit_scale = float(self.data.get('metadata', {}).get('unit_scale', 1.0))
        self.vertices = []
        self.base_nvert = len(self.geom.vertices)
        for i, p in enumerate(positions):
            v = convert_vector_from_coord(p, coord)
            v *= unit_scale
            self.vertices.append(self.geom.vertices.new((v.x, v.y, v.z)))

    def paste_polygons(self, polygons, materials):
        count = 0
        for poly in polygons:
            vert_indices = poly.get('vertices', [])
            face_verts = [self.vertices[i] for i in vert_indices]
            p = self.geom.polygons.new(face_verts)
            self.polygons.append(p)
            attributes = poly.get('attributes', {})
            if 'material_index' in attributes:
                try:
                    material_index = int(attributes['material_index'])
                    if material_index is not None and 0 <= material_index < len(materials):
                        self.geom.polygons[count].materialTag = materials[material_index].get('name', '')
                except Exception:
                    pass
            count += 1

    def select_edge(self, vertices):
        i0, i1 = vertices[0], vertices[1]
        i0 += self.base_nvert
        i1 += self.base_nvert
        if i0 >= self.mesh.PointCount() or i1 >= self.mesh.PointCount():
            return False
        self.point_accessor.SelectByIndex(i0)
        id0 = self.point_accessor.ID()
        self.point_accessor.SelectByIndex(i1)
        id1 = self.point_accessor.ID()
        #print(f"select_edge: [{i0} {i1}] vertices {vertices} id0 {id0} id1 {id1}")
        if id0 is None or id1 is None:
            return False
        self.edge_accessor.SelectEndpoints(id0, id1)
        return True
    
    def paste_edges(self, edges):
        vmap = self.lookupMap(lx.symbol.i_VMAP_SUBDIV, "Subdivision")
        #print(f"paste_edges: vmap {vmap} edges {len(self.geom.edges)}")
        if vmap is None:
            return
        for e in edges:
            if self.select_edge(e.get('vertices', [])) == False:
                continue
            weight = e.get('attributes', {}).get('crease_edge', 0.0)
            if weight > 0.0 and vmap is not None:
                self.setSubdivWeight(vmap, weight)


    def get_type_name(self, type):
        if type == 'base_color':
            return 'diffColor'
        return None

    def paste_textures(self, material, mask):
        base_dir = self.data.get('metadata', {}).get('custom', {}).get('base_dir', None)
        for tex in material.get('textures', []):
            effect = self.get_type_name(tex.get('type', ''))
            if not effect:
                continue
            img_path = tex.get('image')
            uv_map_name = tex.get('uv_map') or ''
            if not img_path:
                continue
            if base_dir and not os.path.isabs(img_path):
                candidate = os.path.join(base_dir, img_path)
                if os.path.exists(candidate):
                    img_path = candidate
            name = os.path.basename(img_path)
            #print(f"image_path {img_path} name {name}")
            layer = self.scene.addItem('imageMap')
            layer.SetName(name)
            layer.channel('effect').set(effect)
            layer.setParent(mask, index=1)
            graph = layer.itemGraph(lx.symbol.sGRAPH_SHADELOC)
            #print(f"-- add imageMap layer {layer.name}")
            item_image = None
            item_txtrLocator = None
            for it in graph.forward():
                #print(f"linked item {it.name} type {it.type}")
                if it.type == 'videoStill':
                    it.channel('filename').set(img_path)
                    item_image = it
                elif it.type == 'txtrLocator':
                    it.channel('projType').set('uv')
                    if uv_map_name:
                        it.channel('uvMap').set(uv_map_name)
                    item_txtrLocator = it
            if item_image is None:
                item = self.scene.addItem('videoStill')
                graph.AddLink(layer, item)
                item.channel('filename').set(img_path)
            if item_txtrLocator is None:
                item = self.scene.addItem('txtrLocator')
                graph.AddLink(layer, item)
                item.channel('projType').set('uv')
                if uv_map_name:
                    item.channel('uvMap').set(uv_map_name)
                

    def paste_materials(self, materials):
        for material in materials:
            name = material.get('name', '')
            col = material.get('base_color', ())
            mat = self.scene.addMaterial(name='M_' + name)
            mat.channel('diffCol').set(col)
            mask = self.scene.addItem('mask', name=name)
            mask.channel('ptag').set(name)
            mat.setParent(mask, index=1)
            self.paste_textures(material, mask)


    def paste_uv_sets(self, uv_sets):
        for uv_set in uv_sets:
            name = uv_set.get('name', '')
            vmap = self.lookupMap(lx.symbol.i_VMAP_TEXTUREUV, name)
            if not vmap:
                vmap = self.geom.vmaps.addUVMap(name)
            for face_uvs in uv_set.get('uvs', []):
                index = face_uvs.get('index')
                values = face_uvs.get('values', [])
                p = self.polygons[index]
                for i, v in enumerate(p.vertices):
                    uv = values[i]
                    p.setUV(uv, v, uvmap=vmap)

    def paste_vertex_groups(self, vertex_groups):
        for vertex_group in vertex_groups:
            name = vertex_group.get('name', '')
            weights = vertex_group.get('weights', [])
            vmap = self.lookupMap(lx.symbol.i_VMAP_WEIGHT, name)
            if not vmap:
                vmap = self.geom.vmaps.addWeightMap(name)
            for w_data in weights:
                index = w_data.get('index')
                weight = w_data.get('weight', 0.0)
                v = self.vertices[index]
                self.setWeight(vmap, v, weight)

    def paste_vertex_shapekeys(self, shapekeys):
        coord = self.data.get('metadata', {}).get('coordinate_system', '').lower()
        unit_scale = float(self.data.get('metadata', {}).get('unit_scale', 1.0))
        base_positions = None
        for shapekey in shapekeys:
            name = shapekey.get('name')
            if name == 'Basis':
                base_positions = shapekey.get('positions', [])
                continue
            use_relative = shapekey.get('relative', True)
            map_type = lx.symbol.i_VMAP_MORPH if use_relative else lx.symbol.i_VMAP_SPOT
            vmap = self.lookupMap(lx.symbol.i_VMAP_WEIGHT, name)
            if not vmap:
                vmap = self.geom.vmaps.addMorphMap(name, (map_type is lx.symbol.i_VMAP_SPOT))
            for pos_data in shapekey.get('positions', []):
                index = pos_data.get('index')
                pos = modo.Vector3(pos_data.get('position'))
                v = self.vertices[index]
                if use_relative:
                    if base_positions is None:
                        pos -= v.position
                    else:
                        base_data = base_positions[index]
                        pos -= modo.Vector3(base_data.get('position'))
                pos = convert_vector_from_coord(pos, coord)
                pos *= unit_scale
                self.setMorph(vmap, v, pos)

    def paste_edge_freestyle(self, freestyle_edges):
        name = '_Freestyle'
        vmap = self.lookupMap(lx.symbol.i_VMAP_EPCK, name)
        if not vmap:
            vmap = self.geom.vmaps.addMap(lx.symbol.i_VMAP_EPCK, name)
        for data in freestyle_edges:
            vertices = data.get('vertices')
            use_freestyle_mark = data.get('use_freestyle_mark', 0)
            if use_freestyle_mark:
                if self.select_edge(vertices) == False:
                    continue
                self.setEdgePick(vmap)

    # Main paste function
    def paste(self, external_clipboard='tempfile', new_mesh=False):
        #print(f'Pasting from external clipboard: {external_clipboard}, new_mesh={new_mesh}')
        if external_clipboard == 'tempfile':
            path = get_cpmf_tempfile_path()
            lx.out(f'Read file from: {path}')
            if not path:
                lx.out({'ERROR'}, 'No file path specified for import')
                return False
            try:
                txt = read_tempfile(path)
            except Exception as e:
                lx.out({'ERROR'}, f'Failed to read file: {e}')
                return False
        else:
            try:
                txt = clipboard_paste()
            except Exception as e:
                lx.out({'ERROR'}, f'Failed to read clipboard: {e}')
                return False

        # parse
        try:
            self.data = json.loads(txt)
        except Exception as e:
            lx.out({'ERROR'}, f'Invalid JSON: {e}')
            return False
        
        # Add a new mesh object to the scene and grab the geometry object
        self.scene = modo.Scene()

        for obj_data in self.data.get('objects', []):
            mesh_data = obj_data.get('mesh', {})
            positions = mesh_data.get('positions', [])
            edges = mesh_data.get('edges', [])
            polygons = mesh_data.get('polygons', [])
            materials = mesh_data.get('materials', []) or mesh_data.get('mesh', {}).get('materials', []) or mesh_data.get('materials', [])
            uv_sets = mesh_data.get('uv_sets', [])
            shapekeys = mesh_data.get('shapekeys', [])
            vertex_groups = mesh_data.get('vertex_groups', [])
            freestyle_edges = mesh_data.get('freestyle_edges', [])

            if new_mesh == True:
                mesh = self.scene.addMesh("Mesh")
                self.scene.select(mesh)
            else:
                mesh = self.scene.selectedByType("mesh")
                if mesh:
                    mesh = mesh[0]
                    self.scene.select(mesh)
                else:
                    lx.out({'ERROR'}, 'No mesh selected to paste into')
                    return False

            self.mesh = mesh
            self.geom = mesh.geometry
            self.geom.setAccessMode(value='write')
        
            # paste positions to geometry and apply unit scale
            if positions:
                self.paste_vertices(positions)

            # paste polygons data to geometry
            if polygons:
                self.paste_polygons(polygons, materials)

            # paste materials data to geometry
            if materials:
                self.paste_materials(materials)

            # paste uv sets data to geometry
            if uv_sets:
                self.paste_uv_sets(uv_sets)

            # paste vertex groups data to geometry
            if vertex_groups:
                self.paste_vertex_groups(vertex_groups)

            # paste vertex shapekeys data to geometry
            if shapekeys:
                self.paste_vertex_shapekeys(shapekeys)

            self.geom.setMeshEdits()

            # Apply edge vertex map values using layer scan since ModoGeometry 
            # does not work for writing edge vmaps
            layer_svc = lx.service.Layer()
            layer_scan = lx.object.LayerScan(layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_EDIT))
            if layer_scan.test() == False:
                return

            self.mesh = lx.object.Mesh (layer_scan.MeshEdit (0))
            self.edge_accessor = lx.object.Edge (self.mesh.EdgeAccessor ())
            self.point_accessor = lx.object.Point (self.mesh.PointAccessor ())

            # paste edges data to geometry
            if edges:
                self.paste_edges(edges)

            # paste freestyle data to geometry as EdgePick map
            if freestyle_edges:
                self.paste_edge_freestyle(freestyle_edges)

            layer_scan.SetMeshChange(0, lx.symbol.f_MESHEDIT_MAP_OTHER)
            layer_scan.Apply()


    def get_material_index(self, name, material_out):
        for i, mat in enumerate(material_out):
            if mat['name'] == name:
                return i
        return 0
          
    def get_edge_index(self, edge):
        for i, e in enumerate(self.edges):
            if e == edge:
                return i
        return 0

    def extract_uv_sets(self):
        if len(self.polygons) == 0:
            return None
        uv_sets = []
        for vmap in self.geom.vmaps.uvMaps:
            uv_set = {
                'name': vmap.name,
                'uvs': []
            }
            for i, p in enumerate(self.polygons):
                face_uvs = {
                    'index': i,
                    'values': []
                }
                for v in p.vertices:
                    uv = p.getUV(v, uvmap=vmap)
                    face_uvs['values'].append([uv[0], uv[1]])
                uv_set['uvs'].append(face_uvs)
            uv_sets.append(uv_set)
        if len(uv_sets) == 0:
            return None
        return uv_sets

    def extract_vertex_groups(self):
        if len(self.vertices) == 0:
            return None
        vertex_groups = []
        for vmap in self.geom.vmaps:
            if vmap.map_type != lx.symbol.i_VMAP_WEIGHT:
                continue
            vg_data = {
                'name': vmap.name,
                'weights': []
            }
            for v in self.vertices:
                w = vmap[v.index]
                if w is not None:
                    vg_data['weights'].append({'index': self.index(v), 'weight': w[0]})
            vertex_groups.append(vg_data)
        if len(vertex_groups) == 0:
            return None
        return vertex_groups

    def extract_vertex_shapekeys(self):
        if len(self.geom.vmaps.morphMaps) == 0:
            return None
        if len(self.vertices) == 0:
            return None
        shapekeys = []
        # Add Basis positions
        sk_data = {
            'name': 'Basis',
            'relative': True,
            'positions': []
        }
        for v in self.vertices:
            sk_data['positions'].append({'index': self.index(v), 'position':[v.position[0], v.position[1], v.position[2]]})
        shapekeys.append(sk_data)
        # Add all morph and spot vertex maps
        for vmap in self.geom.vmaps.morphMaps:
            relative = True if (vmap.map_type == lx.symbol.i_VMAP_MORPH) else False
            sk_data = {
                'name': vmap.name,
                'relative': relative,
                'positions': []
            }
            for v in self.vertices:
                co = vmap.getAbsolutePosition(v.index)
                sk_data['positions'].append({'index': self.index(v), 'position':[co[0], co[1], co[2]]})
            shapekeys.append(sk_data)
        if len(shapekeys) == 0:
            return None
        return shapekeys

    def extract_edge_freestyle(self):
        if len(self.edges) == 0:
            return None
        freestyle_edges = []
        for vmap in self.geom.vmaps:
            if vmap.map_type != lx.symbol.i_VMAP_EPCK:
                continue
            if vmap.name != '_Freestyle':
                continue
            for e in self.edges:
                storageBuffer = lx.object.storage('f', 1)
                if e.MapEvaluate(vmap.id, storageBuffer) == True:
                    freestyle_edges.append({'vertices': [self.index(v) for v in e.vertices], 'use_freestyle_mark': 1})
        if len(freestyle_edges) == 0:
            return None
        return freestyle_edges
    
    # get Blender's type name from effect channel of 'imageMap' item
    def get_effect_name(self, layer):
        channel = layer.channel("effect")
        if channel is None:
            return None
        effect = channel.get()
        if effect == 'diffColor':
                return 'base_color'
        return None
    
    def get_imageMap_items(self):
        return [item for item in self.scene.items() if item.type == 'imageMap']

    # extract textures from the material
    def extract_textures(self, material):
        # mask item must be the parent of material
        mask = material.parent
        textures = []
        # iterate all child items under mask item
        if not mask:
            layers = self.get_imageMap_items()
        else:
            layers = mask.children()
        if not layers:
            return None
        for layer in layers:
            # imageMap may have texture image and uv map
            if layer.type == 'imageMap':
                # convert Modo's effect name to Blender's one
                effect = self.get_effect_name(layer)
                if effect is None:
                    continue
                texture = {
                    'type': effect
                }
                # find texture locator and videoStill item linking to the image map
                for it in layer.itemGraph (lx.symbol.sGRAPH_SHADELOC).forward ():
                    # Texture Locator
                    if it.type == 'txtrLocator':
                        # check if the projection type is 'uv'
                        channel = it.channel("projType")
                        if channel is None:
                            continue
                        projType = channel.get()
                        if projType != 'uv':
                            continue
                        # get UV map name
                        channel = it.channel("uvMap")
                        if channel is None:
                            continue
                        uvmap = channel.get()
                        texture['uv_map'] = uvmap
                    # Video Still (image file)
                    elif it.type == 'videoStill':
                        channel = it.channel("filename")
                        if channel is None:
                            continue
                        filename = channel.get()
                        texture['image'] = filename

                if texture['uv_map'] is not None and texture['image'] is not None:
                    textures.append(texture)
        if len(textures) == 0:
            return None
        return textures

    # copy all materials
    def copy_materials(self):
        if len(self.polygons) == 0:
            return None
        # Query Existing Materials
        self.materials = []
        for material in self.scene.items("advancedMaterial"):
            mask = material.parent
            if not mask or mask.type != 'mask':
                name = 'Default'
            else:
                lx.out(f"material {material.name} mask {mask.name}")
                name = mask.channel('ptag').get()
            diffCol = material.channel('diffCol').get()

            mat_data = {
                'name': name,
                'base_color': diffCol
            }
            # extract texture data for the material
            textures = self.extract_textures(material)
            if textures is not None:
                mat_data['textures'] = textures
            lx.out(f'Found material: {name} color: {diffCol}')
            self.materials.append(mat_data)
        return self.materials
    
    # copy all selected vertices
    def copy_vertices(self):
        if len(self.vertices) == 0:
            return None
        positions = []
        for v in self.vertices:
            positions.append([v.position[0], v.position[1], v.position[2]])
        return positions

    # copy all selected edges
    def copy_edges(self):
        if len(self.polygons) == 0:
            return None
        if len(self.edges) == 0:
            return None
        # crease edges
        crease_edges = [0.0] * len(self.geom.edges)
        for vmap in self.geom.vmaps:
            if vmap.map_type == lx.symbol.i_VMAP_SUBDIV:
                lx.out(f'Found vmap: {vmap.name} type: {vmap.map_type} len: {len(vmap)}')
                storageBuffer = lx.object.storage('f', 1)
                for i, e in enumerate(self.geom.edges):
                    if e.MapEvaluate(vmap.id, storageBuffer) == True:
                        w = storageBuffer.get()
                        crease_edges[i] = w[0]
                break
        # edges
        edges = []
        for i, e in enumerate(self.geom.edges):
            v0, v1 = e.vertices
            if self.selected(v0) and self.selected(v1):
                edges.append({
                    'vertices': [self.index(v) for v in e.vertices],
                    'attributes': {
                        'crease_edge': crease_edges[i]
                    }
                })
        return edges

    # copy all selected polygons
    def copy_polygons(self):
        if len(self.polygons) == 0:
            return None
        polygons = []
        for p in self.polygons:
            p_attrs = {
                'material_index': self.get_material_index(p.materialTag, self.materials)
            }
            polygons.append({
                'vertices': [self.index(v) for v in p.vertices],
                'attributes': p_attrs
            })
        return polygons


    # Main copy function
    def copy(self, external_clipboard='tempfile'):
        lx.out(f'Copying to external clipboard: {external_clipboard}')

        self.scene = modo.Scene()
        mesh = self.scene.selectedByType("mesh")
        if not mesh:
            return False

        layer_svc = lx.service.Layer()
        layer_scan = lx.object.LayerScan(layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_PRIMARY | lx.symbol.f_LAYERSCAN_MARKALL))
        if layer_scan.test() == False:
            return False
        
        # get the current selection type
        sel_svc = lx.service.Selection()
        types = lx.object.storage('i', 4)
        types.set((lx.symbol.iSEL_VERTEX, lx.symbol.iSEL_EDGE, lx.symbol.iSEL_POLYGON, 0))
        self.selType = sel_svc.CurrentType(types)
    
        # selection mark for mesh element
        mesh_svc = lx.service.Mesh()
        self.mark_select = mesh_svc.ModeCompose (lx.symbol.sMARK_SELECT, None)

        # setup the primary mesh
        self.mesh = mesh[0]
        self.geom = self.mesh.geometry
        self.scene.select(self.mesh)

        # store all selected mesh elements
        selected = self.setup_mesh_elements()
        if not selected:
            self.selType = lx.symbol.iSEL_POLYGON

        # mesh item transform data
        pos = self.mesh.position.get()
        scl = self.mesh.scale.get()
        lx.out(f'Mesh position: {pos} scale: {scl}')

        locator = lx.object.Locator(self.mesh)
        chan_read = lx.object.ChannelRead(self.scene.Channels(None, 0.0))
        transforms = locator.LocalTransform(chan_read)

        quat = modo.Quaternion()
        quat.fromMatrix3(transforms[0])

        # mesh object data
        cobj = {
            'name': self.mesh.name,
            'type': 'MESH',
            'object_transform': {
                'translation': [pos[0], pos[1], pos[2]],
                'rotation_quat': [quat[0], quat[1], quat[2], quat[3]],
                'scale': [scl[0], scl[1], scl[2]]
            },
            'mesh': {}
        }

        # Query Existing Materials
        materials = self.copy_materials()
        if materials:
            cobj['mesh']['materials'] = materials

        # vertices
        positions = self.copy_vertices()
        if positions:
            cobj['mesh']['positions'] = positions

        # edges
        edges = self.copy_edges()
        if edges:
            cobj['mesh']['edges'] = edges

        # polygons
        polygons = self.copy_polygons()
        if polygons:
            cobj['mesh']['polygons'] = polygons

        # export UV maps
        uv_sets = self.extract_uv_sets()
        if uv_sets:
            cobj['mesh']['uv_sets'] = uv_sets

        # export morph maps
        shapekeys = self.extract_vertex_shapekeys()
        if shapekeys:
            cobj['mesh']['shapekeys'] = shapekeys

        # export weight maps
        vertex_groups = self.extract_vertex_groups()
        if vertex_groups:
            cobj['mesh']['vertex_groups'] = vertex_groups

        # export freestyle edges
        freestyle_edges = self.extract_edge_freestyle()
        if freestyle_edges:
            cobj['mesh']['freestyle_edges'] = freestyle_edges

        # CPMF data
        data = {
            'type': 'CPMF',
            'version': '1.0',
            'metadata': {
                'source_app': 'Modo',
                'coordinate_system': 'y_up_rh',
                'unit_scale': 1.0,
                'timestamp': datetime.utcnow().isoformat() + 'Z'
            },
            'objects': [cobj]
        }
        lx.out(f'Generated CPMF v1.0 data {external_clipboard}')

        try:
            txt = json.dumps(data, indent=4)
        except Exception as e:
            logging.error(f'Failed to dump JSON: {e}')
            return False

        # File
        if external_clipboard == 'tempfile':
            try:
                path = get_cpmf_tempfile_path()
                lx.out(f'Temporary file created at: {path}')
                written = write_tempfile(txt, path)
            except Exception as e:
                logging.error(f'Failed to write file: {e}')
                return False
        # Clipboard
        else:
            try:
                clipboard_copy(txt)
            except Exception as e:
                logging.error(f'Clipboard copy failed: {e}')
                # continue: maybe file write still possible

        lx.out('CPMF v1.0 export completed')
        return True