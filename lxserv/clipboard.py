'''
External Clipboard to exchange CPMF data between Modo and Blender

Copyright (C) 2025 Yoshiaki Tazaki All Rights Reserved

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
'''

import re
import lx
import lxu
import lxifc
import modo
import logging
import json
import sys
import subprocess
from datetime import datetime
import os
import tempfile
import pathlib
try:
    import msgpack
    use_msgpack = True
except ImportError:
    msgpack = None
    use_msgpack = False

i_VMAP_SEAM = 1397047629

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
    global use_msgpack
    temp_dir = tempfile.gettempdir()
    if use_msgpack:
        path = os.path.join(temp_dir, "cpmf_clipboard.bin")
    else:
        path = os.path.join(temp_dir, "cpmf_clipboard.json")
    return path

def write_tempfile(json, path=None):
    global use_msgpack
    """Write binary data to path if provided; otherwise create in OS tempdir and return path."""
    suffixes = [s.lower() for s in pathlib.Path(path).suffixes]
    use_binary = '.bin' in suffixes and use_msgpack
    d = os.path.dirname(path)
    if d and not os.path.exists(d):
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
    if use_binary:
        with open(path, 'wb') as f:
            f.write(msgpack.packb(json, use_bin_type=True))
    else:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(json)
    return os.path.abspath(path)

def read_tempfile(path):
    global use_msgpack
    suffixes = [s.lower() for s in pathlib.Path(path).suffixes]
    use_binary = '.bin' in suffixes and use_msgpack
    # try binary first
    if use_binary:
        if os.path.exists(path) == True:
            with open(path, 'rb') as f:
                return msgpack.unpackb(f.read(), raw=False)
    # then json
    elif '.json' in suffixes:
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        raise RuntimeError('Unsupported file format')

# ---------- (the other utility functions are the same as in the v1.5 code) ----------
# For brevity, core functions (coordinate conversion, material helpers, gatherers, etc.)
# are taken from the previous v1.5 implementation. We'll include them here verbatim
# so the addon is self-contained.
# ---------- Coordinate conversion helpers ----------
#
# Modo (Y-Up) -> Blender (Z-Up)
M2B = modo.Matrix3((
    (1.0, 0.0,  0.0), # x_rh
    (0.0, 0.0, -1.0), # z_rh
    (0.0, 1.0,  0.0), #-y_rh
))

# Blender (Z-Up) -> Modo (Y-Up)
B2M= modo.Matrix3((
    (1.0, 0.0,  0.0),
    (0.0, 0.0,  1.0),
    (0.0, -1.0, 0.0),
))

# Modo (Y-Up, Right Hand) -> LightWave (Y-Up, Left Hand)
M2L = modo.Matrix3((
    (1.0, 0.0,  0.0), # x_rh
    (0.0, 1.0,  0.0), # y_rh
    (0.0, 0.0, -1.0), #-z_rh
))
L2M = M2L.inverted()

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
    if "z_up_rh" in key:
        return mat3_mul_vec3(B2M, v)
    elif "y_up_lh" in key:
        return mat3_mul_vec3(L2M, v)

    # Already Modo space
    return v

# --- Quaternion conversion ---
def convert_quaternion_from_coord(q_in, src_coord):
    if not src_coord or "y_up_rh" in src_coord.lower():
        return modo.Quaternion((q_in[0], q_in[1], q_in[2], q_in[3]))

    # Quaternion → Matrix
    Rb = modo.Quaternion((q_in[0], q_in[1], q_in[2], q_in[3])).toMatrix3()

    # Blender rotation → Modo rotation
    if "z_up_rh" in src_coord.lower():
        Rm = B2M * Rb * M2B
    elif "y_up_lh" in src_coord.lower():
        Rm = L2M * Rb * L2B
    else:
        return modo.Quaternion((q_in[0], q_in[1], q_in[2], q_in[3]))

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
        self.vertex_ids = []
        self.edge_ids = []
        self.polygon_ids = []
        self.vmap_ids = []
        self.materials = []
        self.mark_select = None
        self.selType = None
        self.base_nvert = 0

    def selected(self, v):
        return v.TestMarks(self.mark_select)

    def selected_point(self, id):
        v = self.Point(id)
        return v.TestMarks(self.mark_select)

    def selected_polygonn(self, id):
        p = self.Polygon(id)
        return p.TestMarks(self.mark_select)

    def index(self, v):
        return self.vertex_indices[v.Index()]
    
    def lookupMap(self, map_type, name):
        for vmap_id in self.vmap_ids:
            vmap = self.VMap(vmap_id)
            try:
                vmap_name = vmap.Name()
            except:
                continue
            if vmap.Type() == map_type and vmap_name == name:
                return vmap
        return None
    
    def lookupMapAny(self, map_type):
        for vmap_id in self.vmap_ids:
            vmap = self.VMap(vmap_id)
            if vmap.Type() == map_type:
                return vmap
        return None

    def addMap(self, map_type, name):
        id = self.map_accessor.New(map_type, name)
        self.vmap_ids.append(id)
        return self.VMap(id)
    
    def newPoint(self, pos):
        id = self.point_accessor.New(pos)
        self.vertex_ids.append(id)
        return self.Point(id)
    
    def newPolygon(self, vertices):
        points_storage = lx.object.storage()
        points_storage.setType('p')
        points_storage.setSize(len(vertices))
        points_storage.set(vertices)
        id = self.polygon_accessor.New(lx.symbol.iPTYP_FACE, points_storage, len(vertices), 0)
        self.polygon_ids.append(id)
        return self.Polygon(id)

    def getWeight(self, vmap, v):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        if v.MapValue(vmap.ID(), storage) == False:
            return None
        return storage.get()

    def getPick(self, vmap, v):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        if v.MapValue(vmap.ID(), storage) == False:
            return False
        return True

    def getEdgePick(self, vmap, e):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        if e.MapValue(vmap.ID(), storage) == False:
            return False
        return True

    def getUV(self, vmap, p, point_id):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(2)
        if p.MapEvaluate(vmap.ID(), point_id, storage) == False:
            return [0.0, 0.0]
        return storage.get()

    def getColor(self, vmap, p, point_id):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(4)
        if p.MapEvaluate(vmap.ID(), point_id, storage) == False:
            return None
        rgba = storage.get()
        if vmap.Type() == lx.symbol.i_VMAP_RGB:
            return (rgba[0], rgba[1], rgba[2], 1.0)
        else:
            return rgba

    def getAbsolutePosition(self, vmap, v):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(3)
        if v.MapValue(vmap.ID(), storage) == False:
            return None
        if vmap.Type() == lx.symbol.i_VMAP_MORPH:
            pos = v.Pos()
            return [storage[0] + pos[0], storage[1] + pos[1], storage[2] + pos[2]]
        else:
            return [storage[0], storage[1], storage[2]]

    def setWeight(self, vmap, v, weight):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        storage.set([weight])
        v.SetMapValue(vmap.ID(), storage)

    def setMorph(self, vmap, v, pos):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(3)
        storage.set(pos)
        v.SetMapValue(vmap.ID(), storage)

    def setUV(self, vmap, p, point_id, uv):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(2)
        storage.set(uv)
        p.SetMapValue(point_id, vmap.ID(), storage)

    def setCornerColor(self, vmap, p, point_id, color):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(4)
        storage.set(color)
        p.SetMapValue(point_id, vmap.ID(), storage)

    def setPointColor(self, vmap, v, color):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(4)
        storage.set(color)
        v.SetMapValue(vmap.ID(), storage)

    def setEdgePick(self, vmap, e):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        e.SetMapValue(vmap.ID(), storage)

    def setVertexPick(self, vmap, v):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        v.SetMapValue(vmap.ID(), storage)

    def setSubdivWeight(self, vmap, e, weight):
        storage = lx.object.storage()
        storage.setType('f')
        storage.setSize(1)
        storage.set([weight])
        e.SetMapValue(vmap.ID(), storage)

    def PointByIndex(self, index):
        self.point_accessor.SelectByIndex(index)
        return self.point_accessor

    def PolygonByIndex(self, index):
        self.polygon_accessor.SelectByIndex(index)
        return self.polygon_accessor

    def EdgeByIndex(self, index):
        self.edge_accessor.SelectByIndex(index)
        return self.edge_accessor

    def EdgeEndpoints(self, id0, id1):
        self.edge_accessor.SelectEndpoints(id0, id1)
        return self.edge_accessor

    def Point(self, id):
        self.point_accessor.Select(id)
        return self.point_accessor

    def Edge(self, id):
        self.edge_accessor.Select(id)
        return self.edge_accessor

    def Polygon(self, id):
        self.polygon_accessor.Select(id)
        return self.polygon_accessor

    def VMap(self, id):
        self.map_accessor.Select(id)
        return self.map_accessor

    def MaterialTag(self, p):
        loc = lx.object.StringTag(p)
        try:
            tag = loc.Get(lx.symbol.i_POLYTAG_MATERIAL)
        except:
            tag = None
        return tag

    def setMaterialTag(self, p, tag):
        loc = lx.object.StringTag(p)
        return loc.Set(lx.symbol.i_POLYTAG_MATERIAL, tag)

    def PickTag(self, p):
        loc = lx.object.StringTag(p)
        try:
            tag = loc.Get(lx.symbol.i_POLYTAG_PICK)
        except:
            tag = None
        return tag

    def setPickTag(self, p, tag):
        loc = lx.object.StringTag(p)
        return loc.Set(lx.symbol.i_POLYTAG_PICK, tag)

    def setup_mesh_elements(self):
        # store selected vertices
        self.vertex_indices = {}
        self.vertex_ids = []
        index = 0
        for i in range(self.mesh.PointCount()):
            v = self.PointByIndex(i)
            if self.selected(v):
                self.vertex_indices[i] = index
                self.vertex_ids.append(v.ID())
                index += 1

        # store selected edges
        self.edge_ids = []
        for i in range(self.mesh.EdgeCount()):
            e = self.EdgeByIndex(i)
            id0, id1 = e.Endpoints()
            if self.selected_point(id0) and self.selected_point(id1):
                self.edge_ids.append(e.ID())

        # store selected polygons
        self.polygon_ids = []
        for i in range(self.mesh.PolygonCount()):
            p = self.PolygonByIndex(i)
            ptype = p.Type()
            if ptype != lx.symbol.iPTYP_FACE and \
               ptype != lx.symbol.iPTYP_SUBD and \
               ptype != lx.symbol.iPTYP_PSUB:
                continue
            if self.selected(p):
                self.polygon_ids.append(p.ID())

        if self.selType == lx.symbol.iSEL_VERTEX:
            return True if len(self.vertex_ids) > 0 else False
        elif self.selType == lx.symbol.iSEL_EDGE:
            return True if len(self.edge_ids) > 0 else False
        else:
            return True if len(self.polygon_ids) > 0 else False

    def setup_vmap_ids(self):
        # store all vertex maps
        class QueryMapsVisitor(lxifc.Visitor):
            def __init__(self, meshmap, vmap_ids):
                self.meshmap = meshmap
                self.msv = lx.service.Mesh()
                self.vmap_ids = vmap_ids

            def vis_Evaluate(self):
                try:
                    id = self.meshmap.ID()
                    self.vmap_ids.append(id)
                except Exception:
                    return

        self.vmap_ids = []
        self.vmap_uv_ids = []
        self.vmap_morph_ids = []
        self.vmap_weight_ids = []
        self.vmap_color_ids = []
        visitor = QueryMapsVisitor(self.map_accessor, self.vmap_ids)
        self.map_accessor.Enumerate(lx.symbol.iMARK_ANY, visitor, 0)

        for vmap_id in self.vmap_ids:
            vmap = self.VMap(vmap_id)
            if vmap.Type() == lx.symbol.i_VMAP_TEXTUREUV:
                self.vmap_uv_ids.append(vmap_id)
            elif vmap.Type() == lx.symbol.i_VMAP_MORPH:
                self.vmap_morph_ids.append(vmap_id)
            elif vmap.Type() == lx.symbol.i_VMAP_SPOT:
                self.vmap_morph_ids.append(vmap_id)
            elif vmap.Type() == lx.symbol.i_VMAP_WEIGHT:
                self.vmap_weight_ids.append(vmap_id)
            elif vmap.Type() == lx.symbol.i_VMAP_RGB:
                self.vmap_color_ids.append(vmap_id)
            elif vmap.Type() == lx.symbol.i_VMAP_RGBA:
                self.vmap_color_ids.append(vmap_id)
            else:
                continue

    # check if face winding needs to be reversed
    def reverse_face_winding(self):
        rev = True if self.data.get('metadata', {}).get('coordinate_system', '').lower().find('lh') >= 0 else False
        return rev

    # Main copy function
    def copy(self, external_clipboard='tempfile'):
        lx.out(f'Copying to external clipboard: {external_clipboard}')

        self.scene = modo.Scene()

        layer_svc = lx.service.Layer()
        layer_scan = lx.object.LayerScan(layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_ACTIVE | lx.symbol.f_LAYERSCAN_MARKALL))
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
            'objects': []
        }

        items = []
        for layer_index in range(layer_scan.Count()):
            items.append(layer_scan.MeshItem (layer_index))

        for layer_index in range(layer_scan.Count()):
            # setup the accessor
            self.item = lx.object.Item (layer_scan.MeshItem (layer_index))
            self.mesh = lx.object.Mesh (layer_scan.MeshBase (layer_index))
            self.edge_accessor = lx.object.Edge (self.mesh.EdgeAccessor ())
            self.point_accessor = lx.object.Point (self.mesh.PointAccessor ())
            self.polygon_accessor = lx.object.Polygon (self.mesh.PolygonAccessor ())
            self.map_accessor = lx.object.MeshMap (self.mesh.MeshMapAccessor ())

            matrix = modo.Matrix4(layer_scan.MeshTransform (layer_index))
            pos = matrix.position
            scl = (matrix[0][0], matrix[1][1], matrix[2][2])
            quat = modo.Quaternion()
            quat.fromMatrix3(matrix)

            # store all selected mesh elements
            selected = self.setup_mesh_elements()
            if not selected:
                self.selType = lx.symbol.iSEL_POLYGON

            # store all vertex maps
            self.setup_vmap_ids()

            # mesh object data
            cobj = {
                'name': self.item.UniqueName(),
                'type': 'MESH',
                'object_transform': {
                    'translation': [pos[0], pos[1], pos[2]],
                    'rotation_quat': [quat[0], quat[1], quat[2], quat[3]],
                    'scale': [scl[0], scl[1], scl[2]]
                },
            }

            parent = self.get_item_parent(self.item, items)
            if parent is not None:
                cobj['parent'] = parent

            cobj['mesh'] = {}

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
            uv_sets = self.copy_uv_sets()
            if uv_sets:
                cobj['mesh']['uv_sets'] = uv_sets

            # export morph maps
            shapekeys = self.copy_vertex_shapekeys()
            if shapekeys:
                cobj['mesh']['shapekeys'] = shapekeys

            # export weight maps
            vertex_groups = self.copy_vertex_groups()
            if vertex_groups:
                cobj['mesh']['vertex_groups'] = vertex_groups

            # export freestyle edges
            freestyle_edges = self.copy_edge_freestyle()
            if freestyle_edges:
                cobj['mesh']['freestyle_edges'] = freestyle_edges

            # export RGBA maps
            colors = self.copy_colors()
            if colors:
                cobj['mesh']['colors'] = colors

            # export selection sets
            selection_sets = self.copy_selection_sets()
            if selection_sets:
                cobj['mesh']['selection_sets'] = selection_sets

            data['objects'].append(cobj)

        lx.out(f'Generated CPMF v1.0 data {external_clipboard}')

        # JSON dump
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
                write_tempfile(txt, path)
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
    
    # get item parent index
    def get_item_parent(self, item, items):
        modo_item = modo.Item(item)
        parent = modo_item.parent
        if parent == None:
            return None
        for i, ref in enumerate(items):
            if ref == item:
                continue
            modo_ref = modo.Item(ref)
            if modo_ref.name == parent.name:
                return i
        return None

    def get_material_index(self, name, material_out):
        for i, mat in enumerate(material_out):
            if mat['name'] == name:
                return i
        return 0

    def copy_uv_sets(self):
        if len(self.polygon_ids) == 0:
            return None
        uv_sets = []
        for vmap_id in self.vmap_uv_ids:
            vmap = self.VMap(vmap_id)
            uv_set = {
                'name': vmap.Name(),
                'uvs': []
            }
            i = 0
            for poly_id in self.polygon_ids:
                p = self.Polygon(poly_id)
                if self.is_keyhole(p):
                    count = p.GenerateTriangles()
                    for j in range(count):
                        face_uvs = {
                            'index': i,
                            'values': []
                        }
                        i += 1
                        for point_id in p.TriangleByIndex(j):
                            uv = self.getUV(vmap, p, point_id)
                            face_uvs['values'].append([uv[0], uv[1]])
                        uv_set['uvs'].append(face_uvs)
                else:
                    face_uvs = {
                        'index': i,
                        'values': []
                    }
                    i += 1
                    for j in range(p.VertexCount()):
                        point_id = p.VertexByIndex(j)
                        uv = self.getUV(vmap, p, point_id)
                        face_uvs['values'].append([uv[0], uv[1]])
                    uv_set['uvs'].append(face_uvs)
            uv_sets.append(uv_set)
        if len(uv_sets) == 0:
            return None
        return uv_sets

    def copy_colors(self):
        if len(self.polygon_ids) == 0:
            return None
        colors = []
        for vmap_id in self.vmap_color_ids:
            vmap = self.VMap(vmap_id)
            color = {
                'name': vmap.Name(),
                'domain': 'CORNER',
                'data_type': 'FLOAT_COLOR',
                'colors': []
            }
            i = 0
            for poly_id in self.polygon_ids:
                p = self.Polygon(poly_id)
                if self.is_keyhole(p):
                    count = p.GenerateTriangles()
                    for j in range(count):
                        face_color = {
                            'index': i,
                            'values': []
                        }
                        i += 1
                        n = 0
                        for point_id in p.TriangleByIndex(j):
                            rgba = self.getColor(vmap, p, point_id)
                            if rgba is None:
                                face_color['values'].append([0.0, 0.0, 0.0, 0.0])
                            else:
                                face_color['values'].append([rgba[0], rgba[1], rgba[2], rgba[3]])
                                n += 1
                        if n > 0:
                            color['colors'].append(face_color)
                else:
                    face_color = {
                        'index': i,
                        'values': []
                    }
                    i += 1
                    n = 0
                    for j in range(p.VertexCount()):
                        point_id = p.VertexByIndex(j)
                        rgba = self.getColor(vmap, p, point_id)
                        if rgba is None:
                            face_color['values'].append([0.0, 0.0, 0.0, 0.0])
                        else:
                            face_color['values'].append([rgba[0], rgba[1], rgba[2], rgba[3]])
                            n += 1
                    if n > 0:
                        color['colors'].append(face_color)
            colors.append(color)
        if len(colors) == 0:
            return None
        return colors

    def copy_vertex_groups(self):
        if len(self.vertex_ids) == 0:
            return None
        vertex_groups = []
        for vmap_id in self.vmap_weight_ids:
            vmap = self.VMap(vmap_id)
            vg_data = {
                'name': vmap.Name(),
                'weights': []
            }
            for point_id in self.vertex_ids:
                v = self.Point(point_id)
                w = self.getWeight(vmap, v)
                if w is not None and w[0] != 0.0:
                    vg_data['weights'].append({'index': self.index(v), 'weight': w[0]})
            vertex_groups.append(vg_data)
        if len(vertex_groups) == 0:
            return None
        return vertex_groups

    def copy_vertex_shapekeys(self):
        if len(self.vmap_morph_ids) == 0:
            return None
        if len(self.vertex_ids) == 0:
            return None
        shapekeys = []
        # Add Basis positions
        sk_data = {
            'name': 'Basis',
            'relative': True,
            'positions': []
        }
        for point_id in self.vertex_ids:
            v = self.Point(point_id)
            pos = v.Pos()
            sk_data['positions'].append({'index': self.index(v), 'position':[pos[0], pos[1], pos[2]]})
        shapekeys.append(sk_data)
        # Add all morph and spot vertex maps
        for vmap_id in self.vmap_morph_ids:
            vmap = self.VMap(vmap_id)
            if vmap.Type() == lx.symbol.i_VMAP_SPOT:
                relative = False
            elif vmap.Type() == lx.symbol.i_VMAP_MORPH:
                relative = True
            else:
                continue
            sk_data = {
                'name': vmap.Name(),
                'relative': relative,
                'positions': []
            }
            for point_id in self.vertex_ids:
                v = self.Point(point_id)
                co = self.getAbsolutePosition(vmap, v)
                if co is None:
                    continue
                sk_data['positions'].append({'index': self.index(v), 'position':[co[0], co[1], co[2]]})
            shapekeys.append(sk_data)
        if len(shapekeys) == 0:
            return None
        return shapekeys

    def copy_edge_freestyle(self):
        if len(self.edge_ids) == 0:
            return None
        freestyle_edges = []
        for vmap_id in self.vmap_ids:
            vmap = self.VMap(vmap_id)
            if vmap.Type() != lx.symbol.i_VMAP_EPCK:
                continue
            if vmap.Name() != '_Freestyle':
                continue
            for edge_id in self.edge_ids:
                e = self.Edge(edge_id)
                storageBuffer = lx.object.storage('f', 1)
                if e.MapEvaluate(vmap_id, storageBuffer) == True:
                    id0, id1 = e.Endpoints()
                    freestyle_edges.append({'vertices': [self.index(self.Point(id0)), self.index(self.Point(id1))], 'use_freestyle_mark': 1})
        if len(freestyle_edges) == 0:
            return None
        return freestyle_edges

    def copy_selection_sets(self):
        if len(self.polygon_ids) == 0:
            return None
        selection_sets = []
        for vmap_id in self.vmap_ids:
            vmap = self.VMap(vmap_id)
            # Vertex selection set
            if vmap.Type() == lx.symbol.i_VMAP_PICK:
                sset = {
                    'name': vmap.Name(),
                    'type': 'VERT',
                    'indices': []
                }
                for i, vertex_id in enumerate(self.vertex_ids):
                    v = self.Point(vertex_id)
                    if self.getPick(vmap, v) == True:
                        sset['indices'].append(i)
                selection_sets.append(sset)
            # Edge selection set
            elif vmap.Type() == lx.symbol.i_VMAP_EPCK:
                sset = {
                    'name': vmap.Name(),
                    'type': 'EDGE',
                    'indices': []
                }
                for edge_id in self.edge_ids:
                    e = self.Edge(edge_id)
                    if self.getEdgePick(vmap, e) == True:
                        id0, id1 = e.Endpoints()
                        sset['indices'].append([self.index(self.Point(id0)), self.index(self.Point(id1))])
                selection_sets.append(sset)
        # Polygon selection set
        poly_sset = {}
        i = 0
        for poly_id in self.polygon_ids:
            p = self.Polygon(poly_id)
            if self.is_keyhole(p):
                count = p.GenerateTriangles()
            else:
                count = 1
            tagString = self.PickTag(p)
            if tagString is None:
                i += count
                continue
            tags = tagString.split(";")
            for j in range(count):
                for tag in tags:
                    if tag not in poly_sset:
                        poly_sset[tag] = [i]
                    else:
                        poly_sset[tag].append(i)
                i += 1
        for tag, indices in poly_sset.items():
            sset = {
                'name': tag,
                'type': 'FACE',
                'indices': indices
            }
            selection_sets.append(sset)
                
        if len(selection_sets) == 0:
            return None
        return selection_sets
    
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
    def copy_textures(self, material):
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
        if len(self.polygon_ids) == 0:
            return None
        use_mask_all = True
        for id in self.polygon_ids:
            p = self.Polygon(id)
            if self.MaterialTag(p) is None:
                use_mask_all = False
                break
        # Query Existing Materials
        self.materials = []
        for material in self.scene.items("advancedMaterial"):
            mask = material.parent
            if not mask or mask.type != 'mask':
                if use_mask_all:
                    continue
                name = 'Default'
            else:
                name = mask.channel('ptag').get()
            diffCol = material.channel('diffCol').get()

            mat_data = {
                'name': name,
                'base_color': diffCol
            }
            # extract texture data for the material
            textures = self.copy_textures(material)
            if textures is not None:
                mat_data['textures'] = textures
            self.materials.append(mat_data)
        if len(self.materials) == 0:
            return None
        return self.materials
    
    # copy all selected vertices
    def copy_vertices(self):
        if len(self.vertex_ids) == 0:
            return None
        positions = []
        for id in self.vertex_ids:
            v = self.Point(id)
            pos = v.Pos()
            positions.append([pos[0], pos[1], pos[2]])
        if len(positions) == 0:
            return None
        return positions

    # copy all selected edges
    def copy_edges(self):
        if len(self.polygon_ids) == 0:
            return None
        if len(self.edge_ids) == 0:
            return None
        id_subdiv = None
        id_seam = None
        id_seam_any = None
        for id in self.vmap_ids:
            vmap = self.VMap(id)
            if vmap.Type() == lx.symbol.i_VMAP_SUBDIV:
                id_subdiv = id
            elif vmap.Type() == i_VMAP_SEAM:
                if vmap.Name() == '_Seam':
                    id_seam = id
                id_seam_any = id
        if id_seam is None:
            id_seam = id_seam_any
        # crease edges
        crease_edges = [0.0] * self.mesh.EdgeCount()
        if id_subdiv is not None:
            storageBuffer = lx.object.storage('f', 1)
            for i in range(self.mesh.EdgeCount()):
                e = self.EdgeByIndex(i)
                if e.MapEvaluate(id_subdiv, storageBuffer) == True:
                    w = storageBuffer.get()
                    crease_edges[i] = w[0]
        # uv seam edges
        seam_edges = [False] * self.mesh.EdgeCount()
        if id_seam is not None:
            storageBuffer = lx.object.storage('f', 1)
            for i in range(self.mesh.EdgeCount()):
                e = self.EdgeByIndex(i)
                if e.MapEvaluate(id_seam, storageBuffer) == True:
                    seam_edges[i] = True
            
        # edges
        edges = []
        for i in range(self.mesh.EdgeCount()):
            e = self.EdgeByIndex(i)
            if crease_edges[i] == 0.0 and seam_edges[i] == False:
                continue
            id0, id1 = e.Endpoints()
            if self.selected_point(id0) and self.selected_point(id1):
                edges.append({
                    'vertices': [self.index(self.Point(id0)), self.index(self.Point(id1))],
                    'attributes': {
                        'crease_edge': crease_edges[i],
                        'seam': seam_edges[i]
                    }
                })
        if len(edges) == 0:
            return None
        return edges

    # copy all selected polygons
    def copy_polygons(self):
        if len(self.polygon_ids) == 0:
            return None
        polygons = []
        for id in self.polygon_ids:
            p = self.Polygon(id)
            p_attrs = {
                'material_index': self.get_material_index(self.MaterialTag(p), self.materials)
            }
            if self.is_keyhole(p):
                count = p.GenerateTriangles()
                for i in range(count):
                    vertices = []
                    for id in p.TriangleByIndex(i):
                        v = self.Point(id)
                        vertices.append(self.index(v))
                    polygons.append({
                        'vertices': vertices,
                        'attributes': p_attrs
                    })
            else:
                vertices = []
                for i in range(p.VertexCount()):
                    v = self.Point(p.VertexByIndex(i))
                    vertices.append(self.index(v))
                polygons.append({
                    'vertices': vertices,
                    'attributes': p_attrs
                })
        if len(polygons) == 0:
            return None
        return polygons

    # check if the polygon has bridge edge to connect between outer and an inner loop
    def is_keyhole(self, p):
        count = p.VertexCount()
        if count < 8:
            return False
        edge_map = {}
        for i in range(count):
            id0 = p.VertexByIndex(i)
            id1 = p.VertexByIndex((i + 1) % count)
            if i > 2:
                if (id1, id0) in edge_map:
                    return True
            edge_map[(id0, id1)] = True
        return False

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
            materials = mesh_data.get('materials', [])
            uv_sets = mesh_data.get('uv_sets', [])
            shapekeys = mesh_data.get('shapekeys', [])
            vertex_groups = mesh_data.get('vertex_groups', [])
            freestyle_edges = mesh_data.get('freestyle_edges', [])
            colors = mesh_data.get('colors', [])
            selection_sets = mesh_data.get('selection_sets', [])

            if new_mesh == True:
                mesh = self.scene.addMesh(obj_data['name'])
                self.scene.select(mesh)

            layer_svc = lx.service.Layer()
            scan1 = layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_EDIT | lx.symbol.f_LAYERSCAN_PRIMARY)
            if scan1.test() == False:
                return

            self.item = lx.object.Item (scan1.MeshItem (0))
            self.mesh = lx.object.Mesh (scan1.MeshEdit (0))
            self.edge_accessor = lx.object.Edge (self.mesh.EdgeAccessor ())
            self.point_accessor = lx.object.Point (self.mesh.PointAccessor ())
            self.polygon_accessor = lx.object.Polygon (self.mesh.PolygonAccessor ())
            self.map_accessor = lx.object.MeshMap (self.mesh.MeshMapAccessor ())

            # store all vertex maps
            self.setup_vmap_ids()
        
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

            # paste vertex shapekeys data to geometry
            if colors:
                self.paste_colors(colors)

            #scan1.Update()

            scan1.SetMeshChange(0, lx.symbol.f_MESHEDIT_GEOMETRY)
            scan1.Apply()
            scan1 = None

            # Apply edge vertex map values using layer scan since MeshGetPolyEdge 
            # was crashed at Endpoints method
            scan2 = layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_EDIT | lx.symbol.f_LAYERSCAN_PRIMARY)
            if scan2.test() == False:
                return

            self.item = lx.object.Item (scan2.MeshItem (0))
            self.mesh = lx.object.Mesh (scan2.MeshEdit (0))
            self.edge_accessor = lx.object.Edge (self.mesh.EdgeAccessor ())
            self.point_accessor = lx.object.Point (self.mesh.PointAccessor ())
            self.polygon_accessor = lx.object.Polygon (self.mesh.PolygonAccessor ())
            self.map_accessor = lx.object.MeshMap (self.mesh.MeshMapAccessor ())

            # paste edges data to geometry
            if edges:
                self.paste_edges(edges)

            # paste freestyle data to geometry as EdgePick map
            if freestyle_edges:
                self.paste_edge_freestyle(freestyle_edges)

            # paste selection sets
            if selection_sets:
                self.paste_selection_sets(selection_sets)

            scan2.SetMeshChange(0, lx.symbol.f_MESHEDIT_MAP_OTHER|lx.symbol.f_MESHEDIT_POL_TAGS)
            scan2.Apply()
            scan2 = None

    def paste_vertices(self, positions):
        coord = self.data.get('metadata', {}).get('coordinate_system', '').lower()
        unit_scale = float(self.data.get('metadata', {}).get('unit_scale', 1.0))
        self.vertex_ids = []
        self.base_nvert = self.mesh.PointCount()
        for i, p in enumerate(positions):
            v = convert_vector_from_coord(p, coord)
            v *= unit_scale
            id = self.newPoint((v.x, v.y, v.z))

    def paste_polygons(self, polygons, materials):
        rev = self.reverse_face_winding()
        count = 0
        self.polygon_ids = []
        for poly in polygons:
            vert_indices = poly.get('vertices', [])
            if rev:
                vert_indices.reverse()
            face_verts = [self.vertex_ids[i] for i in vert_indices]
            p = self.newPolygon(face_verts)
            attributes = poly.get('attributes', {})
            if 'material_index' in attributes:
                try:
                    material_index = int(attributes['material_index'])
                    if material_index is not None and 0 <= material_index < len(materials):
                        materialTag = materials[material_index].get('name', '')
                        self.setMaterialTag(p, materialTag)
                except Exception:
                    pass
            count += 1

    def select_edge(self, vertices):
        i0, i1 = vertices[0], vertices[1]
        if i0 >= len(self.vertex_ids) or i1 >= len(self.vertex_ids):
            return None
        id0 = self.vertex_ids[i0]
        id1 = self.vertex_ids[i1]
        if id0 is None or id1 is None:
            return None
        self.edge_accessor.SelectEndpoints(id0, id1)
        return self.edge_accessor
    
    def paste_edges(self, edges):
        # Subdivision map
        vmap = self.lookupMap(lx.symbol.i_VMAP_SUBDIV, "Subdivision")
        if vmap != None:
            for edge in edges:
                e = self.select_edge(edge.get('vertices', []))
                if e is None:
                    continue
                weight = edge.get('attributes', {}).get('crease_edge', 0.0)
                if weight > 0.0:
                    self.setSubdivWeight(vmap, e, weight)
        # UV Seam map
        vmap = self.lookupMap(i_VMAP_SEAM, "_Seam")
        for edge in edges:
            e = self.select_edge(edge.get('vertices', []))
            if e is None:
                continue
            seam = edge.get('attributes', {}).get('seam', False)
            if seam is True:
                if vmap is None:
                    vmap = self.addMap(i_VMAP_SEAM, "_Seam")
                self.setEdgePick(vmap, e)

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
        rev = self.reverse_face_winding()
        for uv_set in uv_sets:
            name = uv_set.get('name', '')
            vmap = self.lookupMap(lx.symbol.i_VMAP_TEXTUREUV, name)
            if not vmap:
                vmap = self.addMap(lx.symbol.i_VMAP_TEXTUREUV, name)
            for face_uvs in uv_set.get('uvs', []):
                index = face_uvs.get('index')
                values = face_uvs.get('values', [])
                if rev:
                    values.reverse()
                poly_id = self.polygon_ids[index]
                p = self.Polygon(poly_id)
                for i in range(p.VertexCount()):
                    point_id = p.VertexByIndex(i)
                    uv = values[i]
                    self.setUV(vmap, p, point_id, uv)


    def paste_colors(self, colors):
        rev = self.reverse_face_winding()
        for color in colors:
            name = color.get('name', '')
            domain = color.get('domain', '')
            vmap = self.lookupMap(lx.symbol.i_VMAP_RGBA, name)
            if not vmap:
                vmap = self.addMap(lx.symbol.i_VMAP_RGBA, name)
            if domain == 'CORNER':
                for face_color in color.get('colors', []):
                    index = face_color.get('index')
                    values = face_color.get('values', [])
                    if rev:
                        values.reverse()
                    poly_id = self.polygon_ids[index]
                    p = self.Polygon(poly_id)
                    for i in range(p.VertexCount()):
                        point_id = p.VertexByIndex(i)
                        color = values[i]
                        self.setCornerColor(vmap, p, point_id, color)
            elif domain == 'POINT':
                for point_color in color.get('colors', []):
                    index = point_color.get('index')
                    color = point_color.get('values', [])
                    point_id = self.vertex_ids[index]
                    v = self.Point(point_id)
                    self.setPointColor(vmap, v, color)

    def paste_vertex_groups(self, vertex_groups):
        for vertex_group in vertex_groups:
            name = vertex_group.get('name', '')
            weights = vertex_group.get('weights', [])
            vmap = self.lookupMap(lx.symbol.i_VMAP_WEIGHT, name)
            if not vmap:
                vmap = self.addMap(lx.symbol.i_VMAP_WEIGHT, name)
            for w_data in weights:
                index = w_data.get('index')
                weight = w_data.get('weight', 0.0)
                v = self.Point(self.vertex_ids[index])
                self.setWeight(vmap, v, weight)

    def paste_vertex_shapekeys(self, shapekeys):
        coord = self.data.get('metadata', {}).get('coordinate_system', '').lower()
        unit_scale = float(self.data.get('metadata', {}).get('unit_scale', 1.0))
        base_positions = None
        for shapekey in shapekeys:
            name = shapekey.get('name')
            lx.out(f"shapekey {name} {name.lower()}")
            if name.lower() == 'basis':
                base_positions = shapekey.get('positions', [])
                continue
            use_relative = shapekey.get('relative', True)
            map_type = lx.symbol.i_VMAP_MORPH if use_relative else lx.symbol.i_VMAP_SPOT
            vmap = self.lookupMap(lx.symbol.i_VMAP_WEIGHT, name)
            if not vmap:
                vmap = self.addMap(map_type, name)
            for pos_data in shapekey.get('positions', []):
                index = pos_data.get('index')
                pos = modo.Vector3(pos_data.get('position'))
                pos = convert_vector_from_coord(pos, coord)
                pos *= unit_scale
                v = self.Point(self.vertex_ids[index])
                if use_relative:
                    if base_positions is None:
                        pos -= v.Pos()
                    else:
                        base_data = base_positions[index]
                        base_pos = modo.Vector3(base_data.get('position'))
                        base_pos = convert_vector_from_coord(base_pos, coord)
                        base_pos *= unit_scale
                        pos -= base_pos
                self.setMorph(vmap, v, pos)

    def paste_edge_freestyle(self, freestyle_edges):
        name = '_Freestyle'
        vmap = self.lookupMap(lx.symbol.i_VMAP_EPCK, name)
        if not vmap:
            vmap = self.addMap(lx.symbol.i_VMAP_EPCK, name)
        for data in freestyle_edges:
            vertices = data.get('vertices')
            use_freestyle_mark = data.get('use_freestyle_mark', 0)
            if use_freestyle_mark:
                e = self.select_edge(vertices)
                if e is None:
                    continue
                self.setEdgePick(vmap, e)

    def paste_selection_sets(self, selection_sets):
        for data in selection_sets:
            name = data.get('name')
            type = data.get('type')
            indices = data.get('indices')
            if type == 'VERT':
                vmap = self.lookupMap(lx.symbol.i_VMAP_PICK, name)
                if not vmap:
                    vmap = self.addMap(lx.symbol.i_VMAP_PICK, name)
                for index in indices:
                    v = self.Point(self.vertex_ids[index])
                    self.setVertexPick(vmap, v)
            elif type == 'EDGE':
                vmap = self.lookupMap(lx.symbol.i_VMAP_EPCK, name)
                if not vmap:
                    vmap = self.addMap(lx.symbol.i_VMAP_EPCK, name)
                for vertices in indices:
                    e = self.select_edge(vertices)
                    if e is None:
                        continue
                    self.setEdgePick(vmap, e)
            elif type == 'FACE':
                for index in indices:
                    p = self.Polygon(self.polygon_ids[index])
                    self.setPickTag(p, name)