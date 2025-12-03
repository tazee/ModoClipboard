#
# External Clipboard Copy and Paste for Modo
#

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
        self.vertex_indices = {}
        self.mark_select = None
        self.selType = None

    def selected(self, v):
        return v.accessor.TestMarks(self.mark_select)

    def vidx(self, v):
        return self.vertex_indices[v]

    def setup_mesh_elements(self):
        # store selected vertices
        self.vertex_indices = {}
        self.vertices.clear()
        index = 0
        for v in self.mesh.geometry.vertices:
            if self.selected(v):
                self.vertex_indices[v] = index
                self.vertices.append(v)
                index += 1

        # store selected edges
        self.edges.clear()
        for e in self.mesh.geometry.edges:
            v0, v1 = e.vertices
            if self.selected(v0) and self.selected(v1):
                self.edges.append(e)

        # store selected polygons
        self.edges.clear()
        for p in self.mesh.geometry.polygons:
            if self.selected(p):
                self.polygons.append(p)

        if self.selType == lx.symbol.iSEL_VERTEX:
            return True if len(self.vertices) > 0 else False
        elif self.selType == lx.symbol.iSEL_EDGE:
            return True if len(self.edges) > 0 else False
        else:
            return True if len(self.polygons) > 0 else False


    # Main paste function
    def paste(self, external_clipboard='CLIPBOARD', new_mesh=False):
        print(f'Pasting from external clipboard: {external_clipboard}, new_mesh={new_mesh}')
        if external_clipboard == 'TEMPFILE':
            path = get_cpmf_tempfile_path()
            print(f'Read file from: {path}')
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
            data = json.loads(txt)
        except Exception as e:
            lx.out({'ERROR'}, f'Invalid JSON: {e}')
            return False
        
        # same import logic as before (coordinate conversion, materials, UVs, groups, shape keys...)
        coord = data.get('metadata', {}).get('coordinate_system', '').lower()
        unit_scale = float(data.get('metadata', {}).get('unit_scale', 1.0))
        base_dir = data.get('metadata', {}).get('custom', {}).get('base_dir', None)
        print(f'Coordinate system: {coord}, unit scale: {unit_scale}, base_dir: {base_dir}')
        
        # Add a new mesh object to the scene and grab the geometry object
        scene = modo.Scene()

        for obj_data in data.get('objects', []):
            mesh_data = obj_data.get('mesh', {})
            positions = mesh_data.get('positions', [])
            edges = mesh_data.get('edges', [])
            polygons = mesh_data.get('polygons', [])
            materials = mesh_data.get('materials', []) or mesh_data.get('mesh', {}).get('materials', []) or mesh_data.get('materials', [])

            if new_mesh == True:
                mesh = scene.addMesh("Mesh")
            else:
                mesh = scene.selectedByType("mesh")
                if mesh:
                    mesh = mesh[0]
                    scene.select(mesh)
                else:
                    lx.out({'ERROR'}, 'No mesh selected to paste into')
                    return False

            self.mesh = mesh
            geo = mesh.geometry
        
            # convert positions and apply unit scale
            vertices = []
            for p in positions:
                v = convert_vector_from_coord(p, coord)
                v *= unit_scale
                print(f'pos: {p} v : {v}')
                vertices.append(geo.vertices.new((v.x, v.y, v.z)))

            count = 0
            for poly in polygons:
                vert_indices = poly.get('vertices', [])
                face_verts = [vertices[i] for i in vert_indices]
                geo.polygons.new(face_verts)
                attributes = poly.get('attributes', {})
                if 'material_index' in attributes:
                    try:
                        material_index = int(attributes['material_index'])
                        if material_index is not None and 0 <= material_index < len(materials):
                            geo.polygons[count].materialTag = materials[material_index].get('name', '')
                    except Exception:
                        pass
                count += 1

            for material in materials:
                name = material.get('name', '')
                col = material.get('base_color', ())
                mat = scene.addMaterial(name='M_' + name)
                mat.channel('diffCol').set(col)
                mask = scene.addItem('mask', name=name)
                mask.channel('ptag').set(name)
                mat.setParent(mask, index=1)

            geo.setMeshEdits()

    def get_material_index(self, name, material_out):
        for i, mat in enumerate(material_out):
            if mat['name'] == name:
                return i
        return 0

    def extract_uv_sets(self):
        uv_sets = []
        for vmap in self.mesh.geometry.vmaps.uvMaps:
            uv_set = {
                'name': vmap.name,
                'uvs': []
            }
            for p in self.mesh.geometry.polygons:
                face_uvs = {
                    'index': p.index,
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
        vertex_groups = []
        for vmap in self.mesh.geometry.vmaps:
            if vmap.map_type != lx.symbol.i_VMAP_WEIGHT:
                continue
            vg_data = {
                'name': vmap.name,
                'weights': []
            }
            for v in self.mesh.geometry.vertices:
                w = vmap[v.index]
                if w is not None:
                    vg_data['weights'].append({'index': v.index, 'weight': w[0]})
            vertex_groups.append(vg_data)
            print(f"-- weight map {vmap.name} size {len(vertex_groups)}")
        if len(vertex_groups) == 0:
            return None
        return vertex_groups

    def extract_vertex_shapekeys(self):
        if len(self.mesh.geometry.vmaps.morphMaps) == 0:
            return None
        shapekeys = []
        # Add Basis positions
        sk_data = {
            'name': 'Basis',
            'relative': True,
            'positions': []
        }
        for v in self.mesh.geometry.vertices:
            sk_data['positions'].append({'index': v.index, 'position':[v.position[0], v.position[1], v.position[2]]})
        shapekeys.append(sk_data)
        # Add all morph and spot vertex maps
        for vmap in self.mesh.geometry.vmaps.morphMaps:
            relative = True if (vmap.map_type == lx.symbol.i_VMAP_MORPH) else False
            sk_data = {
                'name': vmap.name,
                'relative': relative,
                'positions': []
            }
            for v in self.mesh.geometry.vertices:
                co = vmap.getAbsolutePosition(v.index)
                sk_data['positions'].append({'index': v.index, 'position':[co[0], co[1], co[2]]})
            shapekeys.append(sk_data)
            print(f"-- morph map {vmap.name} relative {relative} map_type {vmap.map_type}")
        if len(shapekeys) == 0:
            return None
        return shapekeys

    def get_selected_vertices(self):
        selSrv = lx.service.Selection()
        vertSelTypeCode = selSrv.LookupType(lx.symbol.sSELTYP_VERTEX)
        vTransPacket = lx.object.VertexPacketTranslation(selSrv.Allocate(lx.symbol.sSELTYP_VERTEX))

        numVerts = selSrv.Count(vertSelTypeCode)
        vertex_ids = []

        print(f"** vertices {len(self.mesh.geometry.vertices)} selcount {numVerts}")
        for vi in range(numVerts):
            packetPointer = selSrv.ByIndex(vertSelTypeCode, vi)
            if not packetPointer:
                continue

            vertexID = int(vTransPacket.Vertex(packetPointer))
            item_    = vTransPacket.Item(packetPointer)

            if item_ == self.mesh.geometry._item:
                vertex_ids[v] = vertexID
        return tuple([modo.MeshVertex.fromId(vertexID, self.mesh.geometry) for vertexID in vertex_ids])

    # Main copy function
    def copy(self, external_clipboard='CLIPBOARD'):
        lx.out(f'Copying to external clipboard: {external_clipboard}')
        print("\n".join(dir(modo)))

        scene = modo.Scene()
        mesh = scene.selectedByType("mesh")
        if not mesh:
            return False

        layer_svc = lx.service.Layer()
        layer_scan = lx.object.LayerScan(layer_svc.ScanAllocate(lx.symbol.f_LAYERSCAN_PRIMARY | lx.symbol.f_LAYERSCAN_MARKALL))
        if layer_scan.test() == False:
            return False
        
        sel_svc = lx.service.Selection()
        types = lx.object.storage('i', 4)
        types.set((lx.symbol.iSEL_VERTEX, lx.symbol.iSEL_EDGE, lx.symbol.iSEL_POLYGON, 0))
        self.selType = sel_svc.CurrentType(types)
    
        mesh_svc = lx.service.Mesh()
        self.mark_select = mesh_svc.ModeCompose (lx.symbol.sMARK_SELECT, None)

        self.mesh = mesh[0]
        scene.select(self.mesh)

        # store all selected mesh elements
        self.setup_mesh_elements()

        pos = self.mesh.position.get()
        scl = self.mesh.scale.get()
        lx.out(f'Mesh position: {pos} scale: {scl}')

        locator = lx.object.Locator(self.mesh)
        chan_read = lx.object.ChannelRead(scene.Channels(None, 0.0))
        transforms = locator.LocalTransform(chan_read)

        quat = modo.Quaternion()
        quat.fromMatrix3(transforms[0])

        # Query Existing Materials
        materials_out = []
        for material in scene.items("advancedMaterial"):
            mask = material.parent
            if not mask or mask.type != 'mask':
                name = 'Default'
            else:
                name = mask.channel('ptag').get()
            diffCol = material.channel('diffCol').get()
            mat_data = {
                'name': name,
                'base_color': diffCol,
                'textures': []
            }
            lx.out(f'Found material: {name} color: {diffCol}')
            materials_out.append(mat_data)

        # vertices
        positions = []
        for v in self.mesh.geometry.vertices:
            accessor = v.accessor
            print(f"vertex {v.index} select {accessor.TestMarks(self.mark_select)}")
            positions.append([v.position[0], v.position[1], v.position[2]])

        # crease edges
        crease_edges = [0.0] * len(self.mesh.geometry.edges)
        for vmap in self.mesh.geometry.vmaps:
            if vmap.map_type == lx.symbol.i_VMAP_SUBDIV:
                lx.out(f'Found vmap: {vmap.name} type: {vmap.map_type} len: {len(vmap)}')
                storageBuffer = lx.object.storage('f', 1)
                for i, e in enumerate(self.mesh.geometry.edges):
                    if e.MapEvaluate(vmap.id, storageBuffer) == True:
                        w = storageBuffer.get()
                        crease_edges[i] = w[0]
                break

        # edges
        edges = []
        for i, e in enumerate(self.mesh.geometry.edges):
            edges.append({
                'vertices': [v.index for v in e.vertices],
                'attributes': {
                    'crease_edge': crease_edges[i]
                }
            })

        # faces
        polygons = []
        for p in self.mesh.geometry.polygons:
            p_attrs = {
                'material_index': self.get_material_index(p.materialTag, materials_out)
            }
            polygons.append({
                'vertices': [v.index for v in p.vertices],
                'attributes': p_attrs
            })

        cobj = {
            'name': self.mesh.name,
            'type': 'MESH',
            'object_transform': {
                'translation': [pos[0], pos[1], pos[2]],
                'rotation_quat': [quat[0], quat[1], quat[2], quat[3]],
                'scale': [scl[0], scl[1], scl[2]]
            },
            'mesh': {
                'positions': positions,
                'edges': edges,
                'polygons': polygons,
                'materials': materials_out
            }
        }

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
        if external_clipboard == 'TEMPFILE':
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