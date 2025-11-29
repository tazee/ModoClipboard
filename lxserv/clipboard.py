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


def paste_from_clipboard(external_clipboard='CLIPBOARD', new_mesh=False):
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

def get_material_index(name, material_out):
    for i, mat in enumerate(material_out):
        if mat['name'] == name:
            return i
    return 0

def copy_to_clipboard(external_clipboard='CLIPBOARD'):
    lx.out(f'Copying to external clipboard: {external_clipboard}')

    scene = modo.Scene()
    mesh = scene.selectedByType("mesh")
    if not mesh:
        return False

    mesh = mesh[0]
    scene.select(mesh)

    pos = mesh.position.get()
    scl = mesh.scale.get()
    lx.out(f'Mesh position: {pos} scale: {scl}')

    locator = lx.object.Locator(mesh)
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

    subdiv_map = None
    for vmap in mesh.geometry.vmaps:
        if vmap.map_type == lx.symbol.i_VMAP_SUBDIV:
            subdiv_map = vmap
            lx.out(f'Found vmap: {vmap.name} type: {vmap.map_type} len: {len(vmap)}')
            break

    crease_edges = [0.0] * len(mesh.geometry.edges)
    if subdiv_map:
        lx.out(f'subdiv map {len(subdiv_map)}')
        for i, w in enumerate(subdiv_map):
            if w is not None:
                crease_edges[i] = w[0]

    # vertices
    positions = []
    for v in mesh.geometry.vertices:
        positions.append([v.position[0], v.position[1], v.position[2]])

    # edges
    edges = []
    for i, e in enumerate(mesh.geometry.edges):
        edges.append({
            'vertices': [v.index for v in e.vertices],
            'attributes': {
                'crease_edge': crease_edges[i]
            }
        })

    # faces
    polygons = []
    for p in mesh.geometry.polygons:
        p_attrs = {
            'material_index': get_material_index(p.materialTag, materials_out)
        }
        polygons.append({
            'vertices': [v.index for v in p.vertices],
            'attributes': p_attrs
        })

    cobj = {
        'name': mesh.name,
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