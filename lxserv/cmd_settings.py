#python

'''
Modo command to load and save persist settings

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

import lx
import lxifc
import lxu.command

types = [('tempfile', 'clipboard',),
         ('Temporary File', 'OS Clipboard',)]

class TypePopup(lxifc.UIValueHints):
    def __init__(self, items):
        self._items = items

    def uiv_Flags(self):
        return lx.symbol.fVALHINT_POPUPS

    def uiv_PopCount(self):
        return len(self._items[0])

    def uiv_PopUserName(self,index):
        return self._items[1][index]

    def uiv_PopInternalName(self,index):
        return self._items[0][index]

class PersistData(object):
    def __init__(self):
        self.type = None
        self.type_val = 'tempfile'
        self.replace_mesh = None
        self.replace_mesh_val = 0
        self.replace_material = None
        self.replace_material_val = 0
        self.import_transform = None
        self.import_transform_val = 0

    def get_type(self):
        try:
            return self.type_val.GetString(0)
        except:
            return 'tempfile'

    def set_type(self, type):
        self.type.Append()
        self.type_val.SetString(0, type)

    def get_replace_mesh(self):
        try:
            return self.replace_mesh_val.GetInt(0)
        except:
            return 0

    def set_replace_mesh(self, replace_mesh):
        self.replace_mesh.Append()
        self.replace_mesh_val.SetInt(0, replace_mesh)

    def get_replace_material(self):
        try:
            return self.replace_material_val.GetInt(0)
        except:
            return 0

    def set_replace_material(self, replace_material):
        self.replace_material.Append()
        self.replace_material_val.SetInt(0, replace_material)

    def get_import_transform(self):
        try:
            return self.import_transform_val.GetInt(0)
        except:
            return 0

    def set_import_transform(self, import_transform):
        self.import_transform.Append()
        self.import_transform_val.SetInt(0, import_transform)


persist_data = None

#
# <atom type="ClipboardSettings">
#    <atom type="type">tempfile</atom>
# </atom>
#
class VisClipboardSettings(lxifc.Visitor):
    def vis_Evaluate(self):
        global persist_data
        persist_svc = lx.service.Persistence()

        persist_svc.Start("type", lx.symbol.i_PERSIST_ATOM)
        persist_svc.AddValue(lx.symbol.sTYPE_STRING)
        persist_data.type = persist_svc.End()
        persist_data.type_val = lx.object.Attributes(persist_data.type)

        persist_svc.Start ("replace_mesh", lx.symbol.i_PERSIST_ATOM)
        persist_svc.AddValue (lx.symbol.sTYPE_BOOLEAN)
        persist_data.replace_mesh = persist_svc.End ()
        persist_data.replace_mesh_val = lx.object.Attributes (persist_data.replace_mesh)

        persist_svc.Start ("replace_material", lx.symbol.i_PERSIST_ATOM)
        persist_svc.AddValue (lx.symbol.sTYPE_BOOLEAN)
        persist_data.replace_material = persist_svc.End ()
        persist_data.replace_material_val = lx.object.Attributes (persist_data.replace_material)

        persist_svc.Start ("import_transform", lx.symbol.i_PERSIST_ATOM)
        persist_svc.AddValue (lx.symbol.sTYPE_BOOLEAN)
        persist_data.import_transform = persist_svc.End ()
        persist_data.import_transform_val = lx.object.Attributes (persist_data.import_transform)

        return lx.symbol.e_OK

def persist_setup():
    global persist_data
    if persist_data:
        return
    persist_data = PersistData()
    persist_svc = lx.service.Persistence()
    persist_vis = VisClipboardSettings()
    persist_svc.Configure('ClipboardSettings', persist_vis)

class CmdClipboardSettings(lxu.command.BasicCommand):
    global persist_data
    def __init__(self):
        lxu.command.BasicCommand.__init__(self)
        persist_setup()
        self.dyna_Add('type', lx.symbol.sTYPE_STRING)
        self.basic_SetFlags(0, lx.symbol.fCMDARG_QUERY | lx.symbol.fCMDARG_OPTIONAL)
        self.dyna_Add('replace_mesh', lx.symbol.sTYPE_BOOLEAN)
        self.basic_SetFlags(1, lx.symbol.fCMDARG_QUERY | lx.symbol.fCMDARG_OPTIONAL)
        self.dyna_Add('replace_material', lx.symbol.sTYPE_BOOLEAN)
        self.basic_SetFlags(2, lx.symbol.fCMDARG_QUERY | lx.symbol.fCMDARG_OPTIONAL)
        self.dyna_Add('import_transform', lx.symbol.sTYPE_BOOLEAN)
        self.basic_SetFlags(3, lx.symbol.fCMDARG_QUERY | lx.symbol.fCMDARG_OPTIONAL)

    def arg_UIHints(self, index, hints):
        if index == 0:
            hints.Label("Type")
        elif index == 1:
            hints.Label("Replace Mesh")
        elif index == 2:
            hints.Label("Replace Material")
        elif index == 3:
            hints.Label("Import Transform")

    def arg_UIValueHints(self, index):
        if index == 0:
            return TypePopup(types)

    def basic_Execute(self, msg, flags):
        if self.dyna_IsSet(0):
            persist_data.set_type(self.dyna_String(0))
        if self.dyna_IsSet(1):
            persist_data.set_replace_mesh(self.dyna_Int(1))
        if self.dyna_IsSet(2):
            persist_data.set_replace_material(self.dyna_Int(2))
        if self.dyna_IsSet(3):
            persist_data.set_import_transform(self.dyna_Int(3))

    def cmd_Query(self,index,vaQuery):
        va = lx.object.ValueArray()
        va.set(vaQuery)
        if index == 0:
            va.AddString(persist_data.get_type())
        elif index == 1:
            va.AddInt(persist_data.get_replace_mesh())
        elif index == 2:
            va.AddInt(persist_data.get_replace_material())
        elif index == 3:
            va.AddInt(persist_data.get_import_transform())
        return lx.result.OK

# bless the command to register it as a first class server (plugin)
lx.bless(CmdClipboardSettings, "clipboard.settings")
