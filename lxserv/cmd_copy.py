#python

'''

Modo command to copy selected mesh elements to external 
clipboard using Python script.

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
import lxu.command
import clipboard

class ClipboardCopy(lxu.command.BasicCommand):

    def __init__(self):
        lx.out("ClipboardCopy: initializing")
        lxu.command.BasicCommand.__init__(self)
        self.dyna_Add("cut", lx.symbol.sTYPE_BOOLEAN)
        self.basic_SetFlags(0, lx.symbol.fCMDARG_OPTIONAL)

    def cmd_Flags(self):
        return lx.symbol.fCMD_MODEL | lx.symbol.fCMD_UNDO

    def basic_Enable(self, msg):
        return True

    def cmd_Interact(self):
        pass

    def basic_Execute(self, msg, flags):
        type = lx.eval("clipboard.settings type:?")
        lx.out(f"ClipboardCopy: Executing Copy to External {type}")
        clipboard.ClipboardData().copy(external_clipboard=type)
        cut = self.dyna_Int(0)
        if cut:
            lx.eval("select.delete")

    def cmd_Query(self, index, vaQuery):
        lx.notimpl()


lx.bless(ClipboardCopy, "clipboard.copy")
