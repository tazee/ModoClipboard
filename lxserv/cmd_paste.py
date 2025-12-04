#python

'''

    Modo command to paste the mesh elements from external 
    clipboard using Python script.

'''

import lx
import lxu.command
import clipboard

class ClipboardPaste(lxu.command.BasicCommand):

    def __init__(self):
        lxu.command.BasicCommand.__init__(self)
        self.dyna_Add("new_mesh", lx.symbol.sTYPE_BOOLEAN)
        self.basic_SetFlags(0, lx.symbol.fCMDARG_OPTIONAL)

    def cmd_Flags(self):
        return lx.symbol.fCMD_MODEL | lx.symbol.fCMD_UNDO

    def basic_Enable(self, msg):
        return True

    def cmd_Interact(self):
        pass

    def basic_Execute(self, msg, flags):
        type = lx.eval("clipboard.settings type:?")
        new_mesh = self.dyna_Int(0)
        lx.out(f"ClipboardPaste: Executing Paste to External {new_mesh}")
        clipboard.ClipboardData().paste(external_clipboard=type, new_mesh=new_mesh)

    def cmd_Query(self, index, vaQuery):
        lx.notimpl()


lx.bless(ClipboardPaste, "clipboard.paste")
