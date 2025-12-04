#python

'''

    Modo command to copy selected mesh elements to external 
    clipboard using Python script.

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
