#python

'''

    Moddo command to paste the mesh elements from external 
    clipboard using Python script.

'''

import lx
import lxu.command
import clipboard

class ClipboardPaste(lxu.command.BasicCommand):

    def __init__(self):
        lx.out("ClipboardPaste: initializing")
        lxu.command.BasicCommand.__init__(self)

    def cmd_Flags(self):
        return lx.symbol.fCMD_MODEL | lx.symbol.fCMD_UNDO

    def basic_Enable(self, msg):
        return True

    def cmd_Interact(self):
        pass

    def basic_Execute(self, msg, flags):
        lx.out("ClipboardPaste: Executing Paste to External")
        clipboard.paste_from_clipboard(external_clipboard='CLIPBOARD', new_mesh=False)

    def cmd_Query(self, index, vaQuery):
        lx.notimpl()


lx.bless(ClipboardPaste, "clipboard.paste")
