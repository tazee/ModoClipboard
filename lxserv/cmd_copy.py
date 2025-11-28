#python

'''

    Moddo command to copy selected mesh elements to external 
    clipboard using Python script.

'''

import lx
import lxu.command
import clipboard

class ClipboardCopy(lxu.command.BasicCommand):

    def __init__(self):
        lx.out("ClipboardCopy: initializing")
        lxu.command.BasicCommand.__init__(self)

    def cmd_Flags(self):
        return lx.symbol.fCMD_MODEL | lx.symbol.fCMD_UNDO

    def basic_Enable(self, msg):
        return True

    def cmd_Interact(self):
        pass

    def basic_Execute(self, msg, flags):
        lx.out("ClipboardCopy: Executing Copy to External")
        clipboard.copy_to_clipboard(external_clipboard='CLIPBOARD')

    def cmd_Query(self, index, vaQuery):
        lx.notimpl()


lx.bless(ClipboardCopy, "clipboard.copy")
