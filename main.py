import pptx
import typing
import util.Helper
import importlib

# Reload - Helper util
importlib.reload(util.Helper)

helper = util.Helper.PPTHelper("/mnt/c/Users/kael/Desktop/test.pptx", "target")
helper.to_html()
