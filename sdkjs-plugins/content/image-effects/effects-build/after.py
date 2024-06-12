import sys
sys.path.append("../../build_tools/scripts")
import base

base.configure_common_apps()
base.replaceInFile("../effects-core/deploy/engine/effects.js", "__ATPOSTRUN__=[];", "__ATPOSTRUN__=[onLoadModule];")
base.replaceInFile("../effects-core/deploy/engine/effects_ie.js", "__ATPOSTRUN__=[];", "__ATPOSTRUN__=[onLoadModule];")

base.copy_file("../effects-core/js/library.js", "../effects-core/deploy/effects.js")