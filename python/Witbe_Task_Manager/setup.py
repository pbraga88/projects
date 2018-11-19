from distutils.core import setup
import py2exe
setup(
    console = [
        {
            "script": "WitbeTaskManager.py",
            "icon_resources": [(1, "favicon.ico")]
        }
    ],
    )
