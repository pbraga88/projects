from distutils.core import setup
import py2exe
setup(
    console = [
        {
            "script": "SSIDVerifier.py",
			"icon_resources": [(1, "Wifi.ico")]
        }
    ],
    )
