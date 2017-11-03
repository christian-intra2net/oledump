#!/usr/bin/env python
"""
Installs oledump using distutils

Run:
    python setup.py install

to install this package.

(setup script partly borrowed from oletools)
"""

#--- IMPORTS ------------------------------------------------------------------

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

#from distutils.command.install import INSTALL_SCHEMES

import os, fnmatch
from glob import glob
from os.path import join, relpath


#--- METADATA -----------------------------------------------------------------

name         = "oledump"
version      = '0.0.29'
desc         = "Analyze OLE files (Compound Binary Files)"
author       = "Didier Stevens"
url          = "https://DidierStevens.com"
download_url = "https://didierstevens.com/files/software/"

#--- PACKAGES -----------------------------------------------------------------

packages=[
    "oledump",
]

package_data={
    'oledump': ['*.yara']
}
# --- SCRIPTS ------------------------------------------------------------------

# Entry points to create convenient scripts automatically

entry_points = {
    'console_scripts': [
        'oledump=oledump.oledump:main',
        'oledump_all=oledump.oledump_all:main'
    ],
}


# === MAIN =====================================================================

def main():
    # TODO: warning about Python 2.6
##    # set default location for "data_files" to
##    # platform specific "site-packages" location
##    for scheme in list(INSTALL_SCHEMES.values()):
##        scheme['data'] = scheme['purelib']

    dist = setup(
        name=name,
        version=version,
        description=desc,
        author=author,
        url=url,
        packages=packages,
        package_data = package_data,
        download_url=download_url,
        entry_points=entry_points,
    )


if __name__ == "__main__":
    main()
