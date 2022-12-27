# Author: Tomoki Nishikawa <nishitomo206@gmail.com>
# Copyright (c) 2022 nishitomo206
# License: MIT License

from setuptools import setup, find_packages

# import src.pyopentd
from src import pyopentd

DESCRIPTION = ""
NAME = "pyopentd"
AUTHOR = "Tomoki Nishikawa"
AUTHOR_EMAIL = "nishitomo206@gmail.com"
URL = "https://github.com/nishitomo206/pyopentd"
LICENSE = "MIT License"
DOWNLOAD_URL = "https://github.com/nishitomo206/pyopentd"
VERSION = pyopentd.__version__
# PYTHON_REQUIRES = ">=3.6"

INSTALL_REQUIRES = []

EXTRAS_REQUIRE = {}

# PACKAGES = [
#     'pyopentd'
# ]

CLASSIFIERS = []

# with open('README.rst', 'r') as fp:
#     readme = fp.read()
# with open('CONTACT.txt', 'r') as fp:
#     contacts = fp.read()
# long_description = readme + '\n\n' + contacts

setup(
    name=NAME,
    author=AUTHOR,
    author_email=AUTHOR_EMAIL,
    maintainer=AUTHOR,
    maintainer_email=AUTHOR_EMAIL,
    description=DESCRIPTION,
    #   long_description=long_description,
    license=LICENSE,
    url=URL,
    version=VERSION,
    download_url=DOWNLOAD_URL,
    #   python_requires=PYTHON_REQUIRES,
    install_requires=INSTALL_REQUIRES,
    extras_require=EXTRAS_REQUIRE,
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=CLASSIFIERS,
)
