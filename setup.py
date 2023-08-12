import os
import re

from setuptools import find_packages, setup


def read_version():
    here = os.path.abspath(os.path.dirname(__file__))
    version_file = os.path.join(here, 'mypackage', '_version.py')
    with open(version_file, 'r') as f:
        version_data = f.read()
    version_match = re.search(r"^__version__ = ['\"]([^'\"]*)['\"]", version_data, re.M)
    if version_match:
        return version_match.group(1)
    else:
        raise RuntimeError("Unable to find version string.")


setup(
    name="mypackage",
    version=read_version(),
    packages=find_packages(),
)
