from setuptools import setup, find_packages

with open("requirements.txt") as f:
	install_requires = f.read().strip().split("\n")

# get version from __version__ variable in iiq_check_connect/__init__.py
from iiq_check_connect import __version__ as version

setup(
	name="iiq_check_connect",
	version=version,
	description="Export Customer Data to iiq-check",
	author="itsdave GmbH",
	author_email="dev@itsdave.de",
	packages=find_packages(),
	zip_safe=False,
	include_package_data=True,
	install_requires=install_requires
)
