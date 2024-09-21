# setup.py
from setuptools import setup, find_packages

setup(
    name="tex2docx",
    version="1.0",
    packages=find_packages(where='src'),
    package_dir={"": "src"},
    install_requires=[
        'argparse',
        'regex',
        'tqdm',
    ]
)
