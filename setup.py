# setup.py
from setuptools import setup, find_packages

setup(
    name="tex2docx",
    version="1.1.0",
    packages=find_packages(where='tex2docx'),
    package_dir={"": "tex2docx"},
    install_requires=[
        'argparse',
        'regex',
        'tqdm',
    ]
)
