#!/usr/bin/env python3
from setuptools import setup, find_packages

with open("README.md", encoding='utf8') as readme:
    long_description = readme.read()

setup(
    name="bingo-card-generator",
    version="1.5.0",
    author="Corey Forman",
    url="https://github.com/digitalsleuth/bingo-card-generator",
    description=("Interactive Bingo Card Generator"),
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
    install_requires=[
        "pdfkit",
        "openpyxl"
    ],
    scripts=['bingo-card-generator.py'],
    package_data={'': ['README.md, LICENSE']}
)
