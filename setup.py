#!/usr/bin/env python3
'''Setup for bingo-card-generator'''

from setuptools import setup, find_packages

with open("README.md", encoding='utf8') as readme:
    long_description = readme.read()

setup(
    name="bingo-card-generator",
    version="6.0.0",
    author="Corey Forman",
    url="https://github.com/digitalsleuth/bingo-card-generator",
    description=("Bingo Card Generator"),
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=[
        "PyQt6",
        "openpyxl",
        "pdfkit",
        "Pillow",
        "webcolors"
    ],
    scripts=['bingo_card_generator.py', 'bingo_gui.py'],
    package_data={'': ['README.md, LICENSE']}
)
