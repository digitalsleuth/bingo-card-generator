#!/usr/bin/env python3
'''Setup for bingo-card-generator'''

from setuptools import setup, find_packages, os

with open("README.md", encoding='utf8') as readme:
    long_description = readme.read()
    sep = os.sep
setup(
    name="bingo-card-generator",
    version="6.0.1",
    author="Corey Forman",
    url="https://github.com/digitalsleuth/bingo-card-generator",
    description=("Bingo Card Generator"),
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    py_modules=['bingo_card_generator', 'bingo_gui'],
    data_files=[(os.sep, ['bingo.ico', 'README.md', 'LICENSE.md'])],
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
    entry_points={
        'console_scripts': [
            'bingo-gui = bingo_gui:main',
            'bingo-card-generator = bingo_card_generator:main'
        ]
    },
)
