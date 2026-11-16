# -*- coding: utf-8 -*-
import os
from setuptools import setup, find_packages

long_description = ""
if os.path.exists("README.md"):
    with open("README.md", "r", encoding="utf-8") as f:
        long_description = f.read()

setup(
    name="quatest3-core",
    version="1.0.0",
    author="QUATEST 3 - Room N2",
    description="QUATEST 3 Environmental & Cleanroom Monitoring Framework",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(include=["Core", "Core.*"]),
    include_package_data=True,
    package_data={
        "Core": [
            "Templates/**/*",
            "Templates/*",
            "Rule/**/*",
            "Rule/*",
            "*.docx",
            "*.xlsx",
            "*.png",
            "*.jpg",
            "*.jpeg",
            "*.json",
            "*.md",
            "*.txt",
            "*.ps1",
            "*.exe",
        ],
        "Core.Tools": [
            "*.ps1",
            "*.exe",
        ]
    },
    install_requires=[
        "python-docx>=1.0.0",
        "openpyxl>=3.1.0",
        "PyMuPDF>=1.24.0",
        "lxml>=5.0.0",
        "matplotlib>=3.8.0",
        "pillow>=10.0.0",
        "pywin32>=306",
        "xlrd>=2.0.1",
        "PyPDF2>=3.0.0",
        "markitdown>=0.0.1",
    ],
    entry_points={
        "console_scripts": [
            "quatest3=Core.run:main",
            "lamkq=Core.run:main",
        ],
    },
    python_requires=">=3.9",
)
