from setuptools import setup, find_packages
from os import path

this_directory = path.abspath(path.dirname(__file__))
with open('README.md') as f:
    long_description = f.read()

setup(
    name="SapGuiFramework",
    version="0.0.10",
    author="Jason Duncan",
    author_email="jason.matthew.duncan@gmail.com",
    description="A Framework Library for controlling the SAP GUI desktop client",
    long_description=long_description,
    long_description_content_type='text/markdown',
    url="https://github.com/jduncan8142/SapGuiFramework",
    project_urls={
        "Bug Tracker": "https://github.com/jduncan8142/SapGuiFramework/issues",
        "Documentation": "https://github.com/jduncan8142/SapGuiFramework/wiki"
    },
    packages=find_packages(where="SapGuiFramework"),
    classifiers=(
        "Programming Language :: Python :: 3.10",
        "Operating System :: Microsoft :: Windows",
        "License :: OSI Approved :: MIT License",
    ),
    package_dir={"SapGui": "SapGui", "SapLogger": "SapGui", "SapLogonPad": "SapLogonPad", "Utilities": "Utilities"},
    python_requires=">=3.10",
    install_requires=["pywin32>=302", "mss>=6.1.0", "lxml>=4.6.4", "PySimpleGUI>=4.55.1", "xmltodict>=0.12.0", "thefuzz>=0.19.0", "python-Levenshtein>=0.12.2"]
)