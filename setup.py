from setuptools import setup, find_packages
from os import path

this_directory = path.abspath(path.dirname(__file__))
with open('README.md') as f:
    long_description = f.read()

setup(
    name="SapGuiFramework",
    version="0.0.7",
    author="Jason Duncan",
    author_email="jason.matthew.duncan@gmail.com",
    description="A Robocorp Robot Framework Library for controlling the SAP GUI desktop client",
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
    package_dir={"": "SapGui"},
    python_requires=">=3.10",
    install_requires=["pywin32>=300", "robotframework>=4.1.2"]
)