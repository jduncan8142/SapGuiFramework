from setuptools import setup, find_packages
from os import path
from SapGuiFramework import __version__

this_directory = path.abspath(path.dirname(__file__))
with open('README.md') as f:
    long_description = f.read()

setup(
    name="SapGuiFramework",
    version=__version__,
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
    classifiers=[
        "Programming Language :: Python :: 3.11",
        "Operating System :: Microsoft :: Windows",
        "License :: OSI Approved :: MIT License",
    ],
    package_dir={"Core": "SapGuiFramework\Core", "Logging": "SapGuiFramework\Logging", "Flow": "SapGuiFramework\Flow"},
    python_requires=">=3.11",
    install_requires=["pywin32>=305", "PyYAML>=6.0", "selenium>=4.10.0", "python-dotenv>=1.0.0", "chromedriver-binary-auto>=0.2.6"],
    extras_require={"dev": ["pytest>=7.0", "twine>=4.0.2"]}
)