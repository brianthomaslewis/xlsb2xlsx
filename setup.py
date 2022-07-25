from setuptools import setup
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()

setup(
    name="xlsb2xlsx",
    version="0.1",
    long_description=long_description,
    long_description_content_type="text/markdown",
    description="Convert .xlsb files to .xlsx.",
    author="Brian Lewis",
    packages=["xlsb2xlsx"],
    python_requires=">=3.7",
    install_requires=["aspose-cells", "tqdm"],
)
