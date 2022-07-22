from setuptools import setup

setup(
    name="xlsb2xlsx",
    version=0.1,
    description="Convert .xlsb files to .xlsx",
    author="Brian Lewis",
    packages=["xlsb2xlsx"],
    python_requires=">=3.9",
    install_requires=["aspose-cells"],
)
