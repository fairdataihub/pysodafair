from setuptools import setup, find_packages

setup(
    name="pysoda_fairhub_tools",
    version="0.1.33",
    author="Your Name",
    author_email="cmarroquin@calmi2.org",
    description="Temp pysoda package for Fairhub tools",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/fairdataihub/pysoda",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "openpyxl",
        # Add other dependencies here
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.10.4',
    package_data={
        "pysoda.core.metadata_templates": [".xlsx"],
        "pysoda.schema": [".json"]
    }
)