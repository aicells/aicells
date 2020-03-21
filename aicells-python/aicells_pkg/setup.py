import setuptools
import os

with open("README.md", "r") as fh:
    longDescription = fh.read()

with open("aicells/LICENSE", "r") as fh:
    license = fh.read()

with open('requirements.txt') as f:
    requiredPackages = f.read().strip().splitlines()

def RecursiveListFiles(directory):
    fileList = []

    for (path, directories, fileNames) in os.walk(directory):
        for fileName in fileNames:
            fileNameWithPath = os.path.join('..', path, fileName)
            if not ("__pycache__" in fileNameWithPath):
                fileList.append(fileNameWithPath)

    return fileList

fileList = RecursiveListFiles('aicells')

setuptools.setup(
    name="aicells",
    version="0.0.1",
    author="Gergely Szerovay, László Siller",
    author_email="gergely@szerovay.hu",
    description="AIcells package",
    long_description=longDescription,
    long_description_content_type="text/markdown",
    url="https://github.com/aicalles/aicells",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires='>=3.7',
    install_requires=requiredPackages,
    package_data={'': fileList},
    include_package_data=True,
    license=license
)
