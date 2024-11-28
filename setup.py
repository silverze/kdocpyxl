import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="kdocpyxl",
    version="0.0.2",
    author="silverze",
    author_email="silverze@foxmail.com",
    description="Kingsoft cloud document excel Python read and write library",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/silverze/kdocpyxl",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)