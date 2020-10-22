import setuptools

install_requires = [
    'compoundfiles',
    'compressed_rtf',
]

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='convert-outlook-msg-file',
    version='0.1.0',
    description='Parse Microsoft Outlook MSG files',
    author='Joshua Tauberer',
    author_email='jt@occams.info',
    url='https://github.com/JoshData/convert-outlook-msg-file',
    py_modules=['outlookmsgfile'],
    install_requires=install_requires,
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires='>=3.6',
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
