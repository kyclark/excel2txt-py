import setuptools
import excel2txt

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='excel2txt',
    version=excel2txt.VERSION,
    author='Ken Youens-Clark',
    author_email='kyclark@gmail.com',
    description='Convert Excel files to delimited text',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/kyclark/excel2txt-py',
    packages='.',
    entry_points={
        'console_scripts': [
            'excel2txt=excel2txt:main',
        ],
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
)
