from setuptools import setup, find_packages

setup(
    name='exceltools',
    version='0.1.0',
    packages=find_packages(),
    install_requires=[
        'openpyxl',
        'tabulate'
    ],
    author='Sanjay Ramkumar',
    author_email='ss.ramsanjay@gmail.com',
    description='A Python package to simplify working with Excel files using openpyxl',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/Sanjay-Ramkumar-0/exceltools',
    license="MIT",
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Programming Language :: Python :: 3',
        'Operating System :: OS Independent',
        'License :: OSI Approved :: MIT License'
    ],
    python_requires='>=3.6',
)

