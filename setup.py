from setuptools import setup

setup(
    name='Brightspace Feedback Entry Scraper',
    version='0.1',
    author='Andrew Guo',
    install_requires=['selenium==4.8.0', 'chromedriver-py==109.0.5414.74', 'xlwings>=0.29.1', 'openpyxl>=3.1.0', 'python-dotenv>=0.21.1'],
)