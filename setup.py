from setuptools import setup, find_packages

with open("README.md", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name = 'xbrx_12345_excel_tool',
    version = "0.1.0",
    author="lbcoder",
    author_email="lbcoder@hotmail.com",
    description="A excel helper tool for xbrx",
    long_description = long_description,
    long_description_content_type = "text/markdown",
    url = "https://github.com/weilancys/xbrx_12345_excel_tool",
    packages = find_packages(),
    include_package_data = True,
    install_requires = [
        "openpyxl",
        "xlrd",
        "jinja2",
    ],
    classifiers = [
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    entry_points = {
        'gui_scripts': [
            'xbrx_12345_excel_tool = xbrx_12345_excel_tool:main',
        ],
    }
)