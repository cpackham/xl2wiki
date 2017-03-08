from setuptools import setup

setup(
    name="xl2wiki",
    description="Convert Excel(tm) spreadsheets to MediaWiki table markup",
    license="GPLv3",
    packages=["xl2wiki"],
    entry_points={
        "console_scripts": [
            "xl2wiki=xl2wiki:main",
        ],
    },
    install_requires=["xlrd"])
