# PageSpeed Insights demo tool
*Base on pagespeed insight of google, get information and write to excel to tracking performance.*


## Getting Started
`git clone https://github.com/namdo1983/page_speed_insights_tool.git`

`cd page_speed_insights_tool/`


## Install some library form requirements like: openpyxl, requests, selenium and built-in packages from Python 3.7.9.

If you already have Python with pip installed, you can simply run:

`pip -r requirements.txt`


## Usage

The basic usage is giving a path to a test (or task) file or directory as an argument with possible command line options before the path:

`python main.py`


## Documentation

- ./data: stored urls to run test from main file.
- ./report: stored report excel format after program run complete.
- ./utils: store some class objects that called from main file
