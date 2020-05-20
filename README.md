# PortExtract

Cisco Port Extractor - Parse L2 switchport information, from running-config file/s, and export into an easily readable spreadsheet.
  
### Requirements

* Python3
* Microsoft Excel2016
* [xlsxwriter](https://github.com/jmcnamara/XlsxWriter)


### Getting Started

1. Clone the repo  

`git clone https://byte-of-reyn/portextract`  

2. Install library requirements  

`python -m pip install -r requirements.txt`

### Usage

Single input file  

`python portextract.py -i input-a -o output.xlsx`

Multiple input files  

`python portextract.py -i input-a, input-b -o output.xlsx`
  
### Roadmap 

TODO: Feature add listing
  
### Contact

byte.of.reyn@gmail.com

### License

Distributed under the MIT License. See LICENSE for more information.
