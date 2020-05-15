# PortExtract
---
Cisco Port Extractor  
Used to extract L2 switchport information and export into an easily readable spreadsheet.
  
### Requirements
* [xlsxwriter](https://github.com/jmcnamara/XlsxWriter)
  
### Getting Started
---
1. Clone the repo  
`git clone https://byte-of-reyn/portextract`  
2. Install library requirements  
`python -m pip install requirements.txt`
  
### Usage
---
Single input file  
`python confextract.py -i input-a -o output.xlsx`

Multiple input files  
`python confextract.py -i input-a, input-b -o output.xlsx`
  
### Contact
---
byte.of.reyn@gmail.com

### License
---
Distributed under the MIT License. See LICENSE for more information.
