# vbs-excel-utilities
Excel Unpacker and other excel utilities written in Vbscript


## Installation

```sh
npm i vbs-excel-utilities
```

## Usage

### To Unpack an excel workbook
```sh
npx xlunpack /workbook:Workbook-Name.xlsm /destination:Workbook-Name
```


### To Pack VBA Modules into an excel workbook
```sh
npx xlpack /workbook:Workbook-Name.xlsm /source:Workbook-Name
```