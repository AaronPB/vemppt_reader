# VE MPPT Reader


vemppt_reader reads VE.Direct Protocol and registers values from MPPT Devices into xls files

<img src="images/header_image.JPG" >

## MPPT Connection


This software is tested with a [VE.Direct to USB Interface cable](https://www.victronenergy.com.es/accessories/ve-direct-to-usb-interface)


## Installation

vemppt_reader requires [pyserial](https://pypi.org/project/pyserial/) and [xlwt](https://pypi.org/project/xlwt/) to run


```sh

$ pip install pyserial

$ pip install xlwt

```


## VE.Direct Protocol Documentation

  - [VE.Direct Protocol - Version 3.25](https://www.victronenergy.com.es/download-document/2036/ve.direct-protocol-3.25.pdf)

  - [BlueSolar HEX protocol MPPT](https://www.victronenergy.com.es/download-document/4459/bluesolar-hex-protocol-mppt.pdf)


## TO-DO List

  - Asynchronous message interpreter and register
