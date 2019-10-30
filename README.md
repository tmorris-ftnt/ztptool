# ZTP Tool

ZTP Tool is a small GUI application to assist with setting up Fortinet FortiManager for Zero Touch Provisioning (ZTP) of SD-WAN deployments. ZTP Tool has two main functions:

1) Import devices from an Excel file (Model Device with Meta Data, CLI Templates, SD-WAN templates and Policy Packages applied)
2) Import ADOM from JSON file (with CLI Templates, SD-WAN Templates and Policy packages in place)

Downloads are available on the [releases](https://github.com/tmorris-ftnt/ztptool/releases) page.

<p align="center"><img src="example/screenshot.png" ></p>

## Getting Started with demo_example
Extact the files from the .zip archive somewhere on your computer.
Note: Chrome must be installed the computer.

Create an user on FortiManager version 6.2.1+ with '<rpc-permit read-write'> permissions set with ADOM mode enabled. 

Open ztptool.exe and goto the Import ADOM page.

Fill in the form and select the demo_example.json file.

This will setup your new ADOM with CLI Templates, SDWAN Templates, Policy Packages and Objects. 

Now goto the Import Devices Page. 

Fill in the form and select the demo_example.xlsx file. 

This will populate your FortiManager ADOM with prebuild model devices. 


