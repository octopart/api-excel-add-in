# Octopart Excel Add-In

The Octopart Excel Add-In enables you to access pricing and availability data from right within Microsoft Excel. You can pull part information for your BOM all at once without leaving your spreadsheet.

## Installation
1. In Excel, choose ‘File' > 'Options’ > ‘Add-Ins’, then press ‘Go’ to manage the ‘Excel Add-Ins’.
![](docs/add-ins.png?raw=true)

2. Browse for the OctopartAddIn, make sure it's selected, and press ‘OK’.
![](docs/install.png?raw=true)

3. To use the worksheet functions, simply type “=OCTOPART….” and the list of functions will appear. Refer to Using Functions for documentation on how to use the functions.
![](docs/example.png?raw=true)


## Ribbon
A new ribbon will be added to your toolbar to make use of the new functionality. 
![](docs/ribbon.png?raw=true)


## Excel Functions
The following functions are available through the Add-In. You can also access the guide to each argument by clicking the Function Wizard after you’ve selected any function. Keep in mind that across most functions, the mpn_or_sku field is required; all other fields are optional.
![](docs/using.png?raw=true)

The first function that you’ll need to use to activate the Add-In is:

`=OCTOPART_SET_APIKEY("_apikey_")`

‘_apikey_’ refers to your unique api key as provided by [Octopart](https://octopart.com/my/api). When you've entered it (e.g., =OCTOPART_SET_APIKEY("00000000-0000-0000-0000-000000000000@example.com") the result will read: `Octopart Add-In is ready`. Your api key is used to track your usage of the Add-In.

From here on, the world is your oyster:

```
=OCTOPART_DETAIL_URL(...)
=OCTOPART_DATASHEET_URL(...)
=OCTOPART_AVERAGE_PRICE(...)
=OCTOPART_DISTRIBUTOR_PRICE(...)
=OCTOPART_DISTRIBUTOR_STOCK(...)
=OCTOPART_DISTRIBUTOR_URL(...)
=OCTOPART_DISTRIBUTOR_MOQ(...)
=OCTOPART_DISTRIBUTOR_PACKAGING(...)
=OCTOPART_DISTRIBUTOR_LEAD_TIME(...)
=OCTOPART_DISTRIBUTOR_ORDER_MUTIPLE(...)
=OCTOPART_DISTRIBUTOR_SKU(...)
=OCTOPART_GET_INFO(...)
=OCTOPART_SET_OPTIONS(...)
```

For results that come in a URL format (e.g., for `=octopart_detail_url` or `=octopart_distributor_url`, click the "Format Hyperlinks" button to activate the links:


# Building the Excel Add-In (Windows):

### Required software
  Download and install [Visual Studio](https://www.visualstudio.com/downloads/)

### Generate XLL Add-In:
  Visual Studio -> Open Project/Solution -> ./OctopartExcelAddIn.sln
    Build all Projects in Release mode to generate dependencies for Excel-DNA script.

### Debugging
  Debugging is easy! Simply setup Visual Studio to 'Start External Program' and point it to your installation of Excel. Pass in the XLL as the sole command line argument.

### Generating Help Files
  To build a compiled help file (.chm) the [Microsoft HTML Help Workshop](http://msdn.microsoft.com/en-us/library/windows/desktop/ms669985(v=vs.85).aspx) (HHW) must be installed.
  
  ExcelDnaDoc expects HHW to be installed at C:\Program Files (x86)\HTML Help Workshop\. If it is installed at another location change packages/ExcelDnaDoc/tools/ExcelDnaDoc.exe.config to reference the proper directory before compiling your project.

  Run `.\OctopartXll\BuildHelp.bat` to compile the help file from the source.