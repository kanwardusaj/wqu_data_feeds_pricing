# WQU DATA FEEDS Pricing
We will explore the use of C# and Excel to keep track of the pricing of various properties. We shall look to implement various useful statistical calculations which will allow you to gain further insight into the overall trends of the property market.

In this submission, below tasks are presented :

i)Set up the worksheet when the application is launched for the first time - The main method already calls a method “SetUp”; therefore, you simply have to implement this method, which should create a new Excel workbook titled “property_pricing.xlsx”

ii)Implement the adding of property information to the sheet - The property information headers are as follows:
a. Size (in square feet)
b. Suburb
c. City
d. Market value

iii)Implement statistical methods -. In the skeleton code you find the following four statistical methods already declared:
a. Mean market value
b. Variance in market value
c. Minimum market value
d. Maximum market value

Note - Please follow below steps to configure based on the version of Microsoft Office -

Excel 2016 and above - 
Navigate to the following directory on your windows machine: 
c:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0_71e9bce111e9429c
Inside this directory you should see the following .dll file:
Microsoft.Office.Interop.Excel.dll
Copy the 15.0.0.0_71e9bce111e9429c directory, together with the .dll file it contains into the following directory:
c:\Windows\assembly\GAC_MSIL\office
If the "office" directory doesn't exist, create it.
Don't forget to add in VisualStudio in your project/solution the following COM references: 
Microsoft Office 16.0 Object Library
Microsoft Excel 16.0 Object Library

Excel 2013 -

1.	Start Microsoft Visual Studio 2017 and or (above) or Microsoft Visual Studio .NET.
2.	On the File menu, click New, and then click Project. Select Windows Application from the Visual C# Project types. Form1 is created by default.
3.	Add a reference to Microsoft Excel 11.0 Object Library in Visual Studio 2017 or Microsoft Excel Object Library in Visual Studio .NET. To do this, follow these steps:(optional changes if using previous versions of Office )
i.	On the Project menu, click Add Reference.
ii.	On the COM tab, locate Microsoft Excel Object Library, and then click Select.
In Visual Studio 2017, locate Microsoft Excel 11.0 Object Library on the COM tab.
Note Microsoft Office 2003 includes Primary Interop Assemblies (PIAs). Microsoft Office XP does not include PIAs, but they may be downloaded.
