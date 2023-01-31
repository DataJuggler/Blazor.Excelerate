<img src =https://excelerate.datajuggler.com/Images/LogoTextSparkled.png><br>
Code Generate C# Classes From Excel Header Rows

<img src =https://excelerate.datajuggler.com/Images/ExcelerateLogoSmallWhite.png width="128" height="128">

Live Demo:
<a href=https://excelerate.datajuggler.com target="_blank">https://excelerate.datajuggler.com</a>

Instructions (copied from Index.razor.cs)
1. Prepare your spreadsheet (save as .xlsx extension) and ensure you have a header row, and the column names make good property names. 
   Download MemberData.xlsx to see an example (Downloads\MemberData.xlsx comes with this project and contains 20,000 random names and addresses).
   

   
2. Click the Upload button to upload your spreadsheet and select the sheet to code generate a class for.

   <img src =https://github.com/DataJuggler/SharedRepo/blob/master/Shared/Images/ExcelerateStep1.png><br>
   
3. Type in a namespace for your project and select the sheet in the Sheets ComboBox then click the 'Generate Class' button.

    <img src =https://github.com/DataJuggler/SharedRepo/blob/master/Shared/Images/ExcelerateStep2.png><br>
    
4. Download the zip file and extract the contents to get your C# file.

   <img src =https://github.com/DataJuggler/SharedRepo/blob/master/Shared/Images/ExcelerateStep3.png><br>

Tips / Troubleshooting
It helps to have some rows of data to attempt to determine the data type. Getting the data type is kind of a hack, as I could not find away with EPP Plus to get the column data types.

Working Demo: Excelerate WinForms Demo
https://github.com/DataJuggler/Excelerate.WinForms.Demo

<img src=https://github.com/DataJuggler/SharedRepo/blob/master/Shared/Images/Excelerate%20Win%20Forms%20Demo.png width=512 height=512>






