# Blazor.Excelerate
<img height=256 width=256 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogo.png>

Blazor.Excelerate is built with DataJuggler.Excelerate, which is built using EPP Plus version 4.5.3.3 (last free version).

This Blazor project uses the following DataJuggler Nuget packages:

DataJuggler.Blazor.FileUpload - used to upload Excel worksheets
DataJuggler.Blazor.Components - ImageButton and ComboBox components used
DataJuggler.Excelerate to read Excel and code generate C# classes from Excel worksheets.
DataJuggler.UltimateHelper - I use this for everything



This project comes with a sample spreadsheet with 20,000 Members and 20,000 Addresses *
* This is randomly created sample data, so zip codes may not match streets, cities, etc.

# Excel Uploads

For this to work, the top row must contain a header row for the field names. For best results, some data rows help
to attempt to determine the data type. Obviously not every excel column name will make a good property name,
so try and name your fields something descriptive if possible. Some testing has been done,
but since this is new code and not many spreadhsheets have been tested, it will take some time to 
perfect this.


