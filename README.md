# Blazor.Excelerate
<img height=192 width=192 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogoSmallWhite.png>

Blazor.Excelerate is built with DataJuggler.Excelerate, which is built using EPP Plus version 4.5.3.3 (last free version).

This Blazor project uses the following DataJuggler Nuget packages:

DataJuggler.Blazor.FileUpload - used to upload Excel worksheets

DataJuggler.Blazor.Components - ImageButton, ComboBox and ValidationComponent used

DataJuggler.Excelerate to read Excel and code generate C# classes from Excel worksheets

DataJuggler.UltimateHelper - I use this for everything

Blazor.Excelerate comes with a sample spreadsheet with 20,000 Members and 20,000 Addresses *
* This is randomly created sample data, so zip codes, streets, cities, etc. do not match real places.

# Excel Uploads

For this to work, the top row must contain a header row for the field names. For best results, some data rows help
to attempt to determine the data type. Obviously not every excel column name will make a good property name,
so try and name your fields something descriptive if possible. Some testing has been done,
but since this is new code and not many spreadhsheets have been tested, it will take some time to 
perfect this.

# Current Development
I am building an old files deleter, not sure of the name yet, because I need one for PixelDatabase.Net and this site, and my work needs this also for file clean up.

# Known Issues:
The only known issue now is the Generate Class button seems to need to be clicked on the right side of the button.
Making the drop down for the ComboBox further to the right might fix it, or set the ZIndex on the button to a higher value.
Will work on this next time I do an update.

# Version History:

New Update 10.24.2021: I fixed the combo box issue by adding a ZIndex to the ComboBox.




