# Blazor.Excelerate
<img height=192 width=192 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogoSmallWhite.png>

*******
10.23.2021 - Oops

I managed to mess the combo box up today, so bear with me while i fix the ComboBox for the sheets. I started playing with ZIndex'es and I managed to break what worked.
I am working on a fix or rolling back to something that works as soon as possible.

For now as a workaround, you can code generate the first sheet, so arrange your Excel sheet that way and save until I work this out.

*******

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


