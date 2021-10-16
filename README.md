# Blazor.Excelerate
<img height=64 width=64 src=https://github.com/DataJuggler/Blazor.Excelerate/blob/main/wwwroot/Images/ExcelerateLogo.png>

This Blazor project uses DataJuggler.Blazor.Components and DataJuggler.Excelerate to code generate
C# classes from Excel worksheets.

For this to work, the top row must contain the field names. For best results, some data rows help
to attempt to determine the data type. Obviously not every excel column name makes a good property name,
so try and name your fields something descriptive if possible. Some testing has been done,
but since this is new code and not many spreadhsheets have been looked at, it will take some time to 
perfect this.
