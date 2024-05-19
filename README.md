<<<<<<< HEAD

5.18.2024: New Video

First Ever Opensource Saturday - Sunday Edition
https://youtu.be/kohGlLIBMR0

7.24.2023: New Video

The Best C# Excel Library In The Galaxy
https://youtu.be/uWXiz52cqlg

<img src =https://excelerate.datajuggler.com/Images/LogoTextSparkled.png><br>
Code Generate C# Classes From Excel Header Rows

<img src =https://excelerate.datajuggler.com/Images/ExcelerateLogoSmallWhite.png width="128" height="128">

Live Demo:
<a href=https://excelerate.datajuggler.com target="_blank">https://excelerate.datajuggler.com</a>

Blazor Excelerate is an open source Blazor demo for Nuget DataJuggler.Excelerate, which is built on top of EPP Plus 4.5.3.3 (last free version).

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

# News

Update 7.24.2023:

I created a WinForms (desktop) app that can be installed via NuGet and the dotnet CLI

To Install Via Nuget and DOT NET CLI, navigate to the folder you wish to create your project in

    cd c:\Projects\ExcelerateWinApp
    dotnet new install DataJuggler.ExcelerateWinApp
    dotnet new DataJuggler.ExcelerateWinApp

or

Clone this project from GitHub https://github.com/DataJuggler/ExcelerateWinApp

# Setup Instructions for ExcelerateWinApp, copied from the above project.

1. Create one or more classes from Excel Header Rows at<br><br>

Blazor Excelerate<br>
https://excelerate.datajuggler.com<br>

Download the file MemberData.xlsx from the above site to see an example.
Use ExcelerateWinApp.Objects for the namespace or rename this project to your liking
 
2. Copy the classes created into the Objects folder of ExcelerateWinApp

3. Load Excel Worksheet(s) - Example is included in the UpdateButton_Click event
	
       // load your object(s)
       string workbookPath = FileSelector.Text;

       // Example WorksheetInfo objects           
       WorksheetInfo info = new WorksheetInfo();
       info.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
       info.Path = workbookPath;	

       // Set your SheetName
       info.SheetName = "Address";

       // Example WorksheetInfo objects           
       WorksheetInfo info2 = new WorksheetInfo();
       info2.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
       info2.Path = workbookPath;

       // Set the SheetName for info2
       info2.SheetName = 'States";

       // Example load Worksheets
       Worksheet addressWorksheet = ExcelDataLoader.LoadWorksheet(workbookPath, info);
       Worksheet statesWorksheet = ExcelDataLoader.LoadWorksheet(workbookPath, info2);

5. Load your list of objects
 
        // Examples loading the Address and States sheet from MemberData.xlsx
        List<Address> addresses = Address.Load(addressWorksheet);
        List<States> states = States.Load(statesWorksheet);

6. Perform updates on your List of objects

   For this example, I inserted a column StateName into the Address sheet in Excel and
   added a few state names manually. You must add a few entries so the data type can be
   attempted to be determined. Then I code generated Address and States classes using
   Blazor Excelerate<br>
   https://excelerate.datajuggler.com

   This method set the Address.StateName for each row by looking up the State Name by StateId
	
       /// <summary>
       /// Lookup the StateName for each Address object by StateId
       /// </summary>
       public void FixStateNames(ref List<Address> addresses, List<States> states)
       {
           // verify both lists exists and have at least one item
           if (ListHelper.HasOneOrMoreItems(addresses, states))
           {
              // Iterate the collection of Address objects
              foreach (Address address in addresses)
              {
                  // get a local copy
                  int stateId = address.StateId;

                  // set the stateName
                  address.StateName = states.Where(x => x.Id == stateId).FirstOrDefault().Name;

                  // Increment the value for Graph
                  Graph.Value++;

                  // update the UI every 100
                  if (Graph.Value % 100 == 0)
                  {
                      Refresh();
                      Application.DoEvents();
                   }
              }
           }
        }
	
7. Save your worksheet back to Excel

       // resetup the graph                    
       Graph.Maximum = addresses.Count;
       Graph.Value = 0;

       // change the text
       StatusLabel.Text = "Saving Addresses please wait...";

       // you must convert the list objects to List<IExcelerateObject> before it can be saved
       List<IExcelerateObject> excelerateObjectList = addresses.Cast<IExcelerateObject>().ToList();

       // Now save the worksheet
       SaveWorksheetResponse response = ExcelHelper.SaveWorksheet(excelerateObjectList, addressWorksheet, info, SaveWorksheetCallback, 500);

8. (Optional) Leave a Star on DataJuggler.Excelerate, Blazor Excelerate or this project on GitHub

    DataJuggler.Excelerate
    https://github.com/DataJuggler/Excelerate

    Blazor Excelerate
    https://github.com/DataJuggler/Blazor.Excelerate
	
    Excelerate Win App
    https://github.com/DataJuggler/ExcelerateWinApp

9. (Optional) Subscribe to my YouTube channel
    https://youtube.com/DataJuggler







=======

7.24.2023: New Video

The Best C# Excel Library In The Galaxy
https://youtu.be/uWXiz52cqlg

<img src =https://excelerate.datajuggler.com/Images/LogoTextSparkled.png><br>
Code Generate C# Classes From Excel Header Rows

<img src =https://excelerate.datajuggler.com/Images/ExcelerateLogoSmallWhite.png width="128" height="128">

Live Demo:
<a href=https://excelerate.datajuggler.com target="_blank">https://excelerate.datajuggler.com</a>

Update 1.14.2024: New Video - Import a Cities Database in Excel to SQL Server: https://youtu.be/6LsFP0puuyA

Blazor Excelerate is an open source Blazor demo for Nuget DataJuggler.Excelerate, which is built on top of EPP Plus 4.5.3.3 (last free version).

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

# News

Update 7.24.2023:

I created a WinForms (desktop) app that can be installed via NuGet and the dotnet CLI

To Install Via Nuget and DOT NET CLI, navigate to the folder you wish to create your project in

    cd c:\Projects\ExcelerateWinApp
    dotnet new install DataJuggler.ExcelerateWinApp
    dotnet new DataJuggler.ExcelerateWinApp

or

Clone this project from GitHub https://github.com/DataJuggler/ExcelerateWinApp

# Setup Instructions for ExcelerateWinApp, copied from the above project.

1. Create one or more classes from Excel Header Rows at<br><br>

Blazor Excelerate<br>
https://excelerate.datajuggler.com<br>

Download the file MemberData.xlsx from the above site to see an example.
Use ExcelerateWinApp.Objects for the namespace or rename this project to your liking
 
2. Copy the classes created into the Objects folder of ExcelerateWinApp

3. Load Excel Worksheet(s) - Example is included in the UpdateButton_Click event
	
       // load your object(s)
       string workbookPath = FileSelector.Text;

       // Example WorksheetInfo objects           
       WorksheetInfo info = new WorksheetInfo();
       info.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
       info.Path = workbookPath;	

       // Set your SheetName
       info.SheetName = "Address";

       // Example WorksheetInfo objects           
       WorksheetInfo info2 = new WorksheetInfo();
       info2.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
       info2.Path = workbookPath;

       // Set the SheetName for info2
       info2.SheetName = 'States";

       // Example load Worksheets
       Worksheet addressWorksheet = ExcelDataLoader.LoadWorksheet(workbookPath, info);
       Worksheet statesWorksheet = ExcelDataLoader.LoadWorksheet(workbookPath, info2);

5. Load your list of objects
 
        // Examples loading the Address and States sheet from MemberData.xlsx
        List<Address> addresses = Address.Load(addressWorksheet);
        List<States> states = States.Load(statesWorksheet);

6. Perform updates on your List of objects

   For this example, I inserted a column StateName into the Address sheet in Excel and
   added a few state names manually. You must add a few entries so the data type can be
   attempted to be determined. Then I code generated Address and States classes using
   Blazor Excelerate<br>
   https://excelerate.datajuggler.com

   This method set the Address.StateName for each row by looking up the State Name by StateId
	
       /// <summary>
       /// Lookup the StateName for each Address object by StateId
       /// </summary>
       public void FixStateNames(ref List<Address> addresses, List<States> states)
       {
           // verify both lists exists and have at least one item
           if (ListHelper.HasOneOrMoreItems(addresses, states))
           {
              // Iterate the collection of Address objects
              foreach (Address address in addresses)
              {
                  // get a local copy
                  int stateId = address.StateId;

                  // set the stateName
                  address.StateName = states.Where(x => x.Id == stateId).FirstOrDefault().Name;

                  // Increment the value for Graph
                  Graph.Value++;

                  // update the UI every 100
                  if (Graph.Value % 100 == 0)
                  {
                      Refresh();
                      Application.DoEvents();
                   }
              }
           }
        }
	
7. Save your worksheet back to Excel

       // resetup the graph                    
       Graph.Maximum = addresses.Count;
       Graph.Value = 0;

       // change the text
       StatusLabel.Text = "Saving Addresses please wait...";

       // you must convert the list objects to List<IExcelerateObject> before it can be saved
       List<IExcelerateObject> excelerateObjectList = addresses.Cast<IExcelerateObject>().ToList();

       // Now save the worksheet
       SaveWorksheetResponse response = ExcelHelper.SaveWorksheet(excelerateObjectList, addressWorksheet, info, SaveWorksheetCallback, 500);

8. (Optional) Leave a Star on DataJuggler.Excelerate, Blazor Excelerate or this project on GitHub

    DataJuggler.Excelerate
    https://github.com/DataJuggler/Excelerate

    Blazor Excelerate
    https://github.com/DataJuggler/Blazor.Excelerate
	
    Excelerate Win App
    https://github.com/DataJuggler/ExcelerateWinApp

9. (Optional) Subscribe to my YouTube channel
    https://youtube.com/DataJuggler







>>>>>>> 7ef9ac99e2c3a30d070666db40d34ca17fcc657e
