

#region using statements

using DataJuggler.Excelerate;
using System.ComponentModel;
using DataJuggler.Blazor.Components.Util;
using DataJuggler.Blazor.FileUpload;
using Blazor.Excelerate.Models;
using DataJuggler.Net7.Enumerations;
using Microsoft.AspNetCore.Components;
using System.IO.Compression;
using System.Runtime.Versioning;
using Blazor.Excelerate.Enumerations;
using System.Drawing;
using Microsoft.JSInterop;
using System.Text;
using Timer = System.Timers.Timer;

#endregion

namespace Blazor.Excelerate.Pages
{

    #region class Index
    /// <summary>
    /// This class is the main page for this app
    /// </summary>
    [SupportedOSPlatform("Windows")]
    public partial class Index : IBlazorComponentParent, ISpriteSubscriber
    {
        
        #region Private Variables
        private string sideBarElementText;
        private string sideBarSmallText;
        private string sideBarLargeText;
        private string sideBarLargeTextBold;        
        private string downloadLink;
        private string downloadLink2;        
        private string slogan;
        private List<IBlazorComponent> children;        
        private ComboBox sheetNamesComboBox;
        private ImageButton uploadExcelButton;
        private ImageButton generateClassesButton;
        private ImageButton hideButton;
        private Item selectedTextSizeItem;
        private List<Item> selectedSheets;
        private string buttonUrl;
        private List<string> sheetNames;
        private Workbook workbook;
        private List<Item> sheetItems;
        private bool finishedLoading;
        private string selectClasses;                
        private string excelPath;        
        private ValidationComponent namespaceComponent;
        private string status;
        private string statusStyle;
        private string labelColor;
        private List<CodeGenerationResponse> responses;
        private string downloadPath;
        private string mainContent;
        private string smallheader;
        private string instructionsLineHeight;
        private string grid;
        private Sprite logo;
        private string orangeButton;        
        private bool showProgress;
        private double proressPercent;
        private string versionStyle;
        private BackgroundWorker worker;  
        private string displayStyle;
        private string progressStyle;        
        private string percentString;
        private int percent;
        private FileUpload fileUpload;
        private ImageButton resetFileUploadButton;
        private Sprite invisibleSprite;
        private CreateZipFileResponse response;
        private string loadingExamples;        
        private string savingExamples;        
        private MainContentDisplayEnum mainContentDisplay;
        private string instructionButtonUrl;
        private string loadingButtonUrl;
        private string savingButtonUrl;
        private Color instructionButtonTextColor;
        private Color loadingButtonTextColor;
        private Color savingButtonTextColor;
        private string checkVisibility;
        private Timer timer;
        private string checkMarkClassName;
        
        // 20 megs hard coded for now
        private const int UploadLimit = 20971520;
        private const string SampleMemberDataPath = "../Downloads/MemberData.xlsx";
        private const string FileTooLargeMessage = "Your file must be 20 megs or less for this demo.";
        private const string LoadExample = "List<NASDAQ> nasdaqEntries = NASDAQ.Load(worksheet);";
        #endregion

        #region Constructor
        /// <summary>
        /// Create a new instance of an Index page.
        /// </summary>
        public Index()
        {
            // Perform initializations for this object
            Init();
        }
        #endregion

        #region Events

            #region TimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
            /// <summary>
            /// event is fired when Timer Elapsed
            /// </summary>
            private void TimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
            {
                // Hide the checkmark
                CheckVisibility = "hidden";

                // destroy the timer
                Timer.Dispose();

                // Update the UI
                Refresh();
            }
            #endregion

        #endregion

        #region Methods

            #region ButtonClicked(int buttonNumber, string buttonText)
            /// <summary>
            /// This method serves as the ClickHandler for buttons
            /// </summary>
            /// <param name="buttonNumber"></param>
            /// <param name="buttonText"></param>
            public void ButtonClicked(int buttonNumber, string buttonText)
            {
                if ((buttonNumber == 1) && (HasSheetNamesComboBox))
                {
                    // Generate Classes

                    // Handle Generate Classes - starts the background process
                    HandleGenerateClasses();
                }
                else if ((buttonNumber == 2) && (HasFileUpload))
                {
                    // Reset Button

                    // Hide the SheetNamesComboBox
                    DisplayStyle = "none";

                    // In case a zip file is showing, get rid of it
                    Response = null;

                    // Hide the "This class will only be available for the next hour"
                    Status = "";

                    // reset back to gray                     
                    ButtonUrl = "../Images/ButtonDisabled.png";

                    // Reset
                    FileUpload.Reset();
                }
                else if ((buttonNumber == 3) && (HasHideButton))
                {
                    // Hide Instructions

                    // Set this to None until a button is clicked
                    MainContentDisplay = MainContentDisplayEnum.None;

                    // Hide the button
                    HideButton.SetVisible(false);
                }
                else if (buttonNumber == 4)
                {
                    // Show Instructions
                    MainContentDisplay = MainContentDisplayEnum.Instructions;

                    // if the value for HasHideInstructionsButton is true
                    if (HasHideButton)
                    {
                        // Show the button
                        HideButton.SetVisible(true);
                    }
                }
                else if (buttonNumber == 5)
                {
                    // Show Loading Code Examples
                    MainContentDisplay = MainContentDisplayEnum.LoadingExamples;

                    // if the value for HasHideInstructionsButton is true
                    if (HasHideButton)
                    {
                        // Show the button
                        HideButton.SetVisible(true);
                    }
                }
                else if (buttonNumber == 6)
                {
                    // Show Instructions
                    MainContentDisplay = MainContentDisplayEnum.SavingExamples;

                    // if the value for HasHideInstructionsButton is true
                    if (HasHideButton)
                    {
                        // Show the button
                        HideButton.SetVisible(true);
                    }
                }
                else if (buttonNumber == 7)
                {
                    // Copy the code to the clipboard
                    Copy();
                }

                // Update UI
                Refresh();
            }
            #endregion

            #region Copy()
            /// <summary>
            /// Copies the results to the Clipboard
            /// </summary>
            public async void Copy()
            {
                // This is the code sample
                StringBuilder sb = new StringBuilder();
                sb.Append("// Create a new instance of a 'WorksheetInfo' object.\r\n            WorksheetInfo worksheetInfo = new WorksheetInfo();\r\n    \r\n            // set the properties\r\n            worksheetInfo.SheetName = \"NASDAQ\";\r\n            worksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;\r\n            worksheetInfo.Path = \"C:\\\\Projects\\\\GitHub\\\\StockData\\\\Documents\\\\Stocks\\\\NASDAQ.xlsx\";\r\n    \r\n            // load the worksheet\r\n            Worksheet worksheet = ExcelDataLoader.LoadWorksheet(worksheetInfo.Path, worksheetInfo);\r\n    \r\n            // Load the NASDAQ entries\r\n            List<NASDAQ> nasdaqEntries = NASDAQ.Load(worksheet);");
                string code = sb.ToString();

                // Copy
                await BlazorJSBridge.CopyToClipboard(JSRuntime, code);

                 // Show the check mark
                CheckVisibility = "visible";

                // Force UI to update
                Refresh();

                // Start the timer
                Timer = new Timer(3000);
                Timer.Elapsed += TimerElapsed;
                Timer.Start();
            }
            #endregion

            #region ConvertSheetNames()
            /// <summary>
            /// returns a list of Sheet Names
            /// </summary>
            public List<Item> ConvertSheetNames()
            {
                // initial value
                List<Item> items = null;

                // local
                int count = 0;
                
                // If the SheetNames collection exists and has one or more items
                if (ListHelper.HasOneOrMoreItems(SheetNames))
                {
                    // Create a new collection of 'Item' objects.
                    items = new List<Item>();

                    // Iterate the collection of string objects
                    foreach (string item in SheetNames)
                    {
                        // Increment the value for count
                        count++;

                        // Create a new instance of an 'Item' object.
                        Item newItem = new Item();

                        // set the id
                        newItem.Id = count;

                        // Set the sheetName
                        newItem.Text = item;

                        // add this item
                        items.Add(newItem);
                    }
                }

                // return value
                return items;
            }
            #endregion
            
            #region FindChildByName(string name)
            /// <summary>
            /// method Find Child By Name
            /// </summary>
            public IBlazorComponent FindChildByName(string name)
            {
                return ComponentHelper.FindChildByName(children, name);
            }
            #endregion

            #region GetSheetNames(string path)
            /// <summary>
            /// returns the Sheet Names
            /// </summary>
            public Task<List<string>> GetSheetNames(string path)
            {
                // initial value
                List<string> sheetNames = ExcelDataLoader.GetSheetNames(path);
                
                // return value
                return Task.FromResult(sheetNames);
            }
            #endregion

            #region HandleDiscoverSheets(GetSheetNamesModel model)
            /// <summary>
            /// returns the Discover Sheets
            /// </summary>
            public void HandleDiscoverSheets(GetSheetNamesModel model)
            {
                // Create the Worker
                Worker = new BackgroundWorker();

                // Setup the DoWork
                Worker.DoWork += Worker_DoWork;

                // Setup the Completed method
                Worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

                // Start
                Worker.RunWorkerAsync(model);
            }
            #endregion

            #region HandleGenerateClasses()
            /// <summary>
            /// Handle Generating One Or More Classes
            /// </summary>
            public void HandleGenerateClasses()
            {
                // local
                string namespaceName = "";

                // if the NamespaceComponent exists
                if ((HasNamespaceComponent) && (ListHelper.HasOneOrMoreItems(SheetNamesComboBox.SelectedItems)))
                {
                    // Make sure we have a Namespace
                    bool isValid = NamespaceComponent.Validate();

                    // Set the value
                    NamespaceComponent.IsValid = isValid;

                    // Get a list of SheetNames
                    List<string> sheetNames = new List<string>();

                    // Get the text value
                    namespaceName = NamespaceComponent.Text;

                    // iterate the SelectedSheets
                    foreach (Item item in SheetNamesComboBox.SelectedItems)
                    {  
                        // Get the sheetName
                        sheetNames.Add(item.Text);
                    }

                    // if already valid and sheetName and ExcelPath exist
                    isValid = (isValid && ListHelper.HasOneOrMoreItems(sheetNames)) && (FileHelper.Exists(ExcelPath));

                    // if valid
                    if (isValid)
                    {
                        // Show the Progressbar
                        ShowProgress = true;

                       // if the ProgressBar
                       if (HasInvisibleSprite)
                       {
                            // Start the Timer
                            InvisibleSprite.Start();
                        }

                        // Create a new instance of a 'GenerateClassModel' object.
                        GenerateClassModel model = new GenerateClassModel(sheetNames, namespaceName, excelPath);

                        // Launch Background Worker here

                        // Create the Worker
                        Worker = new BackgroundWorker();

                        // Setup the DoWork
                        Worker.DoWork += Worker_DoWork;

                        // Setup the Completed method
                        Worker.RunWorkerCompleted += Worker_RunWorkerCompleted;

                        // Start
                        Worker.RunWorkerAsync(model);

                        // erase any validation messages
                        Status = "";                        
                    }
                    else
                    {
                        // erase
                        Status = "";

                        // this is first
                        if (!FileHelper.Exists(ExcelPath))
                        {
                            // Set Status
                            Status = "Upload an Excel file with a .xlsx extension.";
                        }
                        else if (!NamespaceComponent.IsValid)
                        {
                            // use a red color
                            LabelColor = "tomato";

                            // Set Status
                            Status = "Namespace is required.";
                        }
                        else if (!ListHelper.HasOneOrMoreItems(sheetNames))
                        {
                            // Set the Status if the SheetNames are not selected
                            Status = "Select one or more sheets to continue.";
                        }
                    }                   
                }
            }
            #endregion

            #region Init()
            /// <summary>
            ///  This method performs initializations for this object.
            /// </summary>
            public void Init()
            {
                // Create a new collection of 'IBlazorComponent' objects.
                Children = new List<IBlazorComponent>();
            
                // Start off with the disabled button
                ButtonUrl = "../Images/ButtonDisabled.png";

                // Default to hidden for the ComboBox
                DisplayStyle = "none";

                // Default to Instructions
                MainContentDisplay = MainContentDisplayEnum.Instructions;

                // Hide the checkmark
                CheckVisibility = "hidden";
            }
            #endregion
            
            #region OnFileUploaded(UploadedFileInfo file)
            /// <summary>
            /// This method On File Uploaded
            /// </summary>
            public void OnFileUploaded(UploadedFileInfo file)
            {
                // if the file was uploaded
                if (!file.Aborted)
                {
                   // Show the Progressbar
                   ShowProgress = true;

                   // if the value for HasInvisibleSprite is true
                   if (HasInvisibleSprite)
                   {
                        // Start the Timer
                        InvisibleSprite.Start();
                    }

                   // Create a model
                   GetSheetNamesModel model = new GetSheetNamesModel();

                   // Set the model
                   model.FullPath = file.FullPath;

                    // Store this for later
                    ExcelPath = file.FullPath;

                    // reload the model
                    HandleDiscoverSheets(model);

                    // Update the UI
                    Refresh();
                }
                else
                {
                    // for debugging only
                    if (file.HasException)
                    {
                        // for debugging only
                        string message = file.Exception.Message;
                    }
                }
            }
            #endregion

            #region OnReset()
            /// <summary>
            /// This method On Reset
            /// </summary>
            public void OnReset()
            {  
                // Erase
                Workbook = null;
            }
            #endregion
            
            #region ReceiveData(Message message)
            /// <summary>
            /// method Receive Data
            /// </summary>
            public void ReceiveData(Message message)
            {
                
            }
            #endregion
            
            #region Refresh()
            /// <summary>
            /// method Refresh
            /// </summary>
            public void Refresh()
            {  
                // increment by 4
                Percent += 4;

                // go a little past 100 for effect
                if (Percent >= 100)
                {
                    // Stop the timer
                    InvisibleSprite.Stop();
                    ShowProgress = false;
                }

                // Update the UI
                InvokeAsync(() =>
                {
                    StateHasChanged();
                });
            }
            #endregion
            
            #region Register(Sprite sprite)
            /// <summary>
            /// method returns the
            /// </summary>
            public void Register(Sprite sprite)
            {
                if (sprite.Name == "Logo")
                {
                    // Set the Logo
                    Logo = sprite;
                }
                else
                {
                    // store the InvisibleSprite, used for the progerss bar
                    InvisibleSprite = sprite;
                }
            }
            #endregion
            
            #region Register(IBlazorComponent component)
            /// <summary>
            /// method returns the
            /// </summary>
            public void Register(IBlazorComponent component)
            {
                // Add this item
                this.Children.Add(component);

                // if this is the FileUpload
                if (component is FileUpload)
                {
                    // store the FileUpload
                    FileUpload = component as FileUpload;
                }
                else
                {
                    if (TextHelper.IsEqual(component.Name, "UploadExcelButton"))
                    {
                        // Store this object
                        this.UploadExcelButton = component as ImageButton;

                        // Setup the ClickHandler
                        this.UploadExcelButton.SetClickHandler(ButtonClicked);
                    }
                    else if (TextHelper.IsEqual(component.Name, "GenerateClassesButton"))
                    {
                        // Store this object
                        this.GenerateClassesButton = component as ImageButton;

                        // Setup the ClickHandler
                        this.GenerateClassesButton.SetClickHandler(ButtonClicked);
                    }
                    else if (TextHelper.IsEqual(component.Name, "SheetNamesComboBox"))
                    {
                        // Register the SheetNamesComboBox
                        this.SheetNamesComboBox = component as ComboBox;
                    }
                    else if (TextHelper.IsEqual(component.Name, "NamespaceComponent"))
                    {
                        // Store the NamespaceComponent
                        NamespaceComponent = component as ValidationComponent;                    
                    }
                    else if (TextHelper.IsEqual(component.Name, "HideButton"))
                    {
                        // Register the HideInstructionsButton
                        HideButton = component as ImageButton;

                        // Setup the ClickHandler
                        HideButton.SetClickHandler(ButtonClicked);
                    }        
                     else if (TextHelper.IsEqual(component.Name, "ResetFileUploadButton"))
                    {
                        // Hide the instructions button
                        ResetFileUploadButton = component as ImageButton;

                        // if the value for HasResetFileUploadButton is true
                        if (HasResetFileUploadButton)
                        {
                            // Setup the ClickHandler
                            ResetFileUploadButton.SetClickHandler(ButtonClicked);
                        }
                    }        
                }
            }
            #endregion

            #region Worker_DoWork(object sender, DoWorkEventArgs e)
            /// <summary>
            /// event is fired when Worker _ Do Work
            /// </summary>
            private async void Worker_DoWork(object sender, DoWorkEventArgs e)
            {
                try
                {
                    // Get the model
                    GetSheetNamesModel getSheetNamesModel = e.Argument as GetSheetNamesModel;

                    // if the model exists
                    if (NullHelper.Exists(getSheetNamesModel))
                    {
                        // Store this for later
                        ExcelPath = getSheetNamesModel.FullPath;

                        // Get the SheetNames
                        getSheetNamesModel.SheetNames = await GetSheetNames(getSheetNamesModel.FullPath);

                        // Set Loaded to true
                        getSheetNamesModel.Loaded = ListHelper.HasOneOrMoreItems(getSheetNamesModel.SheetNames);
                    
                        // Set the result
                        e.Result = getSheetNamesModel;
                    }
                    else
                    {
                        // cast as a GenerateClassModel
                        GenerateClassModel generateClassModel = e.Argument as GenerateClassModel;

                        // Get the selected items from the ComboBox
                        SelectedSheets = SheetNamesComboBox.SelectedItems;

                         // Set the outputFolder
                        string outputFolder = Path.GetFullPath("Data");

                        // Create a new string
                        string newFolder = FileHelper.CreateFileNameWithPartialGuid(Path.Combine(outputFolder, "Temp"), 12, false);

                        // Create the directory
                        Directory.CreateDirectory(newFolder);

                        // Set the newFolder
                        generateClassModel.NewFolderPath = newFolder;

                        // if there are one or more sheets
                        if (ListHelper.HasOneOrMoreItems(selectedSheets))
                        {
                            foreach (string sheetName in generateClassModel.SheetNames)
                            {
                                // Create a new instance of a 'LoadWorksheetInfo' object.
                                WorksheetInfo worksheetInfo = new WorksheetInfo();

                                // Set the SheetName
                                worksheetInfo.SheetName = sheetName;

                                // Load all columns
                                worksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;

                                // Load the worksheet
                                Worksheet worksheet = ExcelDataLoader.LoadWorksheet(generateClassModel.ExcelPath, worksheetInfo);

                                // Create a new codeGenerator
                                CodeGenerator codeGenerator = new CodeGenerator(worksheet, newFolder, sheetName);

                                // Generate a class and set the Namespace
                                CodeGenerationResponse response = codeGenerator.GenerateClassFromWorksheet(generateClassModel.NamespaceName, TargetFrameworkEnum.Net7, false);

                                // if the response exists
                                if (NullHelper.Exists(response))
                                {
                                    // Add this response to the Responses collection
                                    generateClassModel.Responses.Add(response);
                                }
                            }

                            // Set the result
                            e.Result = generateClassModel;
                        }
                    }
                }
                catch (Exception error)
                {
                    // for debugging only
                    DebugHelper.WriteDebugError("Worker_DoWork", "Index.razor.cs", error);

                    // Set the error
                    e.Result = error;
                }
            }
            #endregion
            
            #region Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            /// <summary>
            /// event is fired when Worker _ Run Worker Completed
            /// </summary>
            private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                try
                {
                    // hide this 
                    ShowProgress = false;

                    // if the value for HasInvisibleSprite is true
                    if (HasInvisibleSprite)
                    {
                        // Stop the Timer
                        InvisibleSprite.Stop();
                    }
                
                    // Get the PixelQuery
                    GetSheetNamesModel model = e.Result as GetSheetNamesModel;
                    GenerateClassModel generateClassModel = e.Result as GenerateClassModel;

                    // if the PixelQuery exists and one or more pixels were updated
                    if ((NullHelper.Exists(model)) && (model.Loaded))
                    { 
                        // Set the SheetNames
                        SheetNames = model.SheetNames;

                        // Convert the SheetNames to SheetItems
                        SheetItems = ConvertSheetNames();

                        // if there are one or more SheetItems and the ComboBox exists
                        if ((ListHelper.HasOneOrMoreItems(SheetItems)) && (NullHelper.Exists(SheetNamesComboBox)))
                        {
                            // Show the control
                            DisplayStyle = "inline-block";

                            // Start off not expanded
                            SheetNamesComboBox.Expanded = false;

                            // Switch to an enabled OrangeButton
                            ButtonUrl = "../images/OrangeButton.png";

                            // Set the Items
                            SheetNamesComboBox.Items = SheetItems;

                            // Select the first sheet
                            ChangeEventArgs changeEventArgs = new ChangeEventArgs();
                            changeEventArgs.Value = SheetItems[0].Text;
                            SheetNamesComboBox.SelectionChanged(changeEventArgs);                        
                        }                       
                    }
                    else if (NullHelper.Exists(generateClassModel))
                    {
                        // Get the response
                        Responses = generateClassModel.Responses;

                        // Set the newFileName
                        string newFileName = Path.Combine(generateClassModel.NewFolderPath, NamespaceComponent.Text + ".zip");

                        // if a class was created
                        if ((ListHelper.HasOneOrMoreItems(Responses)) && (Responses[0].Success))
                        {
                            // Create the Response
                            Response = new CreateZipFileResponse();

                            // Everything appeared to work
                            Response.Success = true;

                            // Set the Status
                            Status = "This zip file will only be available for download for the next hour.";

                            // reference System.IO.Compression
                            using (var zip = ZipFile.Open(newFileName, ZipArchiveMode.Create))
                            {
                                // Add each file
                                foreach (CodeGenerationResponse response in Responses)
                                {
                                    zip.CreateEntryFromFile(response.FullPath, response.FileName);

                                    // Delete the .cs file
                                    File.Delete(response.FullPath);
                                }
                            }

                            // Create a fileInfo
                            FileInfo fileInfo = new FileInfo(newFileName);

                            // Get the directory info
                            DirectoryInfo directory = new DirectoryInfo(newFileName).Parent;

                            // Now copy the entire folder
                            string destinationFolder = Path.GetFullPath("wwwroot/Downloads/Classes/") + directory.Name;

                            // Create the directory
                            Directory.CreateDirectory(destinationFolder);

                            // Copy the zip file
                            string destinationFileName = Path.Combine(destinationFolder, fileInfo.Name);

                            // Copy
                            File.Copy(newFileName, destinationFileName);

                            // Delete source directory
                            Directory.Delete(generateClassModel.NewFolderPath, true);

                            // Set the DownloadPath
                            DownloadPath = "../Downloads/Classes/" + directory.Name + "/" + fileInfo.Name;

                            // Change the fileName
                            Response.FileName = fileInfo.Name;

                            // Set the FullPath
                            Response.FullPath = DownloadPath;

                            // use white
                            LabelColor = "white";                            
                        }
                        else
                        {
                            // Set the Status
                            Status = "Oops! Something went wrong.";

                            // use a red color
                            LabelColor = "tomato";
                        }
                    }

                    // Update the UI
                    Refresh();
                }
                catch (Exception error)
                {
                    // log the error
                    DebugHelper.WriteDebugError("Worker_RunWorkerCompleted", "Index.razor.cs", error);
                }
                finally
                {   
                    // Erase the DoWork
                    Worker.DoWork -= Worker_DoWork;

                    // Erase the Completed method
                    Worker.RunWorkerCompleted -= Worker_RunWorkerCompleted;

                    // dispose of the worker
                    Worker.Dispose();

                    // destory the reference
                    Worker = null;
                }
            }
            #endregion

        #endregion

        #region Properties

            #region ButtonUrl
            /// <summary>
            /// This property gets or sets the value for 'ButtonUrl'.
            /// </summary>
            public string ButtonUrl
            {
                get { return buttonUrl; }
                set { buttonUrl = value; }
            }
            #endregion
            
            #region CheckMarkClassName
            /// <summary>
            /// This property gets or sets the value for 'CheckMarkClassName'.
            /// </summary>
            public string CheckMarkClassName
            {
                get { return checkMarkClassName; }
                set { checkMarkClassName = value; }
            }
            #endregion
            
            #region CheckVisibility
            /// <summary>
            /// This property gets or sets the value for 'CheckVisibility'.
            /// </summary>
            public string CheckVisibility
            {
                get { return checkVisibility; }
                set { checkVisibility = value; }
            }
            #endregion
            
            #region Children
            /// <summary>
            /// This property gets or sets the value for 'Children'.
            /// </summary>
            public List<IBlazorComponent> Children
            {
                get { return children; }
                set { children = value; }
            }
            #endregion
            
            #region DisplayStyle
            /// <summary>
            /// This property gets or sets the value for 'DisplayStyle'.
            /// </summary>
            public string DisplayStyle
            {
                get { return displayStyle; }
                set { displayStyle = value; }
            }
            #endregion
            
            #region DownloadLink
            /// <summary>
            /// This property gets or sets the value for 'DownloadLink'.
            /// </summary>
            public string DownloadLink
            {
                get { return downloadLink; }
                set { downloadLink = value; }
            }
            #endregion
            
            #region DownloadLink2
            /// <summary>
            /// This property gets or sets the value for 'DownloadLink2'.
            /// </summary>
            public string DownloadLink2
            {
                get { return downloadLink2; }
                set { downloadLink2 = value; }
            }
            #endregion
            
            #region DownloadPath
            /// <summary>
            /// This property gets or sets the value for 'DownloadPath'.
            /// </summary>
            public string DownloadPath
            {
                get { return downloadPath; }
                set { downloadPath = value; }
            }
            #endregion
            
            #region ExcelPath
            /// <summary>
            /// This property gets or sets the value for 'ExcelPath'.
            /// </summary>
            public string ExcelPath
            {
                get { return excelPath; }
                set { excelPath = value; }
            }
            #endregion
            
            #region FileUpload
            /// <summary>
            /// This property gets or sets the value for 'FileUpload'.
            /// </summary>
            public FileUpload FileUpload
            {
                get { return fileUpload; }
                set { fileUpload = value; }
            }
            #endregion
            
            #region FinishedLoading
            /// <summary>
            /// This property gets or sets the value for 'FinishedLoading'.
            /// </summary>
            public bool FinishedLoading
            {
                get { return finishedLoading; }
                set { finishedLoading = value; }
            }
            #endregion
            
            #region GenerateClassesButton
            /// <summary>
            /// This property gets or sets the value for 'GenerateClassesButton'.
            /// </summary>
            public ImageButton GenerateClassesButton
            {
                get { return generateClassesButton; }
                set { generateClassesButton = value; }
            }
            #endregion
            
            #region Grid
            /// <summary>
            /// This property gets or sets the value for 'Grid'.
            /// </summary>
            public string Grid
            {
                get { return grid; }
                set { grid = value; }
            }
            #endregion
            
            #region HasFileUpload
            /// <summary>
            /// This property returns true if this object has a 'FileUpload'.
            /// </summary>
            public bool HasFileUpload
            {
                get
                {
                    // initial value
                    bool hasFileUpload = (this.FileUpload != null);
                    
                    // return value
                    return hasFileUpload;
                }
            }
            #endregion
            
            #region HasHideButton
            /// <summary>
            /// This property returns true if this object has a 'HideButton'.
            /// </summary>
            public bool HasHideButton
            {
                get
                {
                    // initial value
                    bool hasHideButton = (this.HideButton != null);
                    
                    // return value
                    return hasHideButton;
                }
            }
            #endregion
            
            #region HasInvisibleSprite
            /// <summary>
            /// This property returns true if this object has an 'InvisibleSprite'.
            /// </summary>
            public bool HasInvisibleSprite
            {
                get
                {
                    // initial value
                    bool hasInvisibleSprite = (this.InvisibleSprite != null);
                    
                    // return value
                    return hasInvisibleSprite;
                }
            }
            #endregion
            
            #region HasNamespaceComponent
            /// <summary>
            /// This property returns true if this object has a 'NamespaceComponent'.
            /// </summary>
            public bool HasNamespaceComponent
            {
                get
                {
                    // initial value
                    bool hasNamespaceComponent = (this.NamespaceComponent != null);
                    
                    // return value
                    return hasNamespaceComponent;
                }
            }
            #endregion
            
            #region HasResetFileUploadButton
            /// <summary>
            /// This property returns true if this object has a 'ResetFileUploadButton'.
            /// </summary>
            public bool HasResetFileUploadButton
            {
                get
                {
                    // initial value
                    bool hasResetFileUploadButton = (this.ResetFileUploadButton != null);
                    
                    // return value
                    return hasResetFileUploadButton;
                }
            }
            #endregion
            
            #region HasResponse
            /// <summary>
            /// This property returns true if this object has a 'Response'.
            /// </summary>
            public bool HasResponse
            {
                get
                {
                    // initial value
                    bool hasResponse = (this.Response != null);
                    
                    // return value
                    return hasResponse;
                }
            }
            #endregion
            
            #region HasSelectedSheets
            /// <summary>
            /// This property returns true if this object has a 'SelectedSheets'.
            /// </summary>
            public bool HasSelectedSheets
            {
                get
                {
                    // initial value
                    bool hasSelectedSheets = (this.SelectedSheets != null);
                    
                    // return value
                    return hasSelectedSheets;
                }
            }
            #endregion
            
            #region HasSheetNamesComboBox
            /// <summary>
            /// This property returns true if this object has a 'SheetNamesComboBox'.
            /// </summary>
            public bool HasSheetNamesComboBox
            {
                get
                {
                    // initial value
                    bool hasSheetNamesComboBox = (this.SheetNamesComboBox != null);
                    
                    // return value
                    return hasSheetNamesComboBox;
                }
            }
            #endregion
            
            #region HideButton
            /// <summary>
            /// This property gets or sets the value for 'HideButton'.
            /// </summary>
            public ImageButton HideButton
            {
                get { return hideButton; }
                set { hideButton = value; }
            }
            #endregion
            
            #region InstructionButtonTextColor
            /// <summary>
            /// This property gets or sets the value for 'InstructionButtonTextColor'.
            /// </summary>
            public Color InstructionButtonTextColor
            {
                get { return instructionButtonTextColor; }
                set { instructionButtonTextColor = value; }
            }
            #endregion
            
            #region InstructionButtonUrl
            /// <summary>
            /// This property gets or sets the value for 'InstructionButtonUrl'.
            /// </summary>
            public string InstructionButtonUrl
            {
                get { return instructionButtonUrl; }
                set { instructionButtonUrl = value; }
            }
            #endregion
            
            #region InstructionsLineHeight
            /// <summary>
            /// This property gets or sets the value for 'InstructionsLineHeight'.
            /// </summary>
            public string InstructionsLineHeight
            {
                get { return instructionsLineHeight; }
                set { instructionsLineHeight = value; }
            }
            #endregion
            
            #region InvisibleSprite
            /// <summary>
            /// This property gets or sets the value for 'InvisibleSprite'.
            /// </summary>
            public Sprite InvisibleSprite
            {
                get { return invisibleSprite; }
                set { invisibleSprite = value; }
            }
            #endregion
            
            #region LabelColor
            /// <summary>
            /// This property gets or sets the value for 'LabelColor'.
            /// </summary>
            public string LabelColor
            {
                get { return labelColor; }
                set { labelColor = value; }
            }
            #endregion
            
            #region LoadingButtonTextColor
            /// <summary>
            /// This property gets or sets the value for 'LoadingButtonTextColor'.
            /// </summary>
            public Color LoadingButtonTextColor
            {
                get { return loadingButtonTextColor; }
                set { loadingButtonTextColor = value; }
            }
            #endregion
            
            #region LoadingButtonUrl
            /// <summary>
            /// This property gets or sets the value for 'LoadingButtonUrl'.
            /// </summary>
            public string LoadingButtonUrl
            {
                get { return loadingButtonUrl; }
                set { loadingButtonUrl = value; }
            }
            #endregion
            
            #region LoadingExamples
            /// <summary>
            /// This property gets or sets the value for 'LoadingExamples'.
            /// </summary>
            public string LoadingExamples
            {
                get { return loadingExamples; }
                set { loadingExamples = value; }
            }
            #endregion
            
            #region Logo
            /// <summary>
            /// This property gets or sets the value for 'Logo'.
            /// </summary>
            public Sprite Logo
            {
                get { return logo; }
                set { logo = value; }
            }
            #endregion
            
            #region MainContent
            /// <summary>
            /// This property gets or sets the value for 'MainContent'.
            /// </summary>
            public string MainContent
            {
                get { return mainContent; }
                set { mainContent = value; }
            }
            #endregion
            
            #region MainContentDisplay
            /// <summary>
            /// This property gets or sets the value for 'MainContentDisplay'.
            /// </summary>
            public MainContentDisplayEnum MainContentDisplay
            {
                get { return mainContentDisplay; }
                set 
                {
                    // set the value
                    mainContentDisplay = value;

                    // all are blue for MainContentDisplayEnum.None, which only happens if Hide is clicked
                    InstructionButtonUrl = "../Images/Tab.png";
                    LoadingButtonUrl = "../Images/Tab.png";
                    SavingButtonUrl = "../Images/Tab.png";

                    // Now the ButtonTextColors
                    InstructionButtonTextColor = Color.GhostWhite;
                    LoadingButtonTextColor = Color.GhostWhite;
                    SavingButtonTextColor = Color.GhostWhite;

                    // if Instructions Button is selected
                    if (mainContentDisplay == MainContentDisplayEnum.Instructions)
                    {
                        // Instructions is the SelectedTab
                        InstructionButtonUrl = "../Images/TabSelected.png";

                        // Selected Button Text Color
                        InstructionButtonTextColor = Color.Black;
                    }
                    else if (mainContentDisplay == MainContentDisplayEnum.LoadingExamples)
                    {
                        // Loading Examples Button is the SelectedTab
                        LoadingButtonUrl = "../Images/TabSelected.png";

                        // Selected Button Text Color
                        LoadingButtonTextColor = Color.Black;
                    }
                    else if (mainContentDisplay == MainContentDisplayEnum.SavingExamples)
                    {
                        // Saving Examples Button is the SelectedTab
                        SavingButtonUrl = "../Images/TabSelected.png";

                         // Selected Button Text Color
                        SavingButtonTextColor = Color.Black;
                    }
                }
            }
            #endregion
            
            #region NamespaceComponent
            /// <summary>
            /// This property gets or sets the value for 'NamespaceComponent'.
            /// </summary>
            public ValidationComponent NamespaceComponent
            {
                get { return namespaceComponent; }
                set { namespaceComponent = value; }
            }
            #endregion
            
            #region OrangeButton
            /// <summary>
            /// This property gets or sets the value for 'OrangeButton'.
            /// </summary>
            public string OrangeButton
            {
                get { return orangeButton; }
                set { orangeButton = value; }
            }
            #endregion

            #region Percent
            /// <summary>
            /// This property gets or sets the value for 'Percent'.
            /// </summary>
            public int Percent
            {
                get { return percent; }
                set 
                {
                    // if less than zero
                    if (value < 0)
                    {
                        // set to 0
                        value = 0;
                    }

                    // if greater than 100
                    if (value > 100)
                    {
                        // set to 100
                        value = 100;
                    }

                    // set the value
                    percent = value;

                    // Now set ProgressStyle
                    ProgressStyle = "c100 p[Percent] dark small orange".Replace("[Percent]", percent.ToString());

                    // Set the percentString value
                    PercentString = percent.ToString() + "%";
                }
            }
            #endregion
            
            #region PercentString
            /// <summary>
            /// This property gets or sets the value for 'PercentString'.
            /// </summary>
            public string PercentString
            {
                get { return percentString; }
                set { percentString = value; }
            }
            #endregion
            
            #region ProgressStyle
            /// <summary>
            /// This property gets or sets the value for 'ProgressStyle'.
            /// </summary>
            public string ProgressStyle
            {
                get { return progressStyle; }
                set { progressStyle = value; }
            }
            #endregion
            
            #region ProressPercent
            /// <summary>
            /// This property gets or sets the value for 'ProressPercent'.
            /// </summary>
            public double ProressPercent
            {
                get { return proressPercent; }
                set { proressPercent = value; }
            }
            #endregion
            
            #region ResetFileUploadButton
            /// <summary>
            /// This property gets or sets the value for 'ResetFileUploadButton'.
            /// </summary>
            public ImageButton ResetFileUploadButton
            {
                get { return resetFileUploadButton; }
                set { resetFileUploadButton = value; }
            }
            #endregion
            
            #region Response
            /// <summary>
            /// This property gets or sets the value for 'Response'.
            /// </summary>
            public CreateZipFileResponse Response
            {
                get { return response; }
                set { response = value; }
            }
            #endregion
            
            #region Responses
            /// <summary>
            /// This property gets or sets the value for 'Responses'.
            /// </summary>
            public List<CodeGenerationResponse> Responses
            {
                get { return responses; }
                set { responses = value; }
            }
            #endregion
            
            #region SavingButtonTextColor
            /// <summary>
            /// This property gets or sets the value for 'SavingButtonTextColor'.
            /// </summary>
            public Color SavingButtonTextColor
            {
                get { return savingButtonTextColor; }
                set { savingButtonTextColor = value; }
            }
            #endregion
            
            #region SavingButtonUrl
            /// <summary>
            /// This property gets or sets the value for 'SavingButtonUrl'.
            /// </summary>
            public string SavingButtonUrl
            {
                get { return savingButtonUrl; }
                set { savingButtonUrl = value; }
            }
            #endregion
            
            #region SavingExamples
            /// <summary>
            /// This property gets or sets the value for 'SavingExamples'.
            /// </summary>
            public string SavingExamples
            {
                get { return savingExamples; }
                set { savingExamples = value; }
            }
            #endregion
            
            #region SelectClasses
            /// <summary>
            /// This property gets or sets the value for 'SelectClasses'.
            /// </summary>
            public string SelectClasses
            {
                get { return selectClasses; }
                set { selectClasses = value; }
            }
            #endregion
            
            #region SelectedSheets
            /// <summary>
            /// This property gets or sets the value for 'SelectedSheets'.
            /// </summary>
            public List<Item> SelectedSheets
            {
                get { return selectedSheets; }
                set { selectedSheets = value; }
            }
            #endregion
            
            #region SelectedTextSizeItem
            /// <summary>
            /// This property gets or sets the value for 'SelectedTextSizeItem'.
            /// </summary>
            public Item SelectedTextSizeItem
            {
                get { return selectedTextSizeItem; }
                set { selectedTextSizeItem = value; }
            }
            #endregion
            
            #region SheetItems
            /// <summary>
            /// This property gets or sets the value for 'SheetItems'.
            /// </summary>
            public List<Item> SheetItems
            {
                get { return sheetItems; }
                set { sheetItems = value; }
            }
            #endregion
            
            #region SheetNames
            /// <summary>
            /// This property gets or sets the value for 'SheetNames'.
            /// </summary>
            public List<string> SheetNames
            {
                get { return sheetNames; }
                set { sheetNames = value; }
            }
            #endregion
            
            #region SheetNamesComboBox
            /// <summary>
            /// This property gets or sets the value for 'SheetNamesComboBox'.
            /// </summary>
            public ComboBox SheetNamesComboBox
            {
                get { return sheetNamesComboBox; }
                set { sheetNamesComboBox = value; }
            }
            #endregion
            
            #region ShowProgress
            /// <summary>
            /// This property gets or sets the value for 'ShowProgress'.
            /// </summary>
            public bool ShowProgress
            {
                get { return showProgress; }
                set { showProgress = value; }
            }
            #endregion
            
            #region SideBarElementText
            /// <summary>
            /// This property gets or sets the value for 'SideBarElementText'.
            /// </summary>
            public string SideBarElementText
            {
                get { return sideBarElementText; }
                set { sideBarElementText = value; }
            }
            #endregion
            
            #region SideBarLargeText
            /// <summary>
            /// This property gets or sets the value for 'SideBarLargeText'.
            /// </summary>
            public string SideBarLargeText
            {
                get { return sideBarLargeText; }
                set { sideBarLargeText = value; }
            }
            #endregion
            
            #region SideBarLargeTextBold
            /// <summary>
            /// This property gets or sets the value for 'SideBarLargeTextBold'.
            /// </summary>
            public string SideBarLargeTextBold
            {
                get { return sideBarLargeTextBold; }
                set { sideBarLargeTextBold = value; }
            }
            #endregion
            
            #region SideBarSmallText
            /// <summary>
            /// This property gets or sets the value for 'SideBarSmallText'.
            /// </summary>
            public string SideBarSmallText
            {
                get { return sideBarSmallText; }
                set { sideBarSmallText = value; }
            }
            #endregion
            
            #region Slogan
            /// <summary>
            /// This property gets or sets the value for 'Slogan'.
            /// </summary>
            public string Slogan
            {
                get { return slogan; }
                set { slogan = value; }
            }
            #endregion
            
            #region Smallheader
            /// <summary>
            /// This property gets or sets the value for 'Smallheader'.
            /// </summary>
            public string Smallheader
            {
                get { return smallheader; }
                set { smallheader = value; }
            }
            #endregion
            
            #region Status
            /// <summary>
            /// This property gets or sets the value for 'Status'.
            /// </summary>
            public string Status
            {
                get { return status; }
                set { status = value; }
            }
            #endregion
            
            #region StatusStyle
            /// <summary>
            /// This property gets or sets the value for 'StatusStyle'.
            /// </summary>
            public string StatusStyle
            {
                get { return statusStyle; }
                set { statusStyle = value; }
            }
            #endregion
            
            #region Timer
            /// <summary>
            /// This property gets or sets the value for 'Timer'.
            /// </summary>
            public Timer Timer
            {
                get { return timer; }
                set { timer = value; }
            }
            #endregion
            
            #region UploadExcelButton
            /// <summary>
            /// This property gets or sets the value for 'UploadExcelButton'.
            /// </summary>
            public ImageButton UploadExcelButton
            {
                get { return uploadExcelButton; }
                set { uploadExcelButton = value; }
            }
            #endregion
            
            #region VersionStyle
            /// <summary>
            /// This property gets or sets the value for 'VersionStyle'.
            /// </summary>
            public string VersionStyle
            {
                get { return versionStyle; }
                set { versionStyle = value; }
            }
            #endregion
            
            #region Workbook
            /// <summary>
            /// This property gets or sets the value for 'Workbook'.
            /// </summary>
            public Workbook Workbook
            {
                get { return workbook; }
                set { workbook = value; }
            }
            #endregion
            
            #region Worker
            /// <summary>
            /// This property gets or sets the value for 'Worker'.
            /// </summary>
            public BackgroundWorker Worker
            {
                get { return worker; }
                set { worker = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
