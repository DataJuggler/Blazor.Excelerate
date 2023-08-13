

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
        private ImageButton hideInstructionsButton;
        private Item selectedTextSizeItem;
        private Item selectedSheetItem;
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
        private CodeGenerationResponse response;
        private string downloadPath;
        private string instructions;
        private string instructionsDisplay;
        private string smallheader;
        private string instructionsLineHeight;
        private string grid;
        private Sprite logo;
        private string orangeButton;
        private const string Column1Width = "22%";
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
        
        // 20 megs hard coded for now
        private const int UploadLimit = 20971520;
        private const string SampleMemberDataPath = "../Downloads/MemberData.xlsx";
        private const string FileTooLargeMessage = "Your file must be 20 megs or less for this demo.";
        #endregion

        #region Constructor
        /// <summary>
        /// Create a new instance of an Index page.
        /// </summary>
        public Index()
        {
            // Create a new collection of 'IBlazorComponent' objects.
            Children = new List<IBlazorComponent>();
            
            // Start off with the disabled button
            ButtonUrl = "../Images/ButtonDisabled.png";

            // set to block
            InstructionsDisplay = "grid";  
            
            // Default to hidden for the ComboBox
            DisplayStyle = "none";
        }
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
                if ((buttonNumber == 2) && (HasSheetNamesComboBox))
                {
                    // Handle Generate Classes - starts the background process
                    HandleGenerateClass();
                }
                else if ((buttonNumber == 3) && (HasHideInstructionsButton))
                {
                    // Hide
                    InstructionsDisplay = "none";

                    // Hide the button
                    HideInstructionsButton.SetVisible(false);
                }
                else if ((buttonNumber == 4) && (HasFileUpload))
                {
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

                // Update UI
                Refresh();
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

            #region HandleGenerateClass()
            /// <summary>
            /// Handle Generate Class
            /// </summary>
            public void HandleGenerateClass()
            {
                // local
                string namespaceName = "";

                // if the NamespaceComponent exists
                if (HasNamespaceComponent)
                {
                    // Get the sheetName
                    string sheetName = SheetNamesComboBox.ButtonText;

                    // Get the text value
                    namespaceName = NamespaceComponent.Text;

                    // Make sure we have a Namespace
                    bool isValid = NamespaceComponent.Validate();

                    // Set the value
                    NamespaceComponent.IsValid = isValid;

                    // if already valid and sheetName and ExcelPath exist
                    isValid = isValid && TextHelper.Exists(sheetName, ExcelPath);

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
                        GenerateClassModel model = new GenerateClassModel(sheetName, namespaceName, excelPath);

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
                        else if (!TextHelper.Exists(sheetName))
                        {
                            // Set Status (should always
                            Status = "Sheet not selected or invalid.";
                        }
                    }                   
                }
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
                        this.UploadExcelButton.ClickHandler = ButtonClicked;
                    }
                    else if (TextHelper.IsEqual(component.Name, "GenerateClassesButton"))
                    {
                        // Store this object
                        this.GenerateClassesButton = component as ImageButton;

                        // Setup the ClickHandler
                        this.GenerateClassesButton.ClickHandler = ButtonClicked;
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
                    else if (TextHelper.IsEqual(component.Name, "HideInstructionsButton"))
                    {
                        // Register the HideInstructionsButton
                        HideInstructionsButton = component as ImageButton;

                        // Setup the ClickHandler
                        HideInstructionsButton.ClickHandler = ButtonClicked;
                    }        
                     else if (TextHelper.IsEqual(component.Name, "ResetFileUploadButton"))
                    {
                        // Hide the instructions button
                        ResetFileUploadButton = component as ImageButton;

                        // if the value for HasResetFileUploadButton is true
                        if (HasResetFileUploadButton)
                        {
                            // Setup the ClickHandler
                            ResetFileUploadButton.ClickHandler = ButtonClicked;
                        }
                    }        
                }
            }
            #endregion

            #region SheetSelected(ChangeEventArgs selectedItem)
            /// <summary>
            /// event is fired when On Change
            /// </summary>            
            public void SheetSelected(ChangeEventArgs selectedItem)
            {
                // Set the selectedItem
                SelectedSheetItem = selectedItem.Value as Item;
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

                        // Create a new instance of a 'LoadWorksheetInfo' object.
                        LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                        // Set the SheetName
                        loadWorksheetInfo.SheetName = generateClassModel.SheetName;

                        // Load all columns
                        loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;

                        // Load the worksheet
                        Worksheet worksheet = ExcelDataLoader.LoadWorksheet(generateClassModel.ExcelPath, loadWorksheetInfo);

                        // Set the outputFolder
                        string outputFolder = Path.GetFullPath("Data");

                        // Create a new string
                        string newFolder = FileHelper.CreateFileNameWithPartialGuid(Path.Combine(outputFolder, "Temp"), 12, false);

                        // Create the directory
                        Directory.CreateDirectory(newFolder);

                        // Set the newFolder
                        generateClassModel.NewFolderPath = newFolder;

                        // Create a new codeGenerator
                        CodeGenerator codeGenerator = new CodeGenerator(worksheet, newFolder, generateClassModel.SheetName);

                        // Generate a class and set the Namespace
                        generateClassModel.Response = codeGenerator.GenerateClassFromWorksheet(generateClassModel.NamespaceName, TargetFrameworkEnum.Net7, false);

                        // Set the result
                        e.Result = generateClassModel;                        
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
                        Response = generateClassModel.Response;

                        // Set the newFileName
                        string newFileName = Path.Combine(generateClassModel.NewFolderPath, "Excelerate." + generateClassModel.SheetName + ".zip");

                        // if a class was created
                        if (Response.Success)
                        {
                            // Set the Status
                            Status = "This class will only be available for download for the next hour.";

                            // reference System.IO.Compression
                            using (var zip = ZipFile.Open(newFileName, ZipArchiveMode.Create))
                            {
                                zip.CreateEntryFromFile(response.FullPath, response.FileName);
                            }

                            // Delete the .cs file
                            File.Delete(response.FullPath);

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
            
            #region HasHideInstructionsButton
            /// <summary>
            /// This property returns true if this object has a 'HideInstructionsButton'.
            /// </summary>
            public bool HasHideInstructionsButton
            {
                get
                {
                    // initial value
                    bool hasHideInstructionsButton = (this.HideInstructionsButton != null);
                    
                    // return value
                    return hasHideInstructionsButton;
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
            
            #region HideInstructionsButton
            /// <summary>
            /// This property gets or sets the value for 'HideInstructionsButton'.
            /// </summary>
            public ImageButton HideInstructionsButton
            {
                get { return hideInstructionsButton; }
                set { hideInstructionsButton = value; }
            }
            #endregion
            
            #region Instructions
            /// <summary>
            /// This property gets or sets the value for 'Instructions'.
            /// </summary>
            public string Instructions
            {
                get { return instructions; }
                set { instructions = value; }
            }
            #endregion
            
            #region InstructionsDisplay
            /// <summary>
            /// This property gets or sets the value for 'InstructionsDisplay'.
            /// </summary>
            public string InstructionsDisplay
            {
                get { return instructionsDisplay; }
                set { instructionsDisplay = value; }
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
            public CodeGenerationResponse Response
            {
                get { return response; }
                set { response = value; }
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
            
            #region SelectedSheetItem
            /// <summary>
            /// This property gets or sets the value for 'SelectedSheetItem'.
            /// </summary>
            public Item SelectedSheetItem
            {
                get { return selectedSheetItem; }
                set { selectedSheetItem = value; }
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
