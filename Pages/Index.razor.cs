

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataJuggler.Blazor.Components;
using DataJuggler.Blazor.Components.Interfaces;
using DataJuggler.UltimateHelper;
using DataJuggler.Blazor.Components.Enumerations;
using Microsoft.AspNetCore.Components;
using DataJuggler.Excelerate;
using DataJuggler.Blazor.FileUpload;
using System.IO;
using System.IO.Compression;
using System.Drawing;

#endregion

namespace Blazor.Excelerate.Pages
{

    #region class Index
    /// <summary>
    /// This is the code for the Index page
    /// </summary>
    public partial class Index : IBlazorComponentParent
    {
        
        #region Private Variables
        private string sideBarElementText;
        private string sideBarSmallText;
        private double textSize;
        private string textSizeStyle;
        private string largeTextSizeStyle;
        private string smallTextSizeStyle;
        private string downloadLink;
        private string downloadLink2;
        private string downloadLink2Hover;
        private string slogan;
        private List<IBlazorComponent> children;
        private ComboBox textSizeComboBox;
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
        private double left;
        private string leftStyle;
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

        // 20 megs hard coded for now
        private const int UploadLimit = 20971520;
        private const string SampleMemberDataPath = "../Downloads/MemberData.xlsx";
        private const string FileTooLargeMessage = "Your file must be 20 megs or less for this demo.";
        #endregion
        
        #region Constructor
        /// <summary>
        /// Create a new instance of an 'Index' object.
        /// </summary>
        public Index()
        {
            // Set the DefautTextSize
            TextSize = 1.8;

            // Create a new collection of 'IBlazorComponent' objects.
            Children = new List<IBlazorComponent>();
            
            // Start off with the disabled button
            ButtonUrl = "../Images/ButtonDisabled.png";

            // Set the Left for the ComboBox container
            Left = -150;

            // set to block
            InstructionsDisplay = "block";
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
                    // Handle creating the class
                    HandleGenerateClass();
                }
                else if ((buttonNumber == 3) && (HasHideInstructionsButton))
                {
                    // Hide
                    InstructionsDisplay = "none";

                    // Hide the button
                    HideInstructionsButton.SetVisible(false);
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
            /// method returns the Child By Name
            /// </summary>
            public IBlazorComponent FindChildByName(string name)
            {
                // initial value
                IBlazorComponent child = null;

                // if the value for HasChildren is true
                if (HasChildren)
                {
                    foreach (IBlazorComponent tempChild in Children)
                    {
                        // if this is the item being sought                        
                        if (TextHelper.IsEqual(tempChild.Name, name))
                        {
                            // set the return value
                            child = tempChild;

                            // break out of loop
                            break;
                        }
                    }
                }
                
                // return value
                return child;
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
                    // Make sure we have a Namespace
                    bool isValid = NamespaceComponent.Validate();

                    // Set the value
                    NamespaceComponent.IsValid = isValid;

                    // if valid
                    if (isValid)
                    {
                        // Get the text value
                        namespaceName = NamespaceComponent.Text;

                        // erase any validation messages
                        Status = "";

                        // Get the sheetName
                        string sheetName = SheetNamesComboBox.ButtonText;

                        // Create a new instance of a 'LoadWorksheetInfo' object.
                        LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                        // Set the SheetName
                        loadWorksheetInfo.SheetName = sheetName;

                        // Load all columns
                        loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;

                        // Load the worksheet
                        Worksheet worksheet = ExcelDataLoader.LoadWorksheet(ExcelPath, loadWorksheetInfo);

                        // Set the outputFolder
                        string outputFolder = Path.GetFullPath("Data");

                        // Create a new string
                        string newFolder = FileHelper.CreateFileNameWithPartialGuid(Path.Combine(outputFolder, "Temp"), 12, false);

                        // Create the directory
                        Directory.CreateDirectory(newFolder);

                        // Create a new codeGenerator
                        CodeGenerator codeGenerator = new CodeGenerator(worksheet, newFolder, sheetName);

                        // Generate a class and set the Namespace
                        Response = codeGenerator.GenerateClassFromWorksheet(namespaceName, false);

                        // Set the newFileName
                        string newFileName = Path.Combine(newFolder, "Excelerate." + sheetName + ".zip");

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
                            Directory.Delete(newFolder, true);

                            // Set the DownloadPath
                            DownloadPath = "../Downloads/Classes/" + directory.Name + "/" + fileInfo.Name;

                            // Change the fileName
                            response.FileName = fileInfo.Name;

                            // Set the FullPath
                            response.FullPath = DownloadPath;

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
                    else
                    {
                        // use a red color
                        LabelColor = "tomato";

                        // Set Status
                        Status = "Namespace is required.";
                    }                   
                }
            }
            #endregion
            
            #region OnAfterRenderAsync(bool firstRender)
            /// <summary>
            /// This method is used to verify a user
            /// </summary>
            /// <param name="firstRender"></param>
            /// <returns></returns>
            protected async override Task OnAfterRenderAsync(bool firstRender)
            {
                // call the base
                await base.OnAfterRenderAsync(firstRender);

                // if FinishedLoading is false and the TextSizeComboBox exists
                if ((!FinishedLoading) && (HasTextSizeComboBox))
                {
                    // Create a new instance of a 'ChangeEventArgs' object.
                    ChangeEventArgs changeEventArgs = new ChangeEventArgs();

                    // Set to Medium
                    changeEventArgs.Value = TextSizeEnum.Medium;

                    // Select Medium
                    TextSizeComboBox.SelectionChanged(changeEventArgs);

                    // Only fire once
                    FinishedLoading = true;
                }
            }
            #endregion
            
            #region OnFileUploaded(UploadedFileInfo file)
            /// <summary>
            /// This method On File Uploaded
            /// </summary>
            public async void OnFileUploaded(UploadedFileInfo file)
            {
                // if the file was uploaded
                if (!file.Aborted)
                {
                    // Store this for later
                    ExcelPath = file.FullPath;

                    // Get the SheetNames
                    this.SheetNames = await GetSheetNames(file.FullPath);

                    // Convert the SheetNames to SheetItems
                    SheetItems = ConvertSheetNames();

                    // if there are one or more SheetItems and the ComboBox exists
                    if ((ListHelper.HasOneOrMoreItems(SheetItems)) && (NullHelper.Exists(SheetNamesComboBox)))
                    {
                        // Now show the control
                        SheetNamesComboBox.SetVisible(true);

                        // Reset
                        Left = 18.8;

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

            #region ReadWorkbook(Path path)
            /// <summary>
            /// This method loads an Excel workbook
            /// </summary>
            /// <param name="workbook"></param>
            /// <returns></returns>
            public Task<bool> ReadWorkbook(string path)
            {
                // initial value
                bool workbookLoaded = false;

                try
                {
                    // Create a new instance of a 'LoadWorksheetInfo' object.
                    LoadWorksheetInfo loadWorksheetInfo = new LoadWorksheetInfo();

                    // Load all columns
                    loadWorksheetInfo.LoadColumnOptions = LoadColumnOptionsEnum.LoadAllColumnsExceptExcluded;
                }
                catch (Exception error)
                {
                    // for debugging only for this demo
                    DebugHelper.WriteDebugError("ReadWorkbook", "Index.razor.cs", error);
                }
                
                // return the value of workbookLoaded
                return Task.FromResult(workbookLoaded);
            }
            #endregion

            #region ReceiveData(Message message)
            /// <summary>
            /// method returns the Data
            /// </summary>
            public void ReceiveData(Message message)
            {
                // if a message exists
                if (NullHelper.Exists(message))
                {
                    // if this message is from the TextSizeComboBox
                    if (message.HasSender)
                    {
                        if (message.Sender.Name == TextSizeComboBox.Name)
                        {
                            // Set the TextSize
                            TextSize = SetTextSize(message.Text);

                            // Update the UI
                            Refresh();
                        }
                    }
                }
            }
            #endregion

            #region Refresh()
            /// <summary>
            /// method Refresh
            /// </summary>
            public void Refresh()
            {
                // Update the UI
                InvokeAsync(() =>
                {
                    StateHasChanged();
                });
            }
            #endregion

            #region Refresh(string message)
            /// <summary>
            /// method returns the
            /// </summary>
            public void Refresh(string message)
            {
                // Update the UI
                InvokeAsync(() =>
                {
                    StateHasChanged();
                });
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

                // if this is the TextSizeComboBox registering
                if (TextHelper.IsEqual(component.Name, "TextSizeComboBox"))
                {
                    // store this object
                    this.TextSizeComboBox = component as ComboBox;

                    // Create the items for TextSizes                    
                    TextSizeComboBox.LoadItems(typeof(TextSizeEnum));
                }
                else if (TextHelper.IsEqual(component.Name, "UploadExcelButton"))
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
                    // Hide the instructions button
                    HideInstructionsButton = component as ImageButton;

                    // Setup the ClickHandler
                    HideInstructionsButton.ClickHandler = ButtonClicked;
                }
            }
            #endregion

            #region SelectionChanged(ChangeEventArgs selectedItem)
            /// <summary>
            /// event is fired when On Change
            /// </summary>            
            public void SelectionChanged(ChangeEventArgs selectedItem)
            {
                // Set the selectedItem
                SelectedTextSizeItem = selectedItem.Value as Item;

                // if exists
                if (HasTextSizeComboBox)
                {
                    // Set the Text
                    TextSizeComboBox.SetButtonText(selectedItem.Value.ToString());
                }

                // Update the UI
                Refresh();
            }
            #endregion
            
            #region SetTextSize(string selectedTextSizeText)
            /// <summary>
            /// Set Text Size
            /// </summary>
            public double SetTextSize(string selectedTextSizeText)
            {
                // Default value
                double textSize = 1.8;

                // If the selectedTextSizeText string exists
                if (TextHelper.Exists(selectedTextSizeText))
                {
                    switch (selectedTextSizeText)
                    {
                        case "Extra Small":

                            // Set the value
                            textSize = 1.2;

                            // required
                            break;

                         case "Small":

                            // Set the value
                            textSize = 1.5;

                            // required
                            break;

                         case "Large":

                            // Set the value
                            textSize = 2.1;

                            // required
                            break;

                        case "Extra Large":

                            // Set the value
                            textSize = 2.4;

                            // required
                            break;
                    }
                }

                // return value
                return textSize;
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
            
            #region DownloadLink2Hover
            /// <summary>
            /// This property gets or sets the value for 'DownloadLink2Hover'.
            /// </summary>
            public string DownloadLink2Hover
            {
                get { return downloadLink2Hover; }
                set { downloadLink2Hover = value; }
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
            
            #region HasChildren
            /// <summary>
            /// This property returns true if this object has a 'Children'.
            /// </summary>
            public bool HasChildren
            {
                get
                {
                    // initial value
                    bool hasChildren = (this.Children != null);
                    
                    // return value
                    return hasChildren;
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
            
            #region HasSelectedTextSizeItem
            /// <summary>
            /// This property returns true if this object has a 'SelectedTextSizeItem'.
            /// </summary>
            public bool HasSelectedTextSizeItem
            {
                get
                {
                    // initial value
                    bool hasSelectedTextSizeItem = (this.SelectedTextSizeItem != null);
                    
                    // return value
                    return hasSelectedTextSizeItem;
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
            
            #region HasTextSizeComboBox
            /// <summary>
            /// This property returns true if this object has a 'TextSizeComboBox'.
            /// </summary>
            public bool HasTextSizeComboBox
            {
                get
                {
                    // initial value
                    bool hasTextSizeComboBox = (this.TextSizeComboBox != null);
                    
                    // return value
                    return hasTextSizeComboBox;
                }
            }
            #endregion
            
            #region HasWorkbook
            /// <summary>
            /// This property returns true if this object has a 'Workbook'.
            /// </summary>
            public bool HasWorkbook
            {
                get
                {
                    // initial value
                    bool hasWorkbook = (this.Workbook != null);
                    
                    // return value
                    return hasWorkbook;
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
            
            #region LargeTextSizeStyle
            /// <summary>
            /// This property gets or sets the value for 'LargeTextSizeStyle'.
            /// </summary>
            public string LargeTextSizeStyle
            {
                get { return largeTextSizeStyle; }
                set { largeTextSizeStyle = value; }
            }
            #endregion
            
            #region Left
            /// <summary>
            /// This property gets or sets the value for 'Left'.
            /// </summary>
            public double Left
            {
                get { return left; }
                set 
                { 
                    left = value;

                    // set the value for LeftStyle
                    LeftStyle = left + "%";
                }
            }
            #endregion
            
            #region LeftStyle
            /// <summary>
            /// This property gets or sets the value for 'LeftStyle'.
            /// </summary>
            public string LeftStyle
            {
                get { return leftStyle; }
                set { leftStyle = value; }
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
            
            #region SmallTextSizeStyle
            /// <summary>
            /// This property gets or sets the value for 'SmallTextSizeStyle'.
            /// </summary>
            public string SmallTextSizeStyle
            {
                get { return smallTextSizeStyle; }
                set { smallTextSizeStyle = value; }
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
            
            #region TextSize
            /// <summary>
            /// This property gets or sets the value for 'TextSize'.
            /// </summary>
            public double TextSize
            {
                get { return textSize; }
                set
                {
                    textSize = value;
                    
                    // Set the textSizeStyle
                    textSizeStyle = textSize + "vh";

                    // get the larger text size
                    double largeTextSize = textSize * 1.5;

                    // Small is half of large
                    double smallTextSize = largeTextSize * .5;

                    // Set the value for LargeTextSizeStyle
                    LargeTextSizeStyle = largeTextSize + "vh";

                    // Set the value for SmallTextSizeStyle
                    SmallTextSizeStyle = smallTextSize + "vh";
                }
            }
            #endregion
                       
            #region TextSizeComboBox
            /// <summary>
            /// This property gets or sets the value for 'TextSizeComboBox'.
            /// </summary>
            public ComboBox TextSizeComboBox
            {
                get { return textSizeComboBox; }
                set { textSizeComboBox = value; }
            }
            #endregion
            
            #region TextSizeStyle
            /// <summary>
            /// This property gets or sets the value for 'TextSizeStyle'.
            /// </summary>
            public string TextSizeStyle
            {
                get { return textSizeStyle; }
                set { textSizeStyle = value; }
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
            
            #region Workbook
            /// <summary>
            /// This property gets or sets the value for 'Workbook'.
            /// </summary>
            public Workbook Workbook
            {
                get { return workbook; }
                set 
                { 
                    workbook = value;

                    if (NullHelper.Exists(workbook))
                    {  
                        // Use Orange Image
                        ButtonUrl = "../Images/OrangeButton.png";
                    }
                    else
                    {
                        // Use Disabled Image
                        ButtonUrl = "../Images/ButtonDisabled.png";
                    }
                }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
