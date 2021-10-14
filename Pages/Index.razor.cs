

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
        private string downloadLink;
        private string slogan;
        private List<IBlazorComponent> children;
        private ComboBox textSizeComboBox;
        private ComboBox sheetNamesComboBox;
        private ImageButton uploadExcelButton;
        private ImageButton generateClassesButton;
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
                if (buttonNumber == 1)
                {
                    // Test only
                    this.Workbook = new Workbook();

                    // Update the UI
                    Refresh();
                }
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

                    // Set the value for largeTextSize
                    largeTextSizeStyle = largeTextSize + "vh";
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
