

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
        private Item selectedTextSizeItem;
        private const string SampleMemberDataPath = "../Downloads/MemberData.xlsx";
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
        }
        #endregion

        #region Methods

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
                    if ((message.HasSender) && (message.Sender.Name == TextSizeComboBox.Name))
                    {
                        // Set the TextSize
                        TextSize = SetTextSize(message.Text);

                        // Update the UI
                        Refresh();
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
            
        #endregion
        
        #region Properties
            
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
            
        #endregion
        
    }
    #endregion

}
