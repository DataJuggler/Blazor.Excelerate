

#region using statements

using System;
using System.Collections.Generic;

#endregion

namespace Blazor.Excelerate.Models
{

    #region class GetSheetNamesModel
    /// <summary>
    /// This class is used to get the sheet names from the path in the background.
    /// 
    /// </summary>
    public class GetSheetNamesModel
    {
        
        #region Private Variables
        private string fullPath;
        private List<string> sheetNames;
        private bool loaded;
        private Exception error;        
        #endregion

        #region Properties
            
            #region Error
            /// <summary>
            /// This property gets or sets the value for 'Error'.
            /// </summary>
            public Exception Error
            {
                get { return error; }
                set { error = value; }
            }
            #endregion
            
            #region FullPath
            /// <summary>
            /// This property gets or sets the value for 'FullPath'.
            /// </summary>
            public string FullPath
            {
                get { return fullPath; }
                set { fullPath = value; }
            }
            #endregion
            
            #region Loaded
            /// <summary>
            /// This property gets or sets the value for 'Loaded'.
            /// </summary>
            public bool Loaded
            {
                get { return loaded; }
                set { loaded = value; }
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
            
        #endregion
        
    }
    #endregion

}
