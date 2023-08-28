

#region using statements


#endregion

namespace Blazor.Excelerate.Models
{

    #region class CreateZipFileResponse
    /// <summary>
    /// This class [Enter Class Description]
    /// </summary>
    public class CreateZipFileResponse
    {
        
        #region Private Variables
        private string fileName;
        private string fullPath;
        private bool success;
        #endregion

        #region Properties
        
            #region FileName
            /// <summary>
            /// This property gets or sets the value for 'FileName'.
            /// </summary>
            public string FileName
            {
                get { return fileName; }
                set { fileName = value; }
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
            
            #region HasFileName
            /// <summary>
            /// This property returns true if the 'FileName' exists.
            /// </summary>
            public bool HasFileName
            {
                get
                {
                    // initial value
                    bool hasFileName = (!String.IsNullOrEmpty(this.FileName));
                    
                    // return value
                    return hasFileName;
                }
            }
            #endregion
            
            #region HasFullPath
            /// <summary>
            /// This property returns true if the 'FullPath' exists.
            /// </summary>
            public bool HasFullPath
            {
                get
                {
                    // initial value
                    bool hasFullPath = (!String.IsNullOrEmpty(this.FullPath));
                    
                    // return value
                    return hasFullPath;
                }
            }
            #endregion
            
            #region Success
            /// <summary>
            /// This property gets or sets the value for 'Success'.
            /// </summary>
            public bool Success
            {
                get { return success; }
                set { success = value; }
            }
            #endregion
            
        #endregion
    }
    #endregion

}
