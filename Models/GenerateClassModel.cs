

#region using statements

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DataJuggler.Excelerate;

#endregion

namespace Blazor.Excelerate.Models
{

    #region class GenerateClassModel
    /// <summary>
    /// This class is here so the generate class is processed in the background
    /// so the progress bar works on the main thread.
    /// </summary>
    public class GenerateClassModel
    {
        
        #region Private Variables
        private string sheetName;
        private string namespaceName;
        private string excelPath;
        private string newFolderPath;
        private CodeGenerationResponse response;
        #endregion
        
        #region Constructor(string sheetName, string namespaceName, string excelPath)
        /// <summary>
        /// Create a new instance of a 'GenerateClassModel' object.
        /// </summary>
        public GenerateClassModel(string sheetName, string namespaceName, string excelPath)
        {
            // store
            SheetName = sheetName;
            NamespaceName = namespaceName;
            ExcelPath = excelPath;
        }
        #endregion
        
        #region Properties
            
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
            
            #region NamespaceName
            /// <summary>
            /// This property gets or sets the value for 'NamespaceName'.
            /// </summary>
            public string NamespaceName
            {
                get { return namespaceName; }
                set { namespaceName = value; }
            }
            #endregion
            
            #region NewFolderPath
            /// <summary>
            /// This property gets or sets the value for 'NewFolderPath'.
            /// </summary>
            public string NewFolderPath
            {
                get { return newFolderPath; }
                set { newFolderPath = value; }
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
            
            #region SheetName
            /// <summary>
            /// This property gets or sets the value for 'SheetName'.
            /// </summary>
            public string SheetName
            {
                get { return sheetName; }
                set { sheetName = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
