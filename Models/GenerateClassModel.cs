

#region using statements

using DataJuggler.Excelerate;
using DataJuggler.NET8.Enumerations;

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
        private List<string> sheetNames;
        private string namespaceName;
        private string excelPath;
        private string newFolderPath;
        private TargetFrameworkEnum targetFramework;
        private List<CodeGenerationResponse> responses;
        #endregion
        
        #region Constructor(List<string> sheetNames, string namespaceName, string excelPath, TargetFrameworkEnum targetFramework)
        /// <summary>
        /// Create a new instance of a 'GenerateClassModel' object.
        /// </summary>
        public GenerateClassModel(List<string> sheetNames, string namespaceName, string excelPath, TargetFrameworkEnum targetFramework)
        {
            // store
            SheetNames = sheetNames;
            NamespaceName = namespaceName;
            ExcelPath = excelPath;
            TargetFramework = targetFramework;

            // Create a list to store the Responses
            Responses = new List<CodeGenerationResponse>();
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
            
            #region TargetFramework
            /// <summary>
            /// This property gets or sets the value for 'TargetFramework'.
            /// </summary>
            public TargetFrameworkEnum TargetFramework
            {
                get { return targetFramework; }
                set { targetFramework = value; }
            }
            #endregion
            
        #endregion
        
    }
    #endregion

}
