

#region using statements

using DataJuggler.Excelerate;
using DataJuggler.NET.Data;
using DataJuggler.Excelerate.Interfaces;

#endregion

namespace Blazor.Excelerate.Models
{

    #region class NASDAQ : IExcelerateObject
    public class NASDAQ : IExcelerateObject
    {

        #region Private Variables
        private string changedColumns;
        private string industry;
        private int iPOYear;
        private bool loading;
        private string name;
        private Guid rowId;
        private string sector;
        private string symbol;
        private int volume;
        #endregion

        #region Methods

            #region Load(Row row)
            /// <summary>
            /// This method loads a NASDAQ object from a Row.
            /// </Summary>
            /// <param name="row">The row which the row.Columns[x].ColumnValue will be used to load this object.</param>
            public void Load(Row row)
            {
                // If the row exists and the row's column collection exists
                if ((NullHelper.Exists(row)) && (row.HasColumns))
                {
                    // Turn Loading On
                    Loading = true;

                    // set values
                    Symbol = row.Columns[0].StringValue;
                    Name = row.Columns[1].StringValue;
                    IPOYear = row.Columns[2].IntValue;
                    Volume = row.Columns[3].IntValue;
                    Sector = row.Columns[4].StringValue;
                    Industry = row.Columns[5].StringValue;

                    // Turn Loading Off
                    Loading = false;
                }

                // Set RowId
                RowId = row.Id;
            }
            #endregion

            #region Load(Worksheet worksheet)
            /// <summary>
            /// This method loads a list of NASDAQ objects from a Worksheet.
            /// </Summary>
            /// <param name="worksheet">The worksheet which the rows collection will be used to load a list of NASDAQ objects.</param>
            public static List<NASDAQ> Load(Worksheet worksheet)
            {
                // Initial value
                List<NASDAQ> nASDAQList = new List<NASDAQ>();
                
                // If the worksheet exists and the row's collection exists
                if ((NullHelper.Exists(worksheet)) && (worksheet.HasRows))
                {
                    // Iterate the worksheet.Rows collection
                    foreach (Row row in worksheet.Rows)
                    {
                        // If the row is not a HeaderRow and row's column collection exists
                        if ((!row.IsHeaderRow) && (row.HasColumns))
                        {
                            // Create a new instance of a NASDAQ object.
                            NASDAQ nASDAQ = new NASDAQ();
                            
                            // Load this object
                            nASDAQ.Load(row);
                            
                            // Add this object to the list
                            nASDAQList.Add(nASDAQ);
                        }
                    }
                }
                
                // return value
                return nASDAQList;
            }
            #endregion

            #region NewRow(Row row)
            /// <summary>
            /// This method creates the columns for the row to save a new NASDAQ object.
            /// </Summary>
            /// <param name="row">The row which the Columns will be created for.</param>
            public static Row NewRow(int rowNumber)
            {
                // initial value
                Row newRow = new Row();

                // Create Column
                Column symbolColumn = new Column("Symbol", rowNumber, 1, DataManager.DataTypeEnum.String);

                // Add this column
                newRow.Columns.Add(symbolColumn);

                // Create Column
                Column nameColumn = new Column("Name", rowNumber, 2, DataManager.DataTypeEnum.String);

                // Add this column
                newRow.Columns.Add(nameColumn);

                // Create Column
                Column iPOYearColumn = new Column("IPOYear", rowNumber, 3, DataManager.DataTypeEnum.Integer);

                // Add this column
                newRow.Columns.Add(iPOYearColumn);

                // Create Column
                Column volumeColumn = new Column("Volume", rowNumber, 4, DataManager.DataTypeEnum.Integer);

                // Add this column
                newRow.Columns.Add(volumeColumn);

                // Create Column
                Column sectorColumn = new Column("Sector", rowNumber, 5, DataManager.DataTypeEnum.String);

                // Add this column
                newRow.Columns.Add(sectorColumn);

                // Create Column
                Column industryColumn = new Column("Industry", rowNumber, 6, DataManager.DataTypeEnum.String);

                // Add this column
                newRow.Columns.Add(industryColumn);

                // return value
                return newRow;
            }
            #endregion

            #region Save(Row row)
            /// <summary>
            /// This method saves a NASDAQ object back to a Row.
            /// </Summary>
            /// <param name="row">The row which the row.Columns[x].ColumnValue will be set to Save back to Excel.</param>
            public Row Save(Row row)
            {
                // If the row exists and the row's column collection exists and the ChangedColumns string is not null or empty
                if ((NullHelper.Exists(row)) && (row.HasColumns) && (TextHelper.Exists(ChangedColumns)))
                {
                    // Parse the changed column indexes
                    List<int> changedColumnIndexes = ExcelHelper.ParseChangedColumnIndexes(ChangedColumns);

                    row.Columns[0].ColumnValue = Symbol;
                    row.Columns[0].HasChanges = changedColumnIndexes.Contains(0);
                    row.Columns[1].ColumnValue = Name;
                    row.Columns[1].HasChanges = changedColumnIndexes.Contains(1);
                    row.Columns[2].ColumnValue = IPOYear;
                    row.Columns[2].HasChanges = changedColumnIndexes.Contains(2);
                    row.Columns[3].ColumnValue = Volume;
                    row.Columns[3].HasChanges = changedColumnIndexes.Contains(3);
                    row.Columns[4].ColumnValue = Sector;
                    row.Columns[4].HasChanges = changedColumnIndexes.Contains(4);
                    row.Columns[5].ColumnValue = Industry;
                    row.Columns[5].HasChanges = changedColumnIndexes.Contains(5);
                }

                // return value
                return row;
            }
            #endregion

        #endregion

        #region Properties

            #region string ChangedColumns
            public string ChangedColumns
            {
                get
                {
                    return changedColumns;
                }
                set
                {
                    changedColumns = value;
                }
            }
            #endregion

            #region string Industry
            public string Industry
            {
                get
                {
                    return industry;
                }
                set
                {
                    industry = value;

                    if (!Loading)
                    {
                        ChangedColumns += 5 + ",";
                    }
                }
            }
            #endregion

            #region int IPOYear
            public int IPOYear
            {
                get
                {
                    return iPOYear;
                }
                set
                {
                    iPOYear = value;

                    if (!Loading)
                    {
                        ChangedColumns += 2 + ",";
                    }
                }
            }
            #endregion

            #region bool Loading
            public bool Loading
            {
                get
                {
                    return loading;
                }
                set
                {
                    loading = value;
                }
            }
            #endregion

            #region string Name
            public string Name
            {
                get
                {
                    return name;
                }
                set
                {
                    name = value;

                    if (!Loading)
                    {
                        ChangedColumns += 1 + ",";
                    }
                }
            }
            #endregion

            #region Guid RowId
            public Guid RowId
            {
                get
                {
                    return rowId;
                }
                set
                {
                    rowId = value;

                    if (!Loading)
                    {
                        ChangedColumns += 6 + ",";
                    }
                }
            }
            #endregion

            #region string Sector
            public string Sector
            {
                get
                {
                    return sector;
                }
                set
                {
                    sector = value;

                    if (!Loading)
                    {
                        ChangedColumns += 4 + ",";
                    }
                }
            }
            #endregion

            #region string Symbol
            public string Symbol
            {
                get
                {
                    return symbol;
                }
                set
                {
                    symbol = value;

                    if (!Loading)
                    {
                        ChangedColumns += 0 + ",";
                    }
                }
            }
            #endregion

            #region int Volume
            public int Volume
            {
                get
                {
                    return volume;
                }
                set
                {
                    volume = value;

                    if (!Loading)
                    {
                        ChangedColumns += 3 + ",";
                    }
                }
            }
            #endregion

        #endregion

    }
    #endregion

}