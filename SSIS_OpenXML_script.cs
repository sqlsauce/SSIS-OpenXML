/* Microsoft SQL Server Integration Services Script Component
*  Write scripts using Microsoft Visual C# 2008.
*  ScriptMain is the entry point class of the script.*/

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{

    #region Script Setup
    // Ensure you've set up output columns and supplied an Excel Connection Manager in the Script Editor UI
    /// <summary>
    /// Set the following variable value to the name of the table in Excel.  To find or set the name of the table in Excel,
    /// go to the Design ribbon in the Table Tools group.  The table name is shown on the left side.
    /// </summary>
    private static readonly string ExcelTableName = "IF - COA - LNS - Centnl";
    /// <summary>
    /// Set this variable to true if you want copious reporting done (for debugging)
    /// </summary>
    private static readonly bool VerboseLogging = true;
    /// <summary>
    /// Fill this method with calls to MapColumn to map Excel column names to SSIS output columns, and provide data types.
    /// </summary>
    private void MapColumns()
    {
        // sample:
        // this.MapColumn("Excel Column Header", "SSIS Column Name", typeof(string));
        this.MapColumn("FA", "FA", typeof(string), true);
        this.MapColumn("USID", "USID", typeof(int), true);
        this.MapColumn("OOF_Vendor_Type", "OOF_Vendor_Type", typeof(string), true); // OOF VENDOR TYPE
        // this.MapColumn("ECD", "ECD", typeof(string), true);
        this.MapColumn("ECD", "ECD", typeof(int), true);
        this.MapColumn("Revised ECD", "Revised_ECD", typeof(string), true);
        this.MapColumn("ACD", "ACD", typeof(string), true);
        this.MapColumn("ACD 2", "ACD_2", typeof(string), true);
        this.MapColumn("Reported", "Reported", typeof(string), true);
    }
    #endregion

    #region Code You Don't Touch
    // The following code is configured based on the information you supplied in the section above,
    // and what columns and connection in the Script Editor UI.
    private const string SCRIPT_NAME = "OpenXML API Script Source for SpreadsheetML";
    private const string LAST_UPDATED = "2013-11-20 23:50";
    private static bool __script_last_updated_logged = false;
    private static IDTSComponentMetaData100 __metadata;
    /// <summary>
    /// The list of Excel to SSIS column maps
    /// </summary>
    private readonly List<ColumnMapping> _columnMappings = new List<ColumnMapping>();
    #endregion

    #region CLASS: ColumnMapping
    private class ColumnMapping
    {
        #region Property Setting Delegates
        public delegate void NullSetter(bool isNull);
        public delegate void StringSetter(string value);
        public delegate void Int32Setter(Int32 value);
        public delegate void DateTimeSetter(DateTime value);
        public delegate void BooleanSetter(bool value);
        //ADD_DATATYPES_HERE
        #endregion

        #region Private Variables
        private readonly string _excelColumnName;
        private int _excelColumnOffset;
        private readonly string _ssisColumnName;
        private readonly System.Type _dataType;
        private readonly bool _treatBlanksAsNulls;
        #endregion

        public NullSetter SetNull;
        public StringSetter SetString;
        public Int32Setter SetInt;
        public DateTimeSetter SetDateTime;
        public BooleanSetter SetBoolean;
        //ADD_DATATYPES_HERE

        #region Constructor
        public ColumnMapping(string excelColumnName, string ssisColumnName, System.Type dataType, bool treatBlanksAsNulls)
        {
            this._excelColumnName = excelColumnName;
            this._excelColumnOffset = -1;
            this._ssisColumnName = ssisColumnName;
            this._dataType = dataType;
            this._treatBlanksAsNulls = treatBlanksAsNulls;
        }
        #endregion

        #region Public Properties
        public System.Type DataType
        {
            get { return this._dataType; }
        }

        public string ExcelColumnName
        {
            get { return this._excelColumnName; }
        }

        public string SSISColumnName
        {
            get { return this._ssisColumnName; }
        }

        public int ExcelColumnOffset
        {
            get { return this._excelColumnOffset; }
            set { this._excelColumnOffset = value; }
        }

        public bool ExcelColumnFound
        {
            get { return (this._excelColumnOffset >= 0); }
        }

        public bool TreatBlanksAsNulls
        {
            get { return this._treatBlanksAsNulls; }
        }
        #endregion

        #region SSIS Buffer Setter
        public void SetSSISBuffer(string value)
        {
            #region String
            if (this._dataType == typeof(string))
            {
                try
                {
                    this.SetString(value);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered setting SSIS column '" + this._ssisColumnName + "' with string value '" + value + "': " + ex.Message, true);
                }
                #endregion
                VerboseLog("Set SSIS column '" + this._ssisColumnName + "' with string value '" + value + "'");
            }
            #endregion
            #region Int
            else if (this._dataType == typeof(int))
            {
                int intValue = 0;
                try
                {
                    intValue = Convert.ToInt32(value);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered converting Excel column '" + this._excelColumnName + "' value '" + value + "' to integer: " + ex.Message, true);
                }
                #endregion
                try
                {
                    this.SetInt(intValue);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered setting SSIS column '" + this._ssisColumnName + "' with int value '" + intValue.ToString() + "': " + ex.Message, true);
                }
                #endregion
                VerboseLog("Set SSIS column '" + this._ssisColumnName + "' with int value '" + intValue.ToString() + "'");
            }
            #endregion
            #region DateTime
            else if (this._dataType == typeof(DateTime))
            {
                DateTime dateValue = new DateTime(1900, 1, 1);
                try
                {
                    dateValue = dateValue.AddDays(Convert.ToDouble(value) - 2);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered converting Excel column '" + this._excelColumnName + "' value '" + value + "' to DateTime: " + ex.Message, true);
                }
                #endregion
                try
                {
                    this.SetDateTime(dateValue);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered setting SSIS column '" + this._ssisColumnName + "' with DateTime value '" + dateValue.ToString() + "': " + ex.Message, true);
                }
                #endregion
                VerboseLog("Set SSIS column '" + this._ssisColumnName + "' with DateTime value '" + dateValue.ToString("yyyy-MM-dd hh:mm:ss") + "'");
            }
            #endregion
            #region Boolean
            else if (this._dataType == typeof(bool))
            {
                bool boolValue = false;
                try
                {
                    if ((value.ToUpper().Trim() == "YES")
                        || (value.ToUpper().Trim() == "Y")
                        || (value.ToUpper().Trim() == "TRUE"))
                    {
                        boolValue = true;
                    }
                    else if ((value.ToUpper().Trim() == "NO")
                        || (value.ToUpper().Trim() == "N")
                        || (value.ToUpper().Trim() == "FALSE"))
                    {
                        boolValue = false;
                    }
                    else
                    {
                        ReportError("Invalid boolean value in column '" + this._excelColumnName + "': '" + value + "'", true);
                    }
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered converting Excel column '" + this._excelColumnName + "' value '" + value + "' to boolean: " + ex.Message, true);
                }
                #endregion
                try
                {
                    this.SetBoolean(boolValue);
                }
                #region catch...
                catch (Exception ex)
                {
                    ReportError("Error encountered setting SSIS column '" + this._ssisColumnName + "' with boolean value '" + boolValue.ToString() + "': " + ex.Message, true);
                }
                #endregion
                VerboseLog("Set SSIS column '" + this._ssisColumnName + "' with Boolean value '" + boolValue.ToString() + "'");
            }
            #endregion
            //ADD_DATATYPES_HERE
            else
            {
                ReportUnhandledDataTypeError(this._dataType);
            }
        }
        #endregion
    }
    #endregion

    #region Sets up map from Excel column to an SSIS column
    public void MapColumn(string excelColumnName, string ssisColumnName, System.Type dataType, bool treatBlanksAsNulls)
    {
        string methodName = "set_" + ssisColumnName.Replace(" ", "");
        VerboseLog("Creating " + dataType.ToString() + " mapping from '" + excelColumnName + "' to '" + ssisColumnName + "' via " + methodName);
        ColumnMapping mapping = new ColumnMapping(excelColumnName, ssisColumnName, dataType, treatBlanksAsNulls);
        #region Code to create delegates I'd have liked to have inside the ColumnMapping class itself if I could pass Output0Buffer...
        mapping.SetNull = (ColumnMapping.NullSetter)Delegate.CreateDelegate(typeof(ColumnMapping.NullSetter), Output0Buffer, methodName + "_IsNull");
        if (dataType == typeof(string))
        {
            mapping.SetString = (ColumnMapping.StringSetter)Delegate.CreateDelegate(typeof(ColumnMapping.StringSetter), Output0Buffer, methodName);
        }
        else if (dataType == typeof(int))
        {
            mapping.SetInt = (ColumnMapping.Int32Setter)Delegate.CreateDelegate(typeof(ColumnMapping.Int32Setter), Output0Buffer, methodName);
        }
        else if (dataType == typeof(DateTime))
        {
            mapping.SetDateTime = (ColumnMapping.DateTimeSetter)Delegate.CreateDelegate(typeof(ColumnMapping.DateTimeSetter), Output0Buffer, methodName);
        }
        else if (dataType == typeof(bool))
        {
            mapping.SetBoolean = (ColumnMapping.BooleanSetter)Delegate.CreateDelegate(typeof(ColumnMapping.BooleanSetter), Output0Buffer, methodName);
        }
        //ADD_DATATYPES_HERE
        else
        {
            ReportUnhandledDataTypeError(dataType);
        }
        this._columnMappings.Add(mapping);
        #endregion
    }
    #endregion

    #region CreateNewOutputRows - the only method called from SSIS, this is the "entry point"
    public override void CreateNewOutputRows()
    {
        #region Set up verbose logging
        __metadata = ComponentMetaData;
        #endregion

        // See http://blogs.msdn.com/b/brian_jones/archive/2008/11/10/reading-data-from-spreadsheetml.aspx
        // http://openxmldeveloper.org/discussions/formats/f/14/p/5029/157797.aspx
        // http://blogs.msdn.com/b/ericwhite/archive/2010/07/21/table-markup-in-open-xml-spreadsheetml.aspx

        #region Configure Column Mapping
        this.MapColumns();
        if (this._columnMappings.Count != ComponentMetaData.OutputCollection[0].OutputColumnCollection.Count)
        {
            string message = this._columnMappings.Count.ToString() + " column relationships have been set up, but the Script Source has "
                + ComponentMetaData.OutputCollection[0].OutputColumnCollection.Count.ToString() + " output columns defined.";
            ReportError(message, true);
            throw new ArgumentException(message);
        }
        VerboseLog(this._columnMappings.Count.ToString() + " column mappings defined.");
        #endregion

        #region Extract Excel file name from connection manager
        string workbookFileName = null;
        SpreadsheetDocument document = null;
        try
        {
            VerboseLog("Extracting Excel file name from connection manager.");
            string connectionString = ComponentMetaData.RuntimeConnectionCollection[0].ConnectionManager.ConnectionString;
            string[] connectionStringParts = connectionString.Split(';');
            foreach (string connectionStringPart in connectionStringParts)
            {
                string[] pair = connectionStringPart.Split('=');
                if (pair[0] == "Data Source")
                {
                    workbookFileName = pair[1];
                    VerboseLog("File name of '" + workbookFileName + "' identified in connection manager.");
                    break;
                }
            }
        }
        #region catch ...
        catch (Exception ex)
        {
            ReportError("Unable to parse connection string: " + ex.Message, true);
        }
        #endregion
        #endregion
        #region Opening Excel file
        if (workbookFileName != null)
        {
            try
            {
                VerboseLog("Attempting to open Excel file.");
                document = SpreadsheetDocument.Open(workbookFileName, false);
            }
            #region catch ...
            catch (Exception ex)
            {
                ReportError("Unable to open '" + workbookFileName + "': " + ex.Message, true);
            }
            #endregion
        }
        VerboseLog("Excel file opened.");
        #endregion
        try
        {
            WorkbookPart workbook = document.WorkbookPart;
            SharedStringTablePart sharedStringTablePart = workbook.SharedStringTablePart;
            #region Unused code for finding ranges
            //ComponentMetaData.FireInformation(0, "", "Got WorkbookPart", "", 0, ref fireAgain);
            //#region Look at Ranges
            //bool foundRange = false;
            //RangeDef rangeDef = new RangeDef();
            //foreach (DefinedName name in workbook.Workbook.GetFirstChild<DefinedNames>())
            //{
            //    ComponentMetaData.FireInformation(0, "", "Looking at defined name '" + name.Name + "'", "", 0, ref fireAgain);
            //    if (name.Name == this._rangeName)
            //    {
            //        ComponentMetaData.FireInformation(0, "", "Saving def", "", 0, ref fireAgain);
            //        rangeDef.Name = name.Name;
            //        string reference = name.InnerText;
            //        ComponentMetaData.FireInformation(0, "", "  reference: " + reference, "", 0, ref fireAgain);
            //        rangeDef.Sheet = reference.Split('!')[0].Trim('\'');
            //        string[] rangeArray = reference.Split('!')[1].Split('$');
            //        rangeDef.StartCol = rangeArray[1];
            //        rangeDef.StartRow = rangeArray[2].TrimEnd(':');
            //        rangeDef.EndCol = rangeArray[3];
            //        rangeDef.EndRow = rangeArray[4];
            //        foundRange = true;
            //        break;
            //    }
            //}
            //ComponentMetaData.FireInformation(0, "", "Done looking for defined names", "", 0, ref fireAgain);
            //if (foundRange)
            //{
            //    string rangeID = workbook.Workbook.Descendants<Sheet>().Where(r => r.Name.Equals(rangeDef.Sheet)).First().Id;
            //    ComponentMetaData.FireInformation(0, "", "Got rangeID " + rangeID, "", 0, ref fireAgain);
            //    WorksheetPart range = (WorksheetPart)workbook.GetPartById(rangeID);
            //    ComponentMetaData.FireInformation(0, "", "Got Range", "", 0, ref fireAgain);
            //}
            //#endregion
            #endregion
            #region Iterate over sheets to find table
            VerboseLog("Searching sheets for table '" + ExcelTableName + "'.");
            Table table = null;
            Worksheet worksheet = null;
            foreach (Sheet sheet in workbook.Workbook.Sheets)
            {
                VerboseLog("Examining sheet '" + sheet.Name + "'.");
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                foreach (TableDefinitionPart tableDefinitionPart in worksheetPart.TableDefinitionParts)
                {
                    VerboseLog("Sheet contains table '" + tableDefinitionPart.Table.DisplayName + "'.");
                    if (tableDefinitionPart.Table.DisplayName == ExcelTableName)
                    {
                        worksheet = worksheetPart.Worksheet;
                        table = tableDefinitionPart.Table;
                        VerboseLog("Sheet and table found.");
                        break;
                    }
                }
                if (table != null)
                {
                    break;
                }
            }
            #endregion
            if (table == null)
            {
                ReportError("Table '" + ExcelTableName + "' wasn't found in '" + workbookFileName + "'.", true);
            }
            else
            {
                string firstColumnHeader = "";
                #region Find Excel Column Offsets
                VerboseLog("Collecting column offsets for mapped columns.");
                int columnIndex = 1;
                foreach (TableColumn tableColumn in table.TableColumns)
                {
                    if (columnIndex == 1)
                    {
                        firstColumnHeader = tableColumn.Name;
                    }
                    foreach (ColumnMapping columnRelationship in this._columnMappings)
                    {
                        if (tableColumn.Name == columnRelationship.ExcelColumnName)
                        {
                            VerboseLog("Found Excel column " + tableColumn.Name + " at offset " + columnIndex.ToString());
                            columnRelationship.ExcelColumnOffset = columnIndex;
                            break;
                        }
                    }
                    columnIndex++;
                }
                #region Throw an error if not all columns were found
                foreach (ColumnMapping columnRelationship in this._columnMappings)
                {
                    if (!columnRelationship.ExcelColumnFound)
                    {
                        string message = "Unable to locate column '" + columnRelationship.ExcelColumnName + "' in table '" + ExcelTableName + "'.";
                        ReportError(message, true);
                    }
                }
                #endregion
                #endregion
                #region Read spreadsheet data into SSIS output buffer
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Elements<Row>();
                #region Find First Row
                UInt32 firstRow = 0;
                VerboseLog("Finding first row of table.");
                foreach (Row row in rows)
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        if (this.CellReferenceToCoordinates(cell.CellReference)[0] == 1)
                        {
                            if (this.GetCellValue(cell, sharedStringTablePart) == firstColumnHeader)
                            {
                                firstRow = row.RowIndex + 1;
                            }
                        }
                    }
                }
                VerboseLog("First row of table is on row " + firstRow.ToString() + ".");
                #endregion
                VerboseLog("Preparing to read " + rows.Count<Row>().ToString() + " table rows from Excel.");
                foreach (Row row in rows)
                {
                    VerboseLog("Reading row " + row.RowIndex.ToString() + ".");
                    if (row.RowIndex < firstRow)
                    {
                        VerboseLog("Skipping non-table or header row.");
                    }
                    else
                    {
                        VerboseLog("Reading data row " + (row.RowIndex - 1).ToString() + ".");
                        bool rowAdded = false;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            foreach (ColumnMapping columnRelationship in this._columnMappings)
                            {
                                if (this.CellReferenceToCoordinates(cell.CellReference)[0] == columnRelationship.ExcelColumnOffset)
                                {
                                    string cellValue = this.GetCellValue(cell, sharedStringTablePart);
                                    if ((cellValue == null)
                                        || ((cellValue == "") && columnRelationship.TreatBlanksAsNulls))
                                    {
                                        // do nothing
                                    }
                                    else
                                    {
                                        if (!rowAdded)
                                        {
                                            Output0Buffer.AddRow();
                                            rowAdded = true;
                                        }
                                        VerboseLog("Excel column '" + columnRelationship.ExcelColumnName + "' contains '" + cellValue + "'.");
                                        columnRelationship.SetSSISBuffer(this.GetCellValue(cell, sharedStringTablePart));
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion
            }
        }
        #region catch ...
        catch (Exception ex)
        {
            ReportError("Unable to open Excel file using OpenXML API: " + ex.Message, true);
        }
        #endregion
        VerboseLog("Closing Excel file");
        document.Close();
    }
    #endregion

    #region Helper functions to change Excel ranges to numeric coordinates/offsets
    private int[] RangeReferenceToCoordinates(string rangeReference)
    {
        int[] coordinates = new int[4];

        string[] cellReferences = rangeReference.Split(':');
        int[] startCellReference = this.CellReferenceToCoordinates(cellReferences[0]);
        int[] endCellReference = this.CellReferenceToCoordinates(cellReferences[1]);

        coordinates[0] = startCellReference[0];
        coordinates[1] = startCellReference[1];
        coordinates[2] = endCellReference[0];
        coordinates[3] = endCellReference[1];

        return coordinates;
    }

    private int[] CellReferenceToCoordinates(string cellReference)
    {
        //bool fireAgain = true;
        int[] coordinates = new int[2];

        //ComponentMetaData.FireInformation(0, "", "CellRef: [" + cellReference + "]", "", 0, ref fireAgain);
        int index;
        cellReference = cellReference.Replace("$", "").Trim();
        #region Collect column letters -> column
        string column = "";
        for (index = 0; index < cellReference.Length; index++)
        {
            if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(cellReference[index]))
            {
                column += cellReference[index];
            }
            else
            {
                break;
            }
        }
        #endregion
        //ComponentMetaData.FireInformation(0, "", "column: [" + column + "]", "", 0, ref fireAgain);
        #region Convert column into number -> coordinates[0]
        for (int power = 0; power < index; power++)
        {
            coordinates[0] += ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(column[column.Length - power - 1]) + 1) * (int)Math.Pow(26, power);
        }
        #endregion
        //ComponentMetaData.FireInformation(0, "", "col coords: [" + coordinates[0].ToString() + "]", "", 0, ref fireAgain);
        #region Convert row into number -> coordinates[1]
        coordinates[1] = Convert.ToInt32(cellReference.Substring(index));
        #endregion
        //ComponentMetaData.FireInformation(0, "", "row coords: [" + coordinates[1].ToString() + "]", "", 0, ref fireAgain);

        return coordinates;
    }
    #endregion

    #region Helper function to read Excel cell values
    private string GetCellValue(Cell cell, SharedStringTablePart sharedStringTablePart)
    {
        if (cell.ChildElements.Count == 0)
        {
            return null;
        }
        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
        {
            return sharedStringTablePart.SharedStringTable.ChildElements[Int32.Parse(cell.CellValue.InnerText)].InnerText;
        }
        else
        {
            return cell.CellValue.InnerText;
        }
    }
    #endregion

    #region Logging functions
    private static void VerboseLog(string message)
    {
        bool pbFireAgain = true;

        if (!__script_last_updated_logged)
        {
            __metadata.FireInformation(0, __metadata.Name, "Excel OpenXML API Source Script (" + LAST_UPDATED + ") running.", "", 0, ref pbFireAgain);
            __script_last_updated_logged = true;
        }
        if (VerboseLogging)
        {
            __metadata.FireInformation(0, __metadata.Name, message, "", 0, ref pbFireAgain);
        }
    }

    private static void ReportError(string message, bool fatal)
    {
        bool pbCancel;
        __metadata.FireError(0, __metadata.Name, message, "", 0, out pbCancel);
        if (fatal)
        {
            throw new ApplicationException(SCRIPT_NAME + " had a fatal error.");
        }
    }

    private static void ReportUnhandledDataTypeError(System.Type dataType)
    {
        // Need to add a Data Type?  Search for ADD_DATATYPES_HERE
        ReportError("This script can't handle " + dataType.ToString() + " types.", true);
        throw new ArgumentException("This script can't handle " + dataType.ToString() + " types.");
    }
    #endregion

    //public override void PreExecute()
    //{
    //    base.PreExecute();
    //    /*
    //      Add your code here for preprocessing or remove if not needed
    //    */
    //}

    //public override void PostExecute()
    //{
    //    base.PostExecute();
    //    /*
    //      Add your code here for postprocessing or remove if not needed
    //      You can set read/write variables here, for example:
    //      Variables.MyIntVar = 100
    //    */
    //}

    //public override void CreateNewOutputRows()
    //{
    //    /*
    //      Add rows by calling the AddRow method on the member variable named "<Output Name>Buffer".
    //      For example, call MyOutputBuffer.AddRow() if your output was named "MyOutput".
    //    */
    //}

}
