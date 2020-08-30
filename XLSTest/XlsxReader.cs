using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;
using System.Windows.Forms;

public class XlsxReader : DataSource    
{
    [System.Serializable]
    public class ColumnValueNameWrapper
    {
        public List<DataSourceData> m_ValueNames ;
        public ColumnValueNameWrapper()
        {
            m_ValueNames = new List<DataSourceData>();
        }
    }

    public string m_FileName;
    public XSSFWorkbook m_WorkBook = null;
    ISheet m_Sheet = null;
    int m_NumRows = 0;
    int m_NumCols = 0;
    public List<string> m_ColumnNames = new List<string>();
    bool[] m_ActiveRows;
    bool[] m_FilteredAndAvailableRows;

    const string FORMULA_ERROR_MESSAGE = "Error - Formulas not supported";

    // public List<List<DataSourceData>> m_ColumnValueLists = new List<List<DataSourceData>>();
    public List<ColumnValueNameWrapper> m_ColumnValueLists = new List<ColumnValueNameWrapper>();

    public List<DataSourceData.DataType> m_ColumnTypes = new List<DataSourceData.DataType>();

    Dictionary<string, Vector2> m_LocalMinMaxMap = new Dictionary<string, Vector2>();

    //public override Dictionary<string, Vector2> m_MinMaxMap
    //{
    //    get
    //    {
    //        return m_LocalMinMaxMap;
    //    }
    //}

    Dictionary<string, DataSourceDataMinMax> m_LocalMinMaxDataSourceData = new Dictionary<string, DataSourceDataMinMax>();

    public override Dictionary<string, DataSourceDataMinMax> m_MinMaxDataSourceData
    {
        get
        {
            return m_LocalMinMaxDataSourceData;
        }
    }


    public override DataSourceError CloseDataSource()
    {
        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        m_WorkBook.Close();
        m_WorkBook = null;
        return DataSourceError.OK;
    }

    /// <summary>
    /// XLSX Issue - for other data forms, data starts at row 0.  For XLSX, data starts at row 1.  Need to adjust read to accomodate other systems.
    /// GetCell - used by outside components that don't know the difference between XLS, XLSX, CSV, etc.
    /// GetActualCell - used internally by calls that DO understand that row 0 is column name and data starts on row 1
    /// </summary>
    public DataSourceError GetActualCell(int row, int col, out DataSourceData cellValue)
    {
        return GetCell(row - 1, col, out cellValue);
    }
    public override DataSourceError GetCell(int row, int col, out DataSourceData cellValue)
    {
        row = row + 1;
        DataSourceData data = new DataSourceData();
        cellValue = data;

        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        if ((row < 0) ||
            (row >= m_NumRows) ||
            (col < 0) ||
            (col >= m_NumCols))
            return DataSourceError.ROW_INDEX_OUTOFRANGE;

        ICell cell = m_Sheet.GetRow(row).GetCell(col);
        if (cell == null)
            return DataSourceError.ERROR;

        ICellStyle style = cell.CellStyle;
        short fmtIndex = style.DataFormat;
        string fmt = m_WorkBook.CreateDataFormat().GetFormat(fmtIndex);

        switch (cell.CellType)
        {
            case CellType.Blank:
            case CellType.Error:
                //case CellType.Formula:
                //case CellType.Unknown:
                //data.m_string = cell.StringCellValue;
                data.m_type = DataSourceData.DataType.UNKNOWN;
                data.m_string = "unknown";
                break;

            case CellType.Formula:
                data.m_type = DataSourceData.DataType.STRING;
                data.m_string = FORMULA_ERROR_MESSAGE; // cell.CellFormula;
                if ((m_ColumnValueLists.Count > col) && (m_ColumnValueLists[col] != null) && (m_ColumnValueLists[col].m_ValueNames != null))
                {
                    int index = m_ColumnValueLists[col].m_ValueNames.FindIndex(r => r.m_string.Equals(data.m_string));
                    if (index < 0)
                        data.m_float = 0;
                    else
                        data.m_float = (float)index / (float)(m_ColumnValueLists[col].m_ValueNames.Count - 1);
                }
                break;

            case CellType.Unknown:
                float testVal = 0;
                bool isPercent = false;
                if (GetFloat(cell.StringCellValue, out testVal, out isPercent) == DataSourceError.OK)
                {
                    if (isPercent)
                    {
                        data.m_type = DataSourceData.DataType.PERCENT;
                        data.m_string = data.m_float.ToString("P2");
                    }
                    else
                    {
                        data.m_type = DataSourceData.DataType.FLOAT;
                        data.m_string = data.m_float.ToString();
                    }
                }
                else
                {
                    data.m_type = DataSourceData.DataType.STRING;
                    data.m_string = cell.StringCellValue.Trim();
                }
                break;

            case CellType.Boolean:
                data.m_bool = cell.BooleanCellValue;
                data.m_type = DataSourceData.DataType.BOOLEAN;
                if (data.m_bool)
                {
                    data.m_string = "true";
                    data.m_float = 1.0f;
                }
                else
                {
                    data.m_string = "false";
                    data.m_float = 0.0f;
                }
                break;

            case CellType.Numeric:
                data.m_float = (float)cell.NumericCellValue;
                if (fmt.Contains("%"))
                {
                    data.m_type = DataSourceData.DataType.PERCENT;
                    data.m_string = data.m_float.ToString("P2");
                }
                //else if (fmt.Contains("$"))
                //{
                //    data.m_type = DataSourceData.DataType.DOLLARS;
                //    data.m_string = data.m_float.ToString("C");
                //}
                else
                {
                    data.m_type = DataSourceData.DataType.FLOAT;
                    data.m_string = data.m_float.ToString();
                }
                break;

            case CellType.String:
                data.m_string = cell.StringCellValue;
                data.m_type = DataSourceData.DataType.STRING;
                if ((m_ColumnValueLists.Count > col) && (m_ColumnValueLists[col] != null) && (m_ColumnValueLists[col].m_ValueNames != null))
                {
                    int index = m_ColumnValueLists[col].m_ValueNames.FindIndex(r => r.m_string.Equals(data.m_string));
                    if (index < 0)
                        data.m_float = 0;
                    else
                        data.m_float = (float)index / (float)(m_ColumnValueLists[col].m_ValueNames.Count - 1);
                }
                break;
        }
        return DataSourceError.OK;
    }

    public override DataSourceError SetCell(int row, int col, DataSourceData cellValue)
    {
        row = row + 1;

        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        if ((row < 0) ||
            (row >= m_NumRows) ||
            (col < 0) ||
            (col >= m_NumCols))
            return DataSourceError.ROW_INDEX_OUTOFRANGE;

        ICell cell = m_Sheet.GetRow(row).GetCell(col);
        if (cell == null)
            return DataSourceError.ERROR;

        ICellStyle style = cell.CellStyle;
        short fmtIndex = style.DataFormat;
        string fmt = m_WorkBook.CreateDataFormat().GetFormat(fmtIndex);

        switch (cell.CellType)
        {
            case CellType.Blank:
            case CellType.Error:
                //case CellType.Formula:
                //case CellType.Unknown:
                //data.m_string = cell.StringCellValue;
                if (cellValue.m_type != DataSourceData.DataType.UNKNOWN)
                    return DataSourceError.WRONG_COLUMN_TYPE;
                break;

            case CellType.Formula:
                if (cellValue.m_type != DataSourceData.DataType.STRING)
                    return DataSourceError.WRONG_COLUMN_TYPE;
                cell.SetCellValue(cellValue.m_string);
                break;

            case CellType.Unknown:
                float testVal = 0;
                bool isPercent = false;
                if (GetFloat(cell.StringCellValue, out testVal, out isPercent) == DataSourceError.OK)
                {
                    if (isPercent)
                    {
                        if (cellValue.m_type != DataSourceData.DataType.PERCENT)
                            return DataSourceError.WRONG_COLUMN_TYPE;
                        cell.SetCellValue((double)cellValue.m_float);
                    }
                    else
                    {
                        if (cellValue.m_type != DataSourceData.DataType.FLOAT)
                            return DataSourceError.WRONG_COLUMN_TYPE;
                        cell.SetCellValue((double)cellValue.m_float);
                    }
                }
                else
                {
                    if (cellValue.m_type != DataSourceData.DataType.STRING)
                        return DataSourceError.WRONG_COLUMN_TYPE;
                    cell.SetCellValue(cellValue.m_string);
                }
                break;

            case CellType.Boolean:
                if (cellValue.m_type != DataSourceData.DataType.BOOLEAN)
                    return DataSourceError.WRONG_COLUMN_TYPE;
                if ((cellValue.m_string == "true") || (cellValue.m_float == 1.0f))
                    cell.SetCellValue(true);
                else
                    cell.SetCellValue(false);
                break;

            case CellType.Numeric:
                if (fmt.Contains("%"))
                {
                    if (cellValue.m_type != DataSourceData.DataType.PERCENT)
                        return DataSourceError.WRONG_COLUMN_TYPE;
                    cell.SetCellValue(cellValue.m_float);
                }
                //else if (fmt.Contains("$"))
                //{
                //    data.m_type = DataSourceData.DataType.DOLLARS;
                //    data.m_string = data.m_float.ToString("C");
                //}
                else
                {
                    if (cellValue.m_type != DataSourceData.DataType.FLOAT)
                        return DataSourceError.WRONG_COLUMN_TYPE;
                    cell.SetCellValue(cellValue.m_float);
                }
                break;

            case CellType.String:
                if (cellValue.m_type != DataSourceData.DataType.STRING)
                    return DataSourceError.WRONG_COLUMN_TYPE;
                cell.SetCellValue(cellValue.m_string);
                break;
        }
        return DataSourceError.OK;
    }

    public override DataSourceError GetColumnNames(out string[] columnNames)
    {
        columnNames = null;
        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        columnNames = m_ColumnNames.ToArray();
        return DataSourceError.OK;
    }

    public override DataSourceError GetColumnType(string columnName, out DataSourceData.DataType dataColumnType)
    {
        dataColumnType = DataSourceData.DataType.UNKNOWN;

        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;
        int idx = m_ColumnNames.IndexOf(columnName);
        if (idx == -1)
            return DataSourceError.COLUMN_NAME_NOT_FOUND;

        if (m_ColumnTypes[idx] == DataSourceData.DataType.UNDEFINED)
        {
            DataSourceData data = new DataSourceData() ;

            int typeRowIdx = 1;
            bool gotInitialType = false;
            for (; typeRowIdx < m_NumRows; typeRowIdx++)
            {
                if (GetActualCell(1, idx, out data) != DataSourceError.OK)
                {
                    m_ColumnTypes[idx] = DataSourceData.DataType.UNKNOWN;
                    return DataSourceError.ERROR;
                }

                switch (data.m_type)
                {
                    case DataSourceData.DataType.BOOLEAN:
                        gotInitialType = true;
                        break;
                    case DataSourceData.DataType.DOLLARS:
                        gotInitialType = true;
                        break;
                    case DataSourceData.DataType.FLOAT:
                        gotInitialType = true;
                        break;
                    case DataSourceData.DataType.PERCENT:
                        gotInitialType = true;
                        break;
                    case DataSourceData.DataType.STRING:
                        if (data.m_string.Trim().Length > 0)
                            gotInitialType = true;
                        break;
                    case DataSourceData.DataType.UNDEFINED:
                    case DataSourceData.DataType.UNKNOWN:
                        break;
                }
                if (gotInitialType)
                    break;
            }

            DataSourceData.DataType colType = data.m_type;
            //for (int i = 2; i < m_NumRows; i++)
            //{
            //    if (GetActualCell(i, idx, out data) != DataSourceError.OK)
            //    {
            //        m_ColumnTypes[idx] = DataSourceData.DataType.UNKNOWN;
            //        return DataSourceError.ERROR;
            //    }

            //    if (data.m_type != colType)
            //    {
            //        m_ColumnTypes[idx] = DataSourceData.DataType.UNKNOWN;
            //        return DataSourceError.ERROR;
            //    }
            //}
            m_ColumnTypes[idx] = colType;
            dataColumnType = m_ColumnTypes[idx];
        }
        else
            dataColumnType = m_ColumnTypes[idx];

        return DataSourceError.OK;
    }

    public override DataSourceError GetColumnValueList(string columnName, out DataSourceData[] columnValueList)
    {
        columnValueList = null;

        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        int idx = m_ColumnNames.IndexOf(columnName);
        if (idx == -1)
            return DataSourceError.COLUMN_NAME_NOT_FOUND;

        DataSourceData.DataType dType = DataSourceData.DataType.UNDEFINED;
        GetColumnType(columnName, out dType) ;
        if (m_ColumnValueLists[idx] == null)
        {
            //m_ColumnValueLists[idx] = new List<DataSourceData>();
            m_ColumnValueLists[idx] = new ColumnValueNameWrapper(); // new List<DataSourceData>();
            for (int i = 1; i < m_NumRows; i++)
            {
                DataSourceData data = new DataSourceData();
                GetActualCell(i, idx, out data);
                if (IndexOf(m_ColumnValueLists[idx].m_ValueNames, data) == -1)  
                    m_ColumnValueLists[idx].m_ValueNames.Add(data);
            }
            columnValueList = m_ColumnValueLists[idx].m_ValueNames.ToArray();

            float min = Int64.MaxValue;
            float max = Int64.MinValue;
            DataSourceData minDSD = null;
            DataSourceData maxDSD = null;

            switch (dType)
            {
                case DataSourceData.DataType.STRING:
                    {
                        min = 0;
                        max = 1;
                        //minDSD = m_ColumnValueLists[idx][0];
                        //maxDSD = m_ColumnValueLists[idx][m_ColumnValueLists[idx].Count - 1];
                        minDSD = m_ColumnValueLists[idx].m_ValueNames[0];
                        maxDSD = m_ColumnValueLists[idx].m_ValueNames[m_ColumnValueLists[idx].m_ValueNames.Count - 1];
                        for (int i = 0; i < m_ColumnValueLists[idx].m_ValueNames.Count; i++)
                        {
                            int len = m_ColumnValueLists[idx].m_ValueNames.Count;
                            if (len > 1)
                            {
                                m_ColumnValueLists[idx].m_ValueNames[i].m_float = (float)i / (float)(len - 1);
                            }
                            else if (len == 1)
                                m_ColumnValueLists[idx].m_ValueNames[i].m_float = 0.5f;
                            else
                                m_ColumnValueLists[idx].m_ValueNames[i].m_float = 0;

                        }
                    }
                    break;

                case DataSourceData.DataType.FLOAT:
                case DataSourceData.DataType.DOLLARS:
                case DataSourceData.DataType.PERCENT:
                    {

                    }
                    break; 
            }
            columnValueList = m_ColumnValueLists[idx].m_ValueNames.ToArray();
        }
        else
        {
            columnValueList = m_ColumnValueLists[idx].m_ValueNames.ToArray();
        }
        return DataSourceError.OK;
    }

    public override DataSourceError GetNumberOfRows(out int numRows)
    {
        numRows = 0;

        if (m_WorkBook == null)
            return DataSourceError.DATASOURCE_NOT_OPEN;

        // this should reflect number of DATA rows 
        numRows = m_NumRows-1;
        return DataSourceError.OK;
    }

    public bool SanityCheck(ISheet sheet, out int rowWithError)
    {
        rowWithError = 0;

        if (sheet == null)
            return false;

        int firstRow = sheet.FirstRowNum;
        int lastRow = sheet.LastRowNum;
        int numRows = (lastRow - firstRow) + 1;

        if (numRows == 0)
            return false;
        try
        {
            int numCols = sheet.GetRow(0).LastCellNum;
            for (int i = 0; i <= numCols; i++)
            {
                ICell cell = sheet.GetRow(0).GetCell(i);
                if (cell == null)
                {
                    numCols = i; // - 1;
                    break;
                }
            }

            for (int i = 0; i < numCols; i++)
            {
                ICell cell = sheet.GetRow(0).GetCell(i);
                if (cell.CellType != CellType.String)
                    return false;

            }
            for (int i = 0; i < numRows; i++)
            {
                if(sheet.GetRow(i) != null)
                {
                    int lastCellInRow = sheet.GetRow(i).LastCellNum;
                    for (int j = 0; j <= lastCellInRow; j++)
                    {
                        ICell cell = sheet.GetRow(i).GetCell(j);
                        if (cell == null)
                        {
                            lastCellInRow = j - 1;
                            break;
                        }
                    }
                    if (lastCellInRow > numCols)
                    {
                        rowWithError = i;
                        return false;
                    }
                }
            }
        }
        catch (SystemException)
        {
            return false;
        }

        return true;
    }
    public override DataSourceError OpenDataSource(string uri)
    {
        Debug.Log("OpenDataSource");
        try
        {
            if (m_WorkBook != null)
            {
                m_WorkBook.Close();
                m_WorkBook = null;
            }
            m_WorkBook = new XSSFWorkbook(uri);
        }
        catch (Exception e)
        {
            Debug.Log("Failed to open |" + uri + "| " + e.ToString());
            return DataSourceError.DATASOURCE_NOT_FOUND;
        }

        if (m_WorkBook == null)
        {
            return DataSourceError.DATASOURCE_NOT_FOUND;
        }
        else
        {
            m_Sheet = m_WorkBook.GetSheetAt(0);
            if (m_Sheet == null)
            {
                m_WorkBook.Close();
                m_WorkBook = null;
                return DataSourceError.DATASOURCE_NOT_OPEN;
            }
            else
            {
                int rowWithError = 0;
                if (!SanityCheck(m_Sheet, out rowWithError))
                {
                    // put up dialog box:
                    //  XLS Format:
                    //  Data must be in the first sheet
                    //  Data must be rectangular block of cells
                    //  Data must start at row 1, column 1
                    //  No blank rows or blank columns
                    //  Row 1 must contain the column names
                    //  No other data should be on this sheet
                    Debug.LogError("SanityCheck for \"" + uri + "\" failed - row " + (rowWithError +1).ToString());
                    return DataSourceError.PARSE_ERROR;
                }

                int firstRow = m_Sheet.FirstRowNum;
                int lastRow = m_Sheet.LastRowNum;
                int iRow = 0;
                int lastCell = m_Sheet.GetRow(iRow).LastCellNum;
                for (int j = 0; j <= lastCell; j++)
                {
                    ICell cell = m_Sheet.GetRow(iRow).GetCell(j);
                    if (cell == null)
                    {
                        lastCell = j - 1;
                        break;
                    }
                }

                for (iRow = lastRow; iRow >= firstRow ; iRow--)
                {
                    if (m_Sheet.GetRow(iRow) != null)
                    {
                        bool isEmpty = false;
                        for (int j = 0; j < lastCell; j++)
                        {
                            ICell cell = m_Sheet.GetRow(iRow).GetCell(j);
                            if (cell != null)
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Blank:
                                    case CellType.Error:
                                    case CellType.Unknown:
                                        isEmpty = true;
                                        break;

                                    case CellType.Formula:
                                        if (cell.CellFormula.Trim().Length == 0)
                                            isEmpty = true;
                                        break;

                                    case CellType.String:
                                        if (cell.StringCellValue.Trim().Length == 0)
                                            isEmpty = true;
                                        break;
                                }

                                if (!isEmpty)
                                    break;
                            }
                            else
                                isEmpty = true;

                            if (!isEmpty)
                                break;
                        }
                        if (!isEmpty)
                            break;
                    }
                }
                lastRow = iRow;

                m_NumRows = (lastRow - firstRow) + 1;
                //m_NumCols = m_Sheet.GetRow(0).LastCellNum;
                m_NumCols = lastCell+1;

                for (int i = 0; i < m_NumCols; i++)
                {
                    ICell cell = m_Sheet.GetRow(0).GetCell(i);
                    if (cell.CellType != CellType.String)
                    {
                        m_WorkBook.Close();
                        m_WorkBook = null;
                        return DataSourceError.DATASOURCE_NOT_OPEN;
                    }
                    else
                    {
                        m_ColumnNames.Add(cell.StringCellValue);
                        m_ColumnValueLists.Add(null); // set up for lazy evaluation
                        m_ColumnTypes.Add(DataSourceData.DataType.UNDEFINED); // set up for lazy evaluation 
                    }
                }

                // alas, I cannot use lazy evaluation - must do so here
                for (int col = 0; col < m_NumCols; col++)
                {
                    DataSourceData[] columnValues = null;
                    GetColumnValueList(m_ColumnNames[col], out columnValues);
                    float min = Int64.MaxValue;
                    float max = Int64.MinValue;
                    DataSourceData minDSD = null;
                    DataSourceData maxDSD = null;

                    DataSourceData.DataType columnDataType = DataSourceData.DataType.UNDEFINED;

                    for (int row = 1; row < m_NumRows; row++)
                    {
                        DataSourceData data = new DataSourceData();
                        if (columnDataType == DataSourceData.DataType.UNDEFINED)
                            columnDataType = data.m_type;

                        if (GetActualCell(row, col, out data) == DataSourceError.OK)
                        {
                            if ((data.m_type == DataSourceData.DataType.DOLLARS) ||
                                (data.m_type == DataSourceData.DataType.FLOAT) ||
                                (data.m_type == DataSourceData.DataType.PERCENT))
                            {
                                if (data.m_float < min)
                                {
                                    min = data.m_float;
                                    minDSD = data;
                                }
                                if (data.m_float > max)
                                {
                                    max = data.m_float;
                                    maxDSD = data;
                                }
                            }
                            else if (data.m_type == DataSourceData.DataType.BOOLEAN)
                            {
                                data.m_float = data.m_bool ? 1 : 0;
                            }
                            else if (data.m_type == DataSourceData.DataType.STRING)
                            {
                                //int index = IndexOf(columnValues, data);
                                int index = Array.FindIndex(columnValues, r => r.m_string == data.m_string);

                                if (index < 0)
                                    data.m_float = 0;
                                else
                                {
                                    int len = columnValues.Length;
                                    if (len > 1)
                                    {
                                        data.m_float = (float)index / (float)(len - 1);
                                    }
                                    else if (len == 1)
                                        data.m_float = 0.5f;
                                    else
                                        data.m_float = 0;

                                    if (data.m_float < min)
                                    {
                                        min = data.m_float;
                                        minDSD = data;
                                    }
                                    if (data.m_float > max)
                                    {
                                        max = data.m_float;
                                        maxDSD = data;
                                    }
                                }
                            }
                        }
                    }
                    if (columnDataType == DataSourceData.DataType.PERCENT)
                    {
                        min = min / 100.0f;
                        max = max / 100.0f;
                    }
                    m_LocalMinMaxMap[m_ColumnNames[col]] = new Vector2(min, max);
                    DataSourceDataMinMax dsdmm = new DataSourceDataMinMax();
                    dsdmm.Min = minDSD;
                    dsdmm.Max = maxDSD;
                    if (columnDataType == DataSourceData.DataType.PERCENT)
                    {
                        dsdmm.Min.m_float = dsdmm.Min.m_float / 100.0f;
                        dsdmm.Min.m_float = dsdmm.Min.m_float / 100.0f;
                    }
                    m_LocalMinMaxDataSourceData[m_ColumnNames[col]] = dsdmm;

                }

            }
        }
        //is number of rows -1 since the first row doesn't count;
        m_ActiveRows = new bool[m_NumRows-1];
        m_FilteredAndAvailableRows = new bool[(m_NumRows-1)];
        for (int i = 0; i < m_ActiveRows.Length; i++)
            m_ActiveRows[i] = true;

        for (int i = 0; i < m_FilteredAndAvailableRows.Length; i++)
            m_FilteredAndAvailableRows[i] = true;
      

        return DataSourceError.OK;
    }

    public bool Init()
    {
        DataSourceError result =  OpenDataSource(m_FileName);
        if (result != DataSourceError.OK)
        {
            return false;
        }
        //for (int row = 0; row < m_NumRows; row++)
        //{
        //    for (int col = 0; col < m_NumCols; col++)
        //    {
        //        DataSourceData data = new DataSourceData();
        //        GetActualCell(row, col, out data);
        //    }
        //}
        return true;
    }

    public override DataSourceError GetActiveRows(out bool[] activeRows)
    {
        activeRows = m_ActiveRows;
        return DataSourceError.OK;
    }

    public override DataSourceError SetActiveRows(bool[] activeRows)
    {
        if (activeRows.Length > m_ActiveRows.Length)
            return DataSourceError.ERROR;

        for (int i = 0; i < m_ActiveRows.Length; i++)
            m_ActiveRows[i] = activeRows[i];

        return DataSourceError.OK;
    }
    

    public override DataSourceError RecalcMinimumsAndMaximums(bool[] ignore = null)
    {
        for (int col = 0; col < m_ColumnNames.Count; col++)
        {
            List<DataSourceData> columnAsDataSourceData = new List<DataSourceData>();
            DataSourceData.DataType columnDataType = m_ColumnTypes[col];

            float min = float.MaxValue;
            float max = float.MinValue;

            DataSourceData minDSD = new DataSourceData();
            DataSourceData maxDSD = new DataSourceData();
            //bool minAndMaxInited = false;

            List<DataSourceData> valueNames = m_ColumnValueLists[col].m_ValueNames;
            //for (int row = 0; row < m_NumRows; row++)
            /* HACK FIX */ for (int row = 0; row < (m_NumRows-1); row++)
            {
                if (!(m_FilteredAndAvailableRows[row] && m_ActiveRows[row]))
                    continue;

                DataSourceData cell = new DataSourceData();
                GetCell(row, col, out cell);
                DataSourceData data = new DataSourceData();
                if ((columnDataType == DataSourceData.DataType.FLOAT) ||
                    (columnDataType == DataSourceData.DataType.PERCENT))
                {
                    float parseVal = 0;
                    //bool isPercent = false;
                    data.m_type = columnDataType;
                    parseVal = cell.m_float;

                    //get min and max of each column

                    data.m_float = parseVal;
                    data.m_string = parseVal.ToString();

                    if (parseVal < min)
                    {
                        min = parseVal;
                        minDSD = data;
                    }

                    if (parseVal > max)
                    {
                        max = parseVal;
                        maxDSD = data;
                    }

                }
                else
                {
                    data.m_type = DataSourceData.DataType.STRING;
                    data.m_string = cell.m_string;

                    int index = valueNames.FindIndex(r => r.m_string == data.m_string);
                    if (index < 0)
                        data.m_float = 0;
                    else
                    {
                        int len = valueNames.Count;
                        if (len > 1)
                        {
                            data.m_float = (float)index / (float)(len - 1);
                        }
                        else if (len == 1)
                            data.m_float = 0.5f;
                        else
                            data.m_float = 0;

                        if (data.m_float < min)
                        {
                            min = data.m_float;
                            minDSD = data;
                        }
                        if (data.m_float > max)
                        {
                            max = data.m_float;
                            maxDSD = data;
                        }
                    }
                }
            }
            m_LocalMinMaxMap[m_ColumnNames[col]] = new Vector2(min, max);
            DataSourceDataMinMax dsdmm = new DataSourceDataMinMax();
            dsdmm.Min = minDSD;
            dsdmm.Max = maxDSD;
            m_LocalMinMaxDataSourceData[m_ColumnNames[col]] = dsdmm;

        }
        return DataSourceError.OK;
    }

    public override DataSourceError GetFilteredAndAvailableRows(out bool[] filteredRow)
    {
        filteredRow = m_FilteredAndAvailableRows;
        return DataSourceError.OK;
    }

    public override DataSourceError SetFilteredAndAvailableRows(bool[] filteredRow)
    {
        if (filteredRow.Length > m_FilteredAndAvailableRows.Length)
            return DataSourceError.ERROR;

        for (int i = 0; i < m_FilteredAndAvailableRows.Length; i++)
            m_FilteredAndAvailableRows[i] = filteredRow[i];

        return DataSourceError.OK;
    }

    public override DataSourceError GetFullColumnValueList(string columnName, out DataSourceData[] fullColumnValueList)
    {
        int index = m_ColumnNames.IndexOf(columnName);
        if (index == -1)
        {
            fullColumnValueList = null;
            return DataSourceError.COLUMN_NAME_NOT_FOUND;
        }

        int numberOfRows;
        GetNumberOfRows(out numberOfRows);
        fullColumnValueList = new DataSourceData[numberOfRows];
        for (int i = 0; i < numberOfRows; i++)
        {
            DataSourceData dsd;
            GetCell(i, index, out dsd);
            fullColumnValueList[i] = dsd;
        }
        return DataSourceError.OK;
    }

    public override DataSourceError GetFilteredAndActiveRows(out bool[] filteredAndActiveRows)
    {

        BitArray activeBitArray = new BitArray(this.m_ActiveRows);
      

        BitArray filterAndAvailableBitArray = new BitArray(this.m_FilteredAndAvailableRows);
        bool[] ret = new bool[filterAndAvailableBitArray.Length];
        BitArray result;
        try
        {
            result = filterAndAvailableBitArray.And(activeBitArray);
        }
        catch(ArgumentException e)
        {
            Debug.LogError("BitArray length in DataSource not matching. Error Message: " + e.StackTrace);
            filteredAndActiveRows = new bool[0];
            return DataSourceError.ROW_INDEX_OUTOFRANGE;
        }
        
       
      
        result.CopyTo(ret, 0);

        filteredAndActiveRows = ret;
        return DataSourceError.OK;
    }

    public override DataSourceError GetFilteredAndAvailableRowCount(out int activeCount)
    {
        if (m_FilteredAndAvailableRows == null || m_FilteredAndAvailableRows.Length == 0)
        {
            activeCount = 0;
            return DataSourceError.ERROR;
        }
        int count = 0;
        for (int i = 0; i < m_FilteredAndAvailableRows.Length; i++)
        {
            if (m_FilteredAndAvailableRows[i])
            {
                count++;
            }
        }
        activeCount = count;
        return DataSourceError.OK;
    }

  public override string GetDataSourceFileName() {
    return m_FileName;
  }
}