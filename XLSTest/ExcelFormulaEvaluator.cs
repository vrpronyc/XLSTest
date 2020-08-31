using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;

using ExcelFormulaParser;
using System.Reflection;
using System.Data;

using Toolsbox.ShuntingYard;

namespace XLSTest
{
    [System.Serializable]
    public enum FormulaReturnType { stringFormula, floatFormula };
    public class FormulaReturnValue
    {
        public FormulaReturnType returnType;
        public string stringValue;
        public float floatValue;
    }

    public class ExcelFormulaEvaluator
    {
        XSSFWorkbook m_WorkBook;
        ISheet m_Sheet;

        public ExcelFormulaEvaluator(XSSFWorkbook workbook)
        {
            m_WorkBook = workbook;
        }

        public float FetchNumericValueFromCell(ICell cell, out bool valid)
        {
            valid = false;
            ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
            if (cell == null)
            {
                Console.WriteLine("cell is null");
                return 0f;
            }
            if (cell.CellType != CellType.Numeric)
            {
                Console.WriteLine("cell is not numeric");
                return 0;
            }
            valid = true;
            return (float)(cell.NumericCellValue);

        }
        public ICell FetchCellFromSheet(string sheetName, int row, int col)
        {
            ISheet sheet = m_WorkBook.GetSheet(sheetName);
            if (sheet == null)
            {
                Console.WriteLine("Could not get sheet \"" + sheetName + "\"");
                return null;
            }

            IRow cellRow = sheet.GetRow(row);
            if (cellRow == null)
            {
                Console.WriteLine($"Could not get row {row} from sheet \"{sheetName}\"");
                return null;
            }
            int iStartCellIdx = cellRow.FirstCellNum;
            ICell cell = cellRow.GetCell(col);
            if (cell == null)
            {
                Console.WriteLine($"Could not get cell {row},{col} from sheet \"{sheetName}\"");
                return null;
            }

            return cell;
        }

        [System.Serializable]
        public enum CellArgumentType {
            NA,
            CAString,       // just a string
            CAValue,        // ##.## 
            CACell,         // LL##
            CARange,        // LL##:LL##
            CASheetCell,    // str!LL##
            CASheetRange    // str!LL##:LL##
        };

        public class CellArgument
        {
            public CellArgumentType caType;
            public string stringVal;
            public float val;
            public int row0;
            public int col0;
            public int row1;
            public int col1;
            public string sheetName;

            public CellArgument()
            {
                caType = CellArgumentType.NA;
                stringVal = string.Empty;
                val = 0;
                row0 = 0;
                col0 = 0;
                row1 = 0;
                col1 = 0;
                sheetName = string.Empty;
            }
        }

        static string CellValueAsString(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    return "Unknown";
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString("R");
                case CellType.String:
                    return cell.StringCellValue; 
                case CellType.Formula:
                    return cell.CellFormula;
                case CellType.Blank:
                    return "";
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return "Error";
                default:
                    return "default";
            }
        }

        // formats of valid cell argument:
        // ##.##
        // LL##
        // LL##:LL##
        // str!LL##
        // str!LL##:LL##
        static CellArgument ParseRange(string rangeSourceString, string argString)
        {
            CellArgument cellArg = new CellArgument();

            //Not simple cell, not sheet string - must be range!
            string[] rangeString = rangeSourceString.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            if (rangeString.Length != 2)
            {
                //Not simple cell, not sheet string, not range - return as string
                cellArg.caType = CellArgumentType.CAString;
                cellArg.stringVal = argString;
                return cellArg;
            }
            else
            {
                int row0 = 0;
                int col0 = 0;
                int row1 = 0;
                int col1 = 0;
                if (!ParseCell(rangeString[0], out row0, out col0))
                {
                    //Not simple cell, not sheet string, not valid range - return as string ;
                    cellArg.caType = CellArgumentType.CAString;
                    cellArg.stringVal = argString;
                    return cellArg;
                }
                else
                {
                    if (!ParseCell(rangeString[1], out row1, out col1))
                    {
                        //Not simple cell, not sheet string, not valid range - return as string ;
                        cellArg.caType = CellArgumentType.CAString;
                        cellArg.stringVal = argString;
                        return cellArg;
                    }
                    else
                    {
                        // is valid range
                        cellArg.caType = CellArgumentType.CARange;
                        cellArg.row0 = row0;
                        cellArg.col0 = col0;
                        cellArg.row1 = row1;
                        cellArg.col1 = col1;
                        return cellArg;

                    }
                }
            }
        }
        static CellArgument ParseCellArgument (string argString)
        {
            CellArgument cellArg = new CellArgument();
            float val = 0;
            if (float.TryParse(argString, out val))
            {
                cellArg.caType = CellArgumentType.CAValue;
                cellArg.val = val;
                return cellArg;
            }

            // is this just a cell?
            int row = 0;
            int col = 0;

            if (!ParseCell(argString, out row, out col))
            {
                // not simple cell - is it a Sheet range?
                string[] sheetString = argString.Split(new char[] { '!' }, StringSplitOptions.RemoveEmptyEntries);
                if (sheetString.Length != 2)
                {
                    if (sheetString.Length > 1) // wierd formate string!string!string....
                    {
                        cellArg.caType = CellArgumentType.CAString;
                        cellArg.stringVal = argString;
                        return cellArg;
                    }
                    // format of sheetString[0] not sheet!something must be LL##:LL##
                    cellArg = ParseRange(sheetString[0], argString);
                    return cellArg;
                }
                else
                {
                    // This is a sheet cell - is it sheet!LL## ?
                    if (ParseCell(sheetString[1], out row, out col))
                    {
                        cellArg.caType = CellArgumentType.CASheetCell;
                        cellArg.sheetName = sheetString[0];
                        cellArg.row0 = row;
                        cellArg.col0 = col;
                        return cellArg;
                    }

                    //Not simple cell, IS sheet string - must be range!
                    cellArg = ParseRange(sheetString[1], argString);
                    if (cellArg.caType == CellArgumentType.CARange)
                    {
                        // valid range, but this is with at sheet name - change accordingly
                        cellArg.caType = CellArgumentType.CASheetRange;
                        cellArg.sheetName = sheetString[0];
                        return cellArg;
                    }
                    else
                        return cellArg;
                }
            }
            else
            {
                cellArg.caType = CellArgumentType.CACell;
                cellArg.row0 = row;
                cellArg.col0 = col;
                return cellArg;
            }
        }
        static public bool ParseCell(string cellString, out int row, out int column)
        {
            row = 0;
            column = 0;
            cellString = cellString.Replace("$", "");
            cellString = cellString.ToLower();

            string[] colLetters = new string[2];
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < cellString.Length; i++)
            {
                if ((cellString[i] < 'a') || (cellString[i] > 'z'))
                {
                    if ((cellString[i] < '0') || (cellString[i] > '9'))
                        return false;
                    else
                    {
                        colLetters[1] = cellString.Substring(i);
                        break;
                    }
                }
                else
                    colLetters[0] = sb.Append(cellString[i]).ToString();
            }

            string colString = colLetters[0];
            string rowString = colLetters[1];

            if (!int.TryParse(rowString, out row))
            {
                return false;
            }
            else
            {
                row = row - 1;
            }

            int colVal = 0;
            int colScale = 1;
            for (int i = (colString.Length - 1); i >= 0; i--)
            {
                int idx = (int)(colString.ToLower()[i]) - (int)'a' + 1;
                colVal += (idx * colScale);
                colScale *= 26;
            }
            column = colVal - 1;

            return true;
        }

        #region excel_formulas

        public FormulaReturnValue MAX(string arg)
        {
            CellArgument cellArg = ParseCellArgument(arg);
            FormulaReturnValue retValue = new FormulaReturnValue();
            retValue.returnType = FormulaReturnType.floatFormula;
            switch (cellArg.caType)
            {
                case CellArgumentType.NA:
                    break;
                case CellArgumentType.CAString:
                    {
                        Console.WriteLine("Invalid Cell Type " + cellArg.caType.ToString() + " \"" + arg + "\"");
                        retValue.floatValue = 0;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CAValue:
                    retValue.floatValue = cellArg.val;
                    return retValue;
                case CellArgumentType.CACell:
                    {
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        string sheetName = m_Sheet.SheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                        bool valid = false;
                        float val = FetchNumericValueFromCell(cell, out valid);
                        if (!valid)
                        {
                            Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                            retValue.floatValue = 0;
                            return retValue;
                        }
                        retValue.floatValue = val;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CARange:
                    {
                        float maxVal = 0;
                        bool first = true;
                        string sheetName = m_Sheet.SheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        for (int row = cellArg.row0; row <= cellArg.row1; row++)
                        {
                            for (int col = cellArg.col0; col <= cellArg.col1; col++)
                            {
                                ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                                if (cell != null)
                                {
                                    bool valid = false;
                                    float val = FetchNumericValueFromCell(cell, out valid);
                                    if (!valid)
                                    {
                                        Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                                    }
                                    else
                                    {
                                        if (first)
                                        {
                                            first = true;
                                            maxVal = val;
                                        }
                                        else
                                        {
                                            if (val > maxVal)
                                                maxVal = val;
                                        }
                                    }
                                }
                            }
                        }
                        retValue.floatValue = maxVal;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CASheetCell:
                    {
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        string sheetName = cellArg.sheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                        bool valid = false;
                        float val = FetchNumericValueFromCell(cell, out valid);
                        if (!valid)
                        {
                            Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                            retValue.floatValue = 0;
                            return retValue;
                        }
                        retValue.floatValue = val;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CASheetRange:
                    {
                        float maxVal = 0;
                        bool first = true;
                        string sheetName = cellArg.sheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        for (int row = cellArg.row0; row <= cellArg.row1; row++)
                        {
                            for (int col = cellArg.col0; col <= cellArg.col1; col++)
                            {
                                ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                                if (cell != null)
                                {
                                    bool valid = false;
                                    float val = FetchNumericValueFromCell(cell, out valid);
                                    if (!valid)
                                    {
                                        Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                                    }
                                    else
                                    {
                                        if (first)
                                        {
                                            first = true;
                                            maxVal = val;
                                        }
                                        else
                                        {
                                            if (val > maxVal)
                                                maxVal = val;
                                        }
                                    }
                                }
                            }
                        }
                        retValue.floatValue = maxVal;
                        return retValue;

                    }
                    break;
                default:
                    break;
            }
            retValue.floatValue = 0;
            return retValue;
        }

        public FormulaReturnValue SUM(string arg)
        {
            CellArgument cellArg = ParseCellArgument(arg);
            FormulaReturnValue retValue = new FormulaReturnValue();
            retValue.returnType = FormulaReturnType.floatFormula;

            switch (cellArg.caType)
            {
                case CellArgumentType.NA:
                    break;
                case CellArgumentType.CAString:
                    {
                        Console.WriteLine("Invalid Cell Type " + cellArg.caType.ToString() + " \"" + arg + "\"");
                        retValue.floatValue = 0;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CAValue:
                    retValue.floatValue = cellArg.val;
                    return retValue;
                    break;
                case CellArgumentType.CACell:
                    {
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        string sheetName = m_Sheet.SheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                        bool valid = false;
                        float val = FetchNumericValueFromCell(cell, out valid);
                        if (!valid)
                        {
                            Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                            retValue.floatValue = 0;
                            return retValue;
                        }
                        retValue.floatValue = val;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CARange:
                    {
                        float sum = 0;
                        string sheetName = m_Sheet.SheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        for (int row = cellArg.row0; row <= cellArg.row1; row++)
                        {
                            for (int col = cellArg.col0; col <= cellArg.col1; col++)
                            {
                                ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                                if (cell != null)
                                {
                                    bool valid = false;
                                    float val = FetchNumericValueFromCell(cell, out valid);
                                    if (!valid)
                                    {
                                        Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                                    }
                                    else
                                    {
                                        sum += val;
                                    }
                                }
                            }
                        }
                        retValue.floatValue = sum;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CASheetCell:
                    {
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        string sheetName = cellArg.sheetName;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                        bool valid = false;
                        float val = FetchNumericValueFromCell(cell, out valid);
                        if (!valid)
                        {
                            Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                            retValue.floatValue = 0;
                            return retValue;
                        }
                        retValue.floatValue = val;
                        return retValue;
                    }
                    break;
                case CellArgumentType.CASheetRange:
                    {
                        string sheetName = cellArg.sheetName;
                        float sum = 0;
                        ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
                        for (int row = cellArg.row0; row <= cellArg.row1; row++)
                        {
                            for (int col = cellArg.col0; col <= cellArg.col1; col++)
                            {
                                ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
                                if (cell != null)
                                {
                                    bool valid = false;
                                    float val = FetchNumericValueFromCell(cell, out valid);
                                    if (!valid)
                                    {
                                        Console.WriteLine($"Could not numeric value cell {row},{col} from sheet \"{sheetName}\"");
                                    }
                                    else
                                    {
                                        sum += val;
                                    }
                                }
                            }
                        }
                        retValue.floatValue = sum;
                        return retValue;
                    }
                    break;
                default:
                    break;
            }
            retValue.floatValue = 0;
            return retValue;
        }

        public FormulaReturnValue VLOOKUP(string arg0, string arg1, string arg2, string arg3)
        {
            FormulaReturnValue lookupVal = new FormulaReturnValue();
            CellArgument cellArg = ParseCellArgument(arg0);
            switch (cellArg.caType)
            {
                case CellArgumentType.NA:
                    break;
                case CellArgumentType.CAString:
                    lookupVal.returnType = FormulaReturnType.stringFormula;
                    lookupVal.stringValue = cellArg.stringVal;
                    break;
                case CellArgumentType.CAValue:
                    lookupVal.returnType = FormulaReturnType.stringFormula;
                    lookupVal.stringValue = cellArg.stringVal;
                    break;
                case CellArgumentType.CACell:
                    {
                        string sheetName = m_Sheet.SheetName;
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        ICell cell = FetchCellFromSheet(sheetName, row, col);
                        if (cell != null)
                        {
                            if (cell.CellType != CellType.Numeric)
                            {
                                if (cell.CellType == CellType.Formula)
                                {
                                    lookupVal = EvaluateCellFormula(m_WorkBook, cell);
                                }
                                else
                                {
                                    lookupVal.returnType = FormulaReturnType.stringFormula;
                                    lookupVal.stringValue = CellValueAsString(cell);
                                }
                            }
                            else
                            {
                                lookupVal.returnType = FormulaReturnType.floatFormula;
                                lookupVal.floatValue = (float)(cell.NumericCellValue);
                            }
                        }
                    }
                    break;
                case CellArgumentType.CARange:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CASheetCell:
                    {
                        string sheetName = cellArg.sheetName;
                        int row = cellArg.row0;
                        int col = cellArg.col0;
                        ICell cell = FetchCellFromSheet(sheetName, row, col);
                        if (cell != null)
                        {
                            if (cell.CellType != CellType.Numeric)
                            {
                                lookupVal.returnType = FormulaReturnType.stringFormula;
                                lookupVal.stringValue = CellValueAsString(cell);
                            }
                            else
                            {
                                lookupVal.returnType = FormulaReturnType.floatFormula;
                                lookupVal.floatValue = (float)(cell.NumericCellValue);
                            }
                        }
                    }
                    break;
                case CellArgumentType.CASheetRange:
                    Console.WriteLine("Invalid argument");
                    return null;
                default:
                    break;
            }

            CellArgument cellRangeArg = ParseCellArgument(arg1);
            switch (cellRangeArg.caType)
            {
                case CellArgumentType.NA:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CAString:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CAValue:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CACell:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CARange:
                    break;
                case CellArgumentType.CASheetCell:
                    Console.WriteLine("Invalid argument");
                    return null;
                case CellArgumentType.CASheetRange:
                    break;
                default:
                    Console.WriteLine("Invalid argument");
                    return null;
            }

            int valueCol = 0;
            if(!int.TryParse(arg2, out valueCol))
            {
                Console.WriteLine("Invalid argument");
                return null;
            }
            valueCol = cellRangeArg.col0 + valueCol - 1;

            string rangeSheetName = m_Sheet.SheetName;
            if (cellRangeArg.caType == CellArgumentType.CASheetRange)
                rangeSheetName = cellRangeArg.sheetName;

            Console.WriteLine("Looking for \"" + (lookupVal.returnType == FormulaReturnType.stringFormula ? lookupVal.stringValue : lookupVal.floatValue.ToString("R")) + "\"");
            for (int i = cellRangeArg.row0; i <= cellRangeArg.row1; i++)
            {
                ICell cell = FetchCellFromSheet(rangeSheetName, i, cellRangeArg.col0);
                if (lookupVal.returnType == FormulaReturnType.stringFormula)
                {
                    if (cell.CellType == CellType.Numeric)
                    {
                        if (Math.Abs(cell.NumericCellValue - (double)lookupVal.floatValue) < 0.000001f)
                        {
                            Console.WriteLine("FOUND VALUE row " + i.ToString() + " cell \"" + (cell == null ? "null" : CellValueAsString(cell)) + "\"");
                            ICell returnCell = FetchCellFromSheet(rangeSheetName, i, valueCol);
                            if (returnCell == null)
                                return null;
                            FormulaReturnValue retValue = new FormulaReturnValue();
                            if (returnCell.CellType == CellType.Numeric)
                            {
                                retValue.returnType = FormulaReturnType.floatFormula;
                                retValue.floatValue = (float)(returnCell.NumericCellValue);
                                return retValue;
                            }
                            else
                            {
                                retValue.returnType = FormulaReturnType.stringFormula;
                                retValue.stringValue = CellValueAsString(returnCell);
                                return retValue;
                            }
                        }
                    }
                    else
                    {
                        if (CellValueAsString(cell) == lookupVal.stringValue)
                        {
                            Console.WriteLine("FOUND STRING row " + i.ToString() + " cell \"" + (cell == null ? "null" : CellValueAsString(cell)) + "\"");
                            ICell returnCell = FetchCellFromSheet(rangeSheetName, i, valueCol);
                            if (returnCell == null)
                                return null;
                            FormulaReturnValue retValue = new FormulaReturnValue();
                            if (returnCell.CellType == CellType.Numeric)
                            {
                                retValue.returnType = FormulaReturnType.floatFormula;
                                retValue.floatValue = (float)(returnCell.NumericCellValue);
                                return retValue;
                            }
                            else
                            {
                                retValue.returnType = FormulaReturnType.stringFormula;
                                retValue.stringValue = CellValueAsString(returnCell);
                                return retValue;
                            }
                        }
                    }
                }
                Console.WriteLine("searching row " + i.ToString() + " cell \"" + (cell == null ? "null" : CellValueAsString(cell)) + "\"");
            }

            return null;
        }

        #endregion

        public static FormulaReturnValue EvaluateCellFormula(XSSFWorkbook workbook, ICell cell)
        {
            if (cell.CellType != CellType.Formula)
                return null ;
            ISheet sheet = cell.Sheet;

            string formula = "=" + cell.CellFormula;

            ExcelFormula excelFormula = new ExcelFormula(formula);
            List<ExcelFormulaToken> tokens = new List<ExcelFormulaToken>();
            foreach (ExcelFormulaToken token in excelFormula)
            {
                Console.WriteLine("Token type \"" + token.Type.ToString() + "\" value \"" + token.Value + "\"");
                tokens.Add(token);
            }
            ExcelFormulaEvaluator formulaEvaluator = new ExcelFormulaEvaluator(workbook);
            FormulaReturnValue retValue = formulaEvaluator.EvaluateFormulaFromTokens(sheet, tokens, string.Empty);

            return retValue;
        }

        FormulaReturnValue EvaluateFormula(XSSFWorkbook workbook, ISheet sheet, ExcelFormula excelFormula)
        {
            m_WorkBook = workbook;
            m_Sheet = sheet;

            List<ExcelFormulaToken> tokens = new List<ExcelFormulaToken>();
            foreach (ExcelFormulaToken token in excelFormula)
            {
                tokens.Add(token);
            }
            return EvaluateFormulaFromTokens(sheet, tokens, string.Empty);
        }

        FormulaReturnValue MakeFormulaCall(string formulaName, List<ExcelFormulaToken> tokens)
        {
            // Due to recursion, this formula should not contain other formulas.
            // Find any arguments that are NOT Operand or Argument

            List<ExcelFormulaToken> resolvedTokens = new List<ExcelFormulaToken>();
            // TODO - resolve not Operand or Argument tokens
            for (int i = 0; i < tokens.Count; i++)
            {
                resolvedTokens.Add(tokens[i]);
            }

            // Now build the argument list and invoke the method
            Type thisType = this.GetType();
            MethodInfo theMethod = thisType.GetMethod(formulaName);
            if (theMethod == null)
            {
                // is this just a cell lookup?
                Console.WriteLine("UNKNOWN FORMULA \"" + formulaName + "\"");
                return null;
            }
            else
            {
                List<object> formulaArguments = new List<object>();
                for (int i = 0; i < resolvedTokens.Count; i++)
                {
                    if (resolvedTokens[i].Type == ExcelFormulaTokenType.Operand)
                    {
                        formulaArguments.Add(resolvedTokens[i].Value);
                    }
                }

                FormulaReturnValue retValue = (FormulaReturnValue)theMethod.Invoke(this, formulaArguments.ToArray());
                return retValue;
            }
        }

        public FormulaReturnValue EvaluateFormulaFromTokens(ISheet sheet, List<ExcelFormulaToken>tokens, string formulaName)
        {
            m_Sheet = sheet;
            List<ExcelFormulaToken> tokensToEvaluate = new List<ExcelFormulaToken>();
            List<ExcelFormulaToken> tokensInFormula = new List<ExcelFormulaToken>();
            bool insideFormula = false;
            for (int i = 0; i < tokens.Count; i++)
            {
                ExcelFormulaToken token = tokens[i];
                switch (token.Type)
                {
                    case ExcelFormulaTokenType.Noop:
                        break;
                    case ExcelFormulaTokenType.Operand:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.Function:
                        {
                            if (token.Value == string.Empty)
                            {
                                if (!insideFormula)
                                    Console.WriteLine("Got end of formula outside of formula");
                                else
                                {
                                    //float val = EvaluateFormulaFromTokens(workbook, sheet, tokensToEvaluate, formulaName);
                                    //ExcelFormulaToken fToken = new ExcelFormulaToken(val.ToString("R"), ExcelFormulaTokenType.Operand);
                                    //tokensToEvaluate.Add(fToken);
                                    FormulaReturnValue retValue = MakeFormulaCall(formulaName, tokensInFormula);
                                    if (retValue != null)
                                    {
                                        ExcelFormulaToken fToken = new ExcelFormulaToken(string.Empty, ExcelFormulaTokenType.Noop);
                                        if (retValue.returnType == FormulaReturnType.floatFormula)
                                            fToken = new ExcelFormulaToken(retValue.floatValue.ToString("R"), ExcelFormulaTokenType.Operand);
                                        else
                                            fToken = new ExcelFormulaToken(retValue.stringValue, ExcelFormulaTokenType.Operand);
                                        tokensToEvaluate.Add(fToken);
                                    }
                                    insideFormula = false;
                                }
                            }
                            else
                            {
                                formulaName = token.Value;
                                insideFormula = true;
                            }
                        }
                        break;
                    case ExcelFormulaTokenType.Subexpression:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.Argument:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.OperatorPrefix:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.OperatorInfix:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.OperatorPostfix:
                        if (insideFormula)
                            tokensInFormula.Add(token);
                        else
                            tokensToEvaluate.Add(token);
                        break;
                    case ExcelFormulaTokenType.Whitespace:
                        break;
                    case ExcelFormulaTokenType.Unknown:
                        break;
                    default:
                        break;
                }
                //if (methodName != string.Empty)
                //{
                //    Type thisType = this.GetType();
                //    MethodInfo theMethod = thisType.GetMethod(methodName);

                //    theMethod.Invoke(this, excelFormula);
                //}
                //if (token.Type == ExcelFormulaTokenType.Function)
                //{
                //    methodName = token.Value;
                //}
                //Type thisType = this.GetType();
                //MethodInfo theMethod = thisType.GetMethod(token.Value);
                //theMethod.Invoke(this, excelFormula);

                Console.WriteLine("Token type \"" + token.Type.ToString() + "\" value \"" + token.Value + "\"");
            }

            bool formulaIsNumeric = true;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < tokensToEvaluate.Count; i++)
            {
                ExcelFormulaToken token = tokensToEvaluate[i];
                switch (token.Type)
                {
                    case ExcelFormulaTokenType.Noop:
                        break;
                    case ExcelFormulaTokenType.Operand:
                        {
                            float val = 0;
                            if (float.TryParse(token.Value, out val))
                            {
                                sb.Append(token.Value);
                            }
                            else
                            {
                                CellArgument cellArg = ParseCellArgument(token.Value);
                                switch (cellArg.caType)
                                {
                                    case CellArgumentType.NA:
                                        break;
                                    case CellArgumentType.CAString:
                                        formulaIsNumeric = false;
                                        sb.Append(token.Value);
                                        break;
                                    case CellArgumentType.CAValue:
                                        sb.Append(token.Value);
                                        break;
                                    case CellArgumentType.CACell:
                                        {
                                            string sheetName = m_Sheet.SheetName;
                                            int row = cellArg.row0;
                                            int col = cellArg.col0;
                                            ICell cell = FetchCellFromSheet(sheetName, row, col);
                                            if (cell != null)
                                            {
                                                if (cell.CellType != CellType.Numeric)
                                                {
                                                    formulaIsNumeric = false;
                                                    sb.Append(CellValueAsString(cell));
                                                }
                                                else
                                                    sb.Append(cell.NumericCellValue.ToString("R"));
                                            }
                                        }
                                        break;
                                    case CellArgumentType.CARange:
                                        sb.Append(token.Value);
                                        break;
                                    case CellArgumentType.CASheetCell:
                                        {
                                            string sheetName = cellArg.sheetName;
                                            int row = cellArg.row0;
                                            int col = cellArg.col0;
                                            ICell cell = FetchCellFromSheet(sheetName, row, col);
                                            if (cell != null)
                                            {
                                                if (cell.CellType != CellType.Numeric)
                                                {
                                                    formulaIsNumeric = false;
                                                    sb.Append(CellValueAsString(cell));
                                                }
                                                else
                                                    sb.Append(cell.NumericCellValue.ToString("R"));
                                            }
                                        }
                                        break;
                                    case CellArgumentType.CASheetRange:
                                        sb.Append(token.Value);
                                        break;
                                    default:
                                        break;
                                }

                            }
                            if (i < (tokensToEvaluate.Count - 1))
                                sb.Append(" ");
                        }
                        break;
                    case ExcelFormulaTokenType.Function:
                        break;
                    case ExcelFormulaTokenType.Subexpression:
                        break;
                    case ExcelFormulaTokenType.Argument:
                        break;
                    case ExcelFormulaTokenType.OperatorPrefix:
                        sb.Append(token.Value);
                        break;
                    case ExcelFormulaTokenType.OperatorInfix:
                        sb.Append(token.Value);
                        if (i < (tokensToEvaluate.Count - 1))
                            sb.Append(" ");
                        break;
                    case ExcelFormulaTokenType.OperatorPostfix:
                        sb.Append(token.Value);
                        if (i < (tokensToEvaluate.Count - 1))
                            sb.Append(" ");
                        break;
                    case ExcelFormulaTokenType.Whitespace:
                        break;
                    case ExcelFormulaTokenType.Unknown:
                        break;
                    default:
                        break;
                }
            }

            FormulaReturnValue returnValue = new FormulaReturnValue();
            if (formulaIsNumeric)
            {
                string formulaString = sb.ToString();
                Console.WriteLine("formulaString \"" + formulaString + "\"");
                ShuntingYardSimpleMath SY = new ShuntingYardSimpleMath();
                List<String> ss = formulaString.Split(' ').ToList();
                for (int i = 0; i < ss.Count; i++)
                {
                    Console.WriteLine("ss " + i.ToString() + ": \"" + ss[i] + "\"");
                }

                Double res = SY.Execute(ss, null);
                Console.WriteLine("SY = " + res.ToString("R"));
                returnValue.returnType = FormulaReturnType.floatFormula;
                returnValue.floatValue = (float)res;
            }
            else
            {
                returnValue.returnType = FormulaReturnType.stringFormula;
                returnValue.stringValue = sb.ToString();
            }
            return returnValue;
        }

    }
}
