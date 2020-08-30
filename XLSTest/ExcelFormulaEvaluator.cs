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

namespace XLSTest
{
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

            return null;
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



        public float MAX(string arg)
        {
            CellArgument cellArg = ParseCellArgument(arg);
            switch (cellArg.caType)
            {
                case CellArgumentType.NA:
                    break;
                case CellArgumentType.CAString:
                    {
                        Console.WriteLine("Invalid Cell Type " + cellArg.caType.ToString() + " \"" + arg + "\"");
                        return 0;
                    }
                    break;
                case CellArgumentType.CAValue:
                    return cellArg.val;
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
                            return 0;
                        }
                        return val;
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
                        return maxVal;
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
                            return 0;
                        }
                        return val;
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
                        return maxVal;
                    }
                    break;
                default:
                    break;
            }
            return 0;
        }

        float EvaluateFormula(XSSFWorkbook workbook, ISheet sheet, ExcelFormula excelFormula)
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

        float MakeFormulaCall(string formulaName, List<ExcelFormulaToken> tokens)
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
            List<object> formulaArguments = new List<object>();
            for (int i = 0; i < resolvedTokens.Count; i++)
            {
                if (resolvedTokens[i].Type == ExcelFormulaTokenType.Operand)
                {
                    formulaArguments.Add(resolvedTokens[i].Value);
                }
            }

            float val = (float)theMethod.Invoke(this, formulaArguments.ToArray());
            return val ;
        }

        public float EvaluateFormulaFromTokens(ISheet sheet, List<ExcelFormulaToken>tokens, string formulaName)
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
                                    float val = MakeFormulaCall(formulaName, tokensInFormula);
                                    ExcelFormulaToken fToken = new ExcelFormulaToken(val.ToString("R"), ExcelFormulaTokenType.Operand);
                                    tokensToEvaluate.Add(fToken);
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

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < tokensToEvaluate.Count; i++)
            {

            }

            return 0;
        }

    }
}
