using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Formula.Functions;

using ExcelFormulaParser;

namespace XLSTest
{
    class Program
    {
        static string BUDGET_EXAMPLE_SHEET_NAME = "Budget Example";
        static public XSSFWorkbook m_WorkBook = null;
        //static ISheet[] m_Sheets;

        static void Usage(string[] args)
        {
            Console.WriteLine($"Usage: {args[0]} file.xls Sheet Cell");
        }

        //static bool ParseCell(string cellString, out int row, out int column)
        //{
        //    row = 0;
        //    column = 0;
        //    cellString = cellString.Replace("$", "");
        //    string[] colLetters = cellString.Split(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
        //    if (colLetters.Length != 2)
        //        return false;
        //    string colString = colLetters[0].ToLower();
        //    for (int i = 0; i < colString.Length; i++)
        //    {
        //        if ((colString[i] < 'a') || (colString[i] > 'z'))
        //            return false;
        //    }
        //    string rowString = cellString.Substring(colString.Length);

        //    if (!int.TryParse(rowString, out row))
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        row = row - 1;
        //    }

        //    int colVal = 0;
        //    int colScale = 1;
        //    for (int i = (colString.Length-1); i >= 0 ; i--)
        //    {
        //        int idx = (int)(colString.ToLower()[i]) - (int)'a' + 1;
        //        colVal += (idx * colScale);
        //        colScale *= 26;
        //    }
        //    column = colVal - 1;

        //    return true;
        //}

        static string CellToString(ICell cell)
        {
            //return cell.ToString();

            string arrayFormulaRangeStr = string.Empty;
            if (cell.IsPartOfArrayFormulaGroup)
                arrayFormulaRangeStr = cell.ArrayFormulaRange.FormatAsString();
            string hyperlinkAddressStr = string.Empty;
            if (cell.Hyperlink != null)
                hyperlinkAddressStr = cell.Hyperlink.Address;
            string commentStr = string.Empty;
            if (cell.CellComment != null)
                commentStr = cell.CellComment.ToString();
            string cellValue = string.Empty;
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    cellValue = "Unknown";
                    break;
                case CellType.Numeric:
                    cellValue = "N " + cell.NumericCellValue.ToString("R");
                    break;
                case CellType.String:
                    cellValue = "S " + cell.StringCellValue;
                    break;
                case CellType.Formula:
                    cellValue = "F " + cell.CellFormula;
                    break;
                case CellType.Blank:
                    cellValue = "Blank";
                    break;
                case CellType.Boolean:
                    cellValue = "B " + cell.BooleanCellValue.ToString();
                    break;
                case CellType.Error:
                    cellValue = "Error " + cell.ErrorCellValue.ToString();
                    break;
                default:
                    break;
            }

            string ret = "afr \"" + arrayFormulaRangeStr + "\"\n";
            ret += "link \"" + hyperlinkAddressStr + "\"\n";
            ret += "comment \"" + commentStr + "\"\n";
            ret += "value \"" + cellValue + "\"";
            return ret;

            /*
            //ICellStyle CellStyle { get; set; }
            bool BooleanCellValue { get; }
            string StringCellValue { get; }
            byte ErrorCellValue { get; }
            IRichTextString RichStringCellValue { get; }
            DateTime DateCellValue { get; }
            double NumericCellValue { get; }
            string CellFormula { get; set; }
            CellType CachedFormulaResultType { get; }
            CellType CellType { get; }
            IRow Row { get; }
            bool IsMergedCell { get; }
            int RowIndex { get; }
            int ColumnIndex { get; }
            bool IsPartOfArrayFormulaGroup { get; }
            */
        }

        static float EvaluateFormula(ICell cell)
        {
            if (cell.CellType != CellType.Formula)
                return 0;
            ISheet sheet = cell.Sheet;

            string formula = "=" + cell.CellFormula;

            ExcelFormula excelFormula = new ExcelFormula(formula);
            List<ExcelFormulaToken> tokens = new List<ExcelFormulaToken>();
            foreach (ExcelFormulaToken token in excelFormula)
            {
                Console.WriteLine("Token type \"" + token.Type.ToString() + "\" value \"" + token.Value + "\"");
                tokens.Add(token);
            }
            ExcelFormulaEvaluator formulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);
            FormulaReturnValue retValue = formulaEvaluator.EvaluateFormulaFromTokens(sheet, tokens, string.Empty);

            return 0;
        }

        static void Main(string[] args)
        {
            if (args.Length < 4)
            {
                Usage(args);
                return;
            }
            string uri = args[1];
            string sheetName = args[2];
            string cellString = args[3];
            int row = 0;
            int col = 0;

            if(!ExcelFormulaEvaluator.ParseCell(cellString, out row, out col))
            {
                Usage(args);
                return;
            }

            m_WorkBook = new XSSFWorkbook(uri);
            for (int i = 0; i < m_WorkBook.NumberOfSheets; i++)
            {
                Console.WriteLine("Workbook " + i.ToString() + ": \"" + m_WorkBook.GetSheetName(i) + "\"");
            }

            //ISheet sheet = m_WorkBook.GetSheet(sheetName);
            //if (sheet == null)
            //{
            //    Console.WriteLine("Could not get sheet \"" + sheetName + "\"");
            //    return;
            //}

            //IRow cellRow = sheet.GetRow(row);
            //if (cellRow == null)
            //{
            //    Console.WriteLine($"Could not get row {row} from sheet \"{sheetName}\"");
            //    return;
            //}
            //int iStartCellIdx = cellRow.FirstCellNum;
            //ICell cell = cellRow.GetCell(col);
            //if (cell == null)
            //{
            //    Console.WriteLine($"Could not get cell {row},{col} from sheet \"{sheetName}\"");
            //    return;
            //}
            ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);

            ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
            if (cell == null)
            {
                Console.WriteLine($"Could not get cell {row},{col} from sheet \"{sheetName}\"");
                return;
            }

            Console.WriteLine("cell " +row.ToString() + "," + col.ToString() + " \"" + cellString + "\" is \"" + CellToString(cell) + "\"");
            if (cell.CellType == CellType.Formula)
            {
                FormulaReturnValue cellValue = ExcelFormulaEvaluator.EvaluateCellFormula(m_WorkBook, cell);
                if (cellValue == null)
                {
                    Console.WriteLine("Cell value NULL");
                }
                else
                {
                    if (cellValue.returnType == FormulaReturnType.stringFormula)
                        Console.WriteLine("cell formula = \"" + cellValue.stringValue + "\"");
                    else
                        Console.WriteLine("cell formula = " + cellValue.floatValue.ToString("R"));
                }
            }
            //m_Sheet = m_WorkBook.GetSheet(BUDGET_EXAMPLE_SHEET_NAME);
            //if (m_Sheet == null)
            //{
            //    Console.WriteLine($"didn't get \"{BUDGET_EXAMPLE_SHEET_NAME}\"");
            //}
        }
    }
}
