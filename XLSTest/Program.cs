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
        static public XSSFWorkbook m_WorkBook = null;

        static void Usage(string[] args)
        {
            Console.WriteLine($"Usage: {args[0]} file.xls Sheet Cell");
        }

        static string CellToString(ICell cell)
        {
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

            ExcelFormulaEvaluator excelFormulaEvaluator = new ExcelFormulaEvaluator(m_WorkBook);

            ICell cell = excelFormulaEvaluator.FetchCellFromSheet(sheetName, row, col);
            if (cell == null)
            {
                Console.WriteLine($"Could not get cell {row},{col} from sheet \"{sheetName}\"");
                return;
            }

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
                        Console.WriteLine("cell = \"" + cellValue.stringValue + "\"");
                    else
                        Console.WriteLine("cell = " + cellValue.floatValue.ToString("R"));
                }
            }
            else
                Console.WriteLine("cell = " + ExcelFormulaEvaluator.CellValueAsString(cell)) ;

        }
    }
}
