using CalConverter.Lib.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalConverter.Lib;
public static class Utils
{
    // A function that takes a string that represents an Excel range like C2:E4 or AA644:AA645
    // and returns a list of strings with all the cells in the range
    public static List<string> GetCells(string range)
    {
        // Initialize an empty list to store the cells
        List<string> cells = new List<string>();

        // Split the range by the colon
        string[] parts = range.Split(':');

        // If the range is invalid, return an empty list
        if (parts.Length != 2)
        {
            return cells;
        }

        // Get the start and end column and row from the range
        var startSplit = SplitRangeRef(parts[0]);
        var endSplit = SplitRangeRef(parts[1]);
        string startCol = startSplit.column;
        string endCol = endSplit.column;
        int startRow = int.Parse(startSplit.row);
        int endRow = int.Parse(endSplit.row);

        // Convert the columns to numbers
        int startColNum = ColumnToNumber(startCol);
        int endColNum = ColumnToNumber(endCol);

        // Loop through the columns and rows in the range
        for (int colNum = startColNum; colNum <= endColNum; colNum++)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                // Convert the column number to a letter
                string col = NumberToColumn(colNum);

                // Concatenate the column and row to form a cell
                string cell = col + row.ToString();

                // Add the cell to the list
                cells.Add(cell);
            }
        }

        // Return the list of cells
        return cells;
    }

    // A helper function that converts a column letter to a number
    public static int ColumnToNumber(string col)
    {
        // Initialize the result to zero
        int result = 0;

        // Loop through each character in the column
        for (int i = 0; i < col.Length; i++)
        {
            // Get the ASCII value of the character
            int value = (int)col[i];

            // Check if the value is valid
            if (value < 65 || value > 90)
            {
                // Return zero if invalid
                return 0;
            }

            // Update the result by multiplying by 26 and adding the value
            result = result * 26 + (value - 64);
        }

        // Return the result
        return result;
    }

    // A helper function that converts a column number to a letter
    public static string NumberToColumn(int num)
    {
        // Initialize an empty string to store the result
        string result = "";

        // Loop until the number is zero
        while (num > 0)
        {
            // Get the remainder of the number divided by 26
            int rem = num % 26;

            // If the remainder is zero, set it to 26 and subtract one from the number
            if (rem == 0)
            {
                rem = 26;
                num--;
            }

            // Convert the remainder to a character and prepend it to the result
            char ch = (char)(rem + 64);
            result = ch + result;

            // Divide the number by 26
            num = num / 26;
        }

        // Return the result
        return result;

    }

    public static (string column, string row) SplitRangeRef(string value)
    {
        // Find the index of the first digit in the value
        int index = value.IndexOfAny("0123456789".ToCharArray());

        // Split the value into letters and numbers
        string letters = value.Substring(0, index);
        string numbers = value.Substring(index);

        // Create a tuple with the letters and numbers
        return (letters, numbers);
    }

    public static string IncrementRow(string value)
    {
        // Find the index of the first digit in the value
        int index = value.IndexOfAny("0123456789".ToCharArray());

        // Split the value into letters and numbers
        string column = value.Substring(0, index);
        int row = int.Parse(value.Substring(index));

        // Create a tuple with the letters and numbers
        return column + (row + 1).ToString();
    }

    public static Cell? GetRelativeCell(this SheetData sheetData, SimpleCellData cell, int rowOffset = 0, int colOffset = 0)
    {
        string value = cell.CellRef;

        // Find the index of the first digit in the value
        int index = value.IndexOfAny("0123456789".ToCharArray());

        // Split the value into letters and numbers
        int column = ColumnToNumber(value.Substring(0, index));
        int row = int.Parse(value.Substring(index));

        // Create a tuple with the letters and numbers
        string nextCell = NumberToColumn(column + colOffset) + (row + rowOffset).ToString();

        return sheetData.Descendants<Cell>().FirstOrDefault(q => q.CellReference == nextCell);
    }

    public static string GetColorName(string hexColor)
    {
        if(hexColor == null || hexColor == "NULL")
        {
            return "";
        }

        try
        {
            // Parse the hex color string to a Color object
            System.Drawing.Color color = ColorTranslator.FromHtml("#" + hexColor);
            if (color.IsNamedColor)
            {
                return color.Name;
            }

            string r = color.R > 200 ? "FF" : "00";
            string g = color.G > 200 ? "FF" : "00";
            string b = color.B > 200 ? "FF" : "00";
            string roundedColor = r + g + b;
            switch (roundedColor)
            {
                case "000000": return "Black";
                case "0000FF": return "Blue";
                case "00FF00": return "Green";
                case "FF0000": return "Red";
                case "00FFFF": return "Cyan";
                case "FFFF00": return "Yellow";
                case "FFFFFF": return "White";
                default: return "Unknown";
            }
        }
        catch (Exception) {

            return "Unknown";
        }
    }

}
