using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace TextReportGenerator
{
  public class ExcelHelper : IDisposable
  {
    #region IDisposable Members
    void IDisposable.Dispose()
    {
      this.CloseExcel();
    }

    #endregion

    public void Initial(string excelFile, int? decimalPlaces = null)
    {
      this._excelFileName = DPApplication.GetValidFilePath(excelFile);
      this._decimalPlaces = Math.Max(0, decimalPlaces ?? 3);

      if (this._application != null)
        _application = null;

      _beforeTime = DateTime.Now;
      Type oExcel = Type.GetTypeFromProgID("Excel.Application");
      if (oExcel == null)
        throw new Exception("Office excel has not been installed!");

      _application = Activator.CreateInstance(oExcel);
      _afterTime = DateTime.Now;

      InvokeHelper.SetProperty(_application, "DisplayAlerts", false);
#if DEBUG
      InvokeHelper.SetProperty(_application, "Visible", true);
#else
	    InvokeHelper.SetProperty(_application, "Visible", false);
#endif
      _workbooks = InvokeHelper.GetProperty(_application, "Workbooks");

      if (System.IO.File.Exists(_excelFileName))
        _workbook = InvokeHelper.CallMethod(_workbooks, "Open", this._excelFileName);
      else
        _workbook = InvokeHelper.CallMethod(_workbooks, "Add", true);

      _worksheets = InvokeHelper.GetProperty(_workbook, "Worksheets");
      _worksheet = InvokeHelper.GetProperty(_workbook, "ActiveSheet");
    }
    public void SetVisible(bool isVisible)
    {
      try
      {
        if (_application != null)
          InvokeHelper.SetProperty(_application, "Visible", isVisible);
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCells(int topRow, int leftColumn, object[,] cellValues)
    {
      try
      {
        WriteCells(_worksheet, topRow, leftColumn, cellValues);
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCells(object[,] cellValues)
    {
      try
      {
        WriteCells(_worksheet, 1, 1, cellValues);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCells(int sheetIndex, int topRow, int leftColumn, object[,] cellValues)
    {
      try
      {
        if (sheetIndex > 0)
        {
          object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
          WriteCells(worksheet, topRow, leftColumn, cellValues);
        }
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCells(object worksheet, int topRow, int leftColumn, object[,] cellValues)
    {
      try
      {
        if (worksheet == null || topRow < 1 || leftColumn < 1 || cellValues == null)
          return;

        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", topRow, leftColumn);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", topRow + cellValues.GetLength(0) - 1, leftColumn + cellValues.GetLength(1) - 1);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        InvokeHelper.SetProperty(range, "Value", cellValues);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void SetRangeNumberFormat(Int32 sheetIndex, Int32 topRow, Int32 leftColumn, Int32 rows, Int32 columns)
    {
      try
      {
        this.SetRangeNumberFormat(sheetIndex, topRow, leftColumn, new Size(columns, rows));
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void SetRangeNumberFormat(Int32 sheetIndex, Int32 topRow, Int32 leftColumn, Size formatRange)
    {
      try
      {
        if (sheetIndex <= 0 || topRow < 1 || leftColumn < 1 || formatRange.IsEmpty)
          return;

        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", topRow, leftColumn);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", topRow + formatRange.Height - 1, leftColumn + formatRange.Width - 1);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        InvokeHelper.SetProperty(range, "NumberFormat", NumberFormatString);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCell(object worksheet, int rowIndex, int columnIndex, object cellValue)
    {
      try
      {
        object cell = InvokeHelper.GetProperty(worksheet, "Cells", rowIndex, columnIndex);

        InvokeHelper.SetProperty(cell, "Value", cellValue);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCell(int sheetIndex, int rowIndex, int columnIndex, object cellValue)
    {
      try
      {
        WriteCell(sheetIndex, rowIndex, columnIndex, cellValue, 1);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void WriteCell(int sheetIndex, int rowIndex, int columnIndex, object cellValue, int colorIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        object cell = InvokeHelper.GetProperty(worksheet, "Cells", rowIndex, columnIndex);

        this.SetCellColor(cell, colorIndex);

        if (cellValue is double)
          InvokeHelper.SetProperty(cell, "NumberFormat", NumberFormatString);
        if (cellValue is double && double.IsNaN((double)cellValue))
          InvokeHelper.SetProperty(cell, "Value", null);
        else
          InvokeHelper.SetProperty(cell, "Value", cellValue);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void MergeCells(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, object cellValue, XlHAlign align = XlHAlign.xlHAlignCenter)
    {
      try
      {
        MergeCells(_worksheet, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cellValue, align);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void MergeCells(int sheetIndex, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, object cellValue, XlHAlign align = XlHAlign.xlHAlignCenter)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        MergeCells(worksheet, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cellValue, align);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void MergeCells(object worksheet, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, object cellValue, XlHAlign align = XlHAlign.xlHAlignCenter)
    {
      try
      {
        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", startRowIndex, startColumnIndex);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", endRowIndex, endColumnIndex);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        InvokeHelper.CallMethod(range, "Merge");
        InvokeHelper.SetProperty(range, "HorizontalAlignment", align);

        if (cellValue is double)
          InvokeHelper.SetProperty(range, "NumberFormat", NumberFormatString);
        if (cellValue is double && double.IsNaN((double)cellValue))
          InvokeHelper.SetProperty(range, "Value", null);
        else
          InvokeHelper.SetProperty(range, "Value", cellValue);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void SetCellColor(int sheetIndex, int rowIndex, int columnIndex, int colorIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        object cell = InvokeHelper.GetProperty(worksheet, "Cells", rowIndex, columnIndex);

        this.SetCellColor(cell, colorIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void SetCellColor(object cell, int colorIndex)
    {
      try
      {
        object cellFont = InvokeHelper.GetProperty(cell, "Font");
        InvokeHelper.SetProperty(cellFont, "ColorIndex", colorIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void SetCellBackground(int sheetIndex, int rowIndex, int columnIndex, int colorIndex)
    {
      try
      {
        if (colorIndex == 2) // format.Background != Brushes.White
          return;

        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        object cell = InvokeHelper.GetProperty(worksheet, "Cells", rowIndex, columnIndex);
        object interior = InvokeHelper.GetProperty(cell, "Interior");
        InvokeHelper.SetProperty(interior, "ColorIndex", colorIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }
    public void FormatCells(int sheetIndex, List<ExcelRangeFormat> formats, int rowOffset = 0)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
        if (worksheet == null)
          return;

        foreach (ExcelRangeFormat format in formats)
        {
          if (format == null)
            continue;

          object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", format.RowStartIndex + rowOffset, format.ColumnStartIndex);
          object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", format.RowEndIndex + rowOffset, format.ColumnEndIndex);
          object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);

          // Merge & Set Value
          if (format.ColumnCount > 1 || format.RowCount > 1)
            InvokeHelper.CallMethod(range, "Merge");
          if (format.ValueObject != null && format.ValueObject.ToString() != String.Empty)
            InvokeHelper.SetProperty(range, "Value2", format.ValueObject);

          // Foreground Color & Bold
          object font = InvokeHelper.GetProperty(range, "Font");
          InvokeHelper.SetProperty(font, "ColorIndex", format.ForegroundColorIndex);
          InvokeHelper.SetProperty(font, "Bold", format.Bold);

          // Background Color
          if (format.BackgroundColorIndex != 2) // format.Background != Brushes.White
          {
            object interior = InvokeHelper.GetProperty(range, "Interior");
            InvokeHelper.SetProperty(interior, "ColorIndex", format.BackgroundColorIndex);
          }

          // Number Format
          if (format.NeedNumberFormat)
            InvokeHelper.SetProperty(range, "NumberFormat", NumberFormatString);

          // Alignment
          if (format.HAlign != null)
            InvokeHelper.SetProperty(range, "HorizontalAlignment", format.HAlign);
          if (format.VAlign != null)
            InvokeHelper.SetProperty(range, "VerticalAlignment", format.VAlign);
        }
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }

    public object ReadCell(int sheetIndex, int rowIndex, int columnIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);

        return ReadCell(worksheet, rowIndex, columnIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public T ReadCell<T>(int sheetIndex, int rowIndex, int columnIndex, T defaultValue)
    {
      var cellValue = this.ReadCell(sheetIndex, rowIndex, columnIndex);
      if (cellValue == null)
        return defaultValue;
      var textValue = cellValue.ToString();
      if (!string.IsNullOrWhiteSpace(textValue))
        return textValue.SafeConvertInvariantStringTo<T>();
      return defaultValue;
    }
    public static object ReadCell(object worksheet, int rowIndex, int columnIndex)
    {
      try
      {
        if (worksheet == null)
          return null;

        object cell = InvokeHelper.GetProperty(worksheet, "Cells", rowIndex, columnIndex);

        var isMergeCell = (bool)InvokeHelper.GetProperty(cell, "MergeCells");
        if (isMergeCell)
        {
          var mergeArea = InvokeHelper.GetProperty(cell, "MergeArea");
          var firstCellRow = InvokeHelper.GetProperty(mergeArea, "Row");
          var firstCellColumn = InvokeHelper.GetProperty(mergeArea, "Column");

          cell = InvokeHelper.GetProperty(worksheet, "Cells", firstCellRow, firstCellColumn);
        }

        return InvokeHelper.GetProperty(cell, "Value");
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static T ReadCell<T>(object worksheet, int rowIndex, int columnIndex, T defaultValue)
    {
      try
      {
        var cellValue = ReadCell(worksheet, rowIndex, columnIndex);
        if (cellValue == null)
          return defaultValue;
        var textValue = cellValue.ToString();
        if (!string.IsNullOrWhiteSpace(textValue))
          return textValue.SafeConvertInvariantStringTo<T>();
        return defaultValue;
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.LogFile(ex.Message);
        return defaultValue;
      }
    }
    public static object[,] ReadAllCells(object worksheet)
    {
      try
      {
        if (worksheet == null)
          return null;

        var usedRangeSize = GetUsedRangeSize(worksheet);

        return ReadCells(worksheet, 1, 1, usedRangeSize.Height, usedRangeSize.Width);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static object[,] ReadAllCellFormulas(object worksheet)
    {
      try
      {
        if (worksheet == null)
          return null;

        var usedRangeSize = GetUsedRangeSize(worksheet);

        return ReadCellFormulas(worksheet, 1, 1, usedRangeSize.Height, usedRangeSize.Width);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }

    public object[,] ReadCells(int sheetIndex, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);

        return ReadCells(worksheet, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public object[,] ReadCellFormulas(int sheetIndex, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);

        return ReadCellFormulas(worksheet, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    // return Formula if it is not empty, otherwise return value
    public object[,] ReadCellFormulaAndValues(int sheetIndex, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);

        return ReadCellFormulaAndValues(worksheet, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }

    public static object[,] ReadCells(object worksheet, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        if (worksheet == null)
          return null;

        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", startRowIndex, startColumnIndex);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", endRowIndex, endColumnIndex);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        return InvokeHelper.GetProperty(range, "Value") as object[,];
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static object[,] ReadCellFormulas(object worksheet, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        if (worksheet == null)
          return null;

        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", startRowIndex, startColumnIndex);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", endRowIndex, endColumnIndex);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        return InvokeHelper.GetProperty(range, "Formula") as object[,];
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static object[,] ReadCellFormulaAndValues(object worksheet, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex)
    {
      try
      {
        if (worksheet == null)
          return null;

        object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", startRowIndex, startColumnIndex);
        object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", endRowIndex, endColumnIndex);
        object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
        var result = InvokeHelper.GetProperty(range, "Formula") as object[,];
        var values = InvokeHelper.GetProperty(range, "Value") as object[,];

        if (result != null && values != null &&
            result.GetLowerBound(0) == values.GetLowerBound(0) && result.GetUpperBound(0) == values.GetUpperBound(0) &&
            result.GetLowerBound(1) == values.GetLowerBound(1) && result.GetUpperBound(1) == values.GetUpperBound(1))
        {
          for (var r = result.GetLowerBound(0); r < result.GetUpperBound(0); r++)
          {
            for (var c = result.GetLowerBound(1); c < result.GetUpperBound(1); c++)

              if (result[r, c] == null)
                result[r, c] = values[r, c];
          }
        }
        return result;

      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }

    public static string CellToString(object cellValue)
    {
      if (cellValue == null)
        return "";

      if (cellValue is int)
      {
        var intValue = (int)cellValue;
        if (intValue == (int)XlCellErrorValue.ErrDiv0) // = -2146826281, 
          return "#Div/0!";

        if (intValue == (int)XlCellErrorValue.ErrNA) // = -2146826246,  
          return "#N/A";
        if (intValue == (int)XlCellErrorValue.ErrName) // = -2146826259,  
          return "#Name?";
        if (intValue == (int)XlCellErrorValue.ErrNull) // = -2146826288, 
          return "#Null!";
        if (intValue == (int)XlCellErrorValue.ErrNum) // = -2146826252,  
          return "#Num!";
        if (intValue == (int)XlCellErrorValue.ErrRef) // = -2146826265,  
          return "#Ref!";
        if (intValue == (int)XlCellErrorValue.ErrValue) // = -2146826273  
          return "#Value!";
      }
      return cellValue.ToString();
    }

    public static XlSheetVisibility GetWorksheetVisibility(object worksheet)
    {
      return (XlSheetVisibility)InvokeHelper.GetProperty(worksheet, "Visible");
    }
    public static void SetWorksheetVisibility(object worksheet, XlSheetVisibility visibility)
    {
      try
      {
        if (worksheet == null)
          return;

        InvokeHelper.SetProperty(worksheet, "Visible", visibility);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }

    public static object GetUsedRange(object worksheet)
    {
      return InvokeHelper.GetProperty(worksheet, "UsedRange");
    }
    public static Int32Rect GetUsedRangeArea(object worksheet)
    {
      var usedRange = GetUsedRange(worksheet);

      var startRow = (int)InvokeHelper.GetProperty(usedRange, "Row");
      var startColumn = (int)InvokeHelper.GetProperty(usedRange, "Column");

      var rows = GetRows(usedRange);
      var columns = GetColumns(usedRange);
      var rowCount = (int)GetRowCount(rows);
      var columnCount = (int)GetColumnCount(columns);

      return new Int32Rect(startColumn, startRow, columnCount, rowCount);
    }
    public static Int32Size GetUsedRangeSize(object worksheet)
    {
      var usedRangeArea = GetUsedRangeArea(worksheet);

      return new Int32Size(usedRangeArea.X + usedRangeArea.Width - 1, usedRangeArea.Y + usedRangeArea.Height - 1);
    }

    public static object GetRows(object range)
    {
      return InvokeHelper.GetProperty(range, "Rows");
    }
    public static object GetColumns(object range)
    {
      return InvokeHelper.GetProperty(range, "Columns");
    }
    public static int GetRowCount(object rows)
    {
      return (int)InvokeHelper.GetProperty(rows, "Count");
    }
    public static int GetColumnCount(object columns)
    {
      return (int)InvokeHelper.GetProperty(columns, "Count");
    }

    public object[] ReadRowValues(int sheetIndex, int rowIndex, int fromColumnIndex, int toColumnIndex)
    {
      try
      {
        var sheet = this.GetSheet(sheetIndex);
        if (sheet == null)
          return null;
        var usedRange = GetUsedRange(sheet);
        if (usedRange == null)
          return null;
        return ReadRowValues(usedRange, rowIndex, fromColumnIndex, toColumnIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public object[] ReadColumnValues(int sheetIndex, int columnIndex, int fromRowIndex, int toRowIndex)
    {
      try
      {
        var sheet = this.GetSheet(sheetIndex);
        if (sheet == null)
          return null;
        var usedRange = GetUsedRange(sheet);
        if (usedRange == null)
          return null;
        return ReadColumnValues(usedRange, columnIndex, fromRowIndex, toRowIndex);
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static object[] ReadRowValues(object range, int rowIndex, int fromColumnIndex, int toColumnIndex)
    {
      try
      {
        if (range == null)
          return null;

        var rows = GetRows(range);
        var columns = GetColumns(range);
        var rangeSize = new Int32Size(GetRowCount(rows), GetColumnCount(columns));

        var result = new List<object>();

        for (int col = fromColumnIndex; col <= toColumnIndex; col++)
        {
          if (col > rangeSize.Width || rowIndex > rangeSize.Height)
            break;

          object cell = InvokeHelper.GetProperty(range, "Item", rowIndex, col);
          object cellValue = InvokeHelper.GetProperty(cell, "Value");

          result.Add(cellValue);

          cellValue = null;
          cell = null;
        }
        return result.ToArray();
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }
    public static object[] ReadColumnValues(object range, int columnIndex, int fromRowIndex, int toRowIndex)
    {
      try
      {
        if (range == null)
          return null;

        var rows = GetRows(range);
        var columns = GetColumns(range);
        var rangeSize = new Int32Size(GetRowCount(rows), GetColumnCount(columns));

        var result = new List<object>();

        for (int row = fromRowIndex; row <= toRowIndex; row++)
        {
          if (row > rangeSize.Height || columnIndex > rangeSize.Width)
            break;

          object cell = InvokeHelper.GetProperty(range, "Item", row, columnIndex);
          object cellValue = InvokeHelper.GetProperty(cell, "Value");

          result.Add(cellValue);

          cellValue = null;
          cell = null;
        }
        return result.ToArray();
      }
      catch (System.Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
        return null;
      }
    }

    public object GetSheet(int sheetIndex /*Start from 1*/)
    {
      return InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
    }
    public object GetSheet(String sheetName)
    {
      int sheetsCount = this.GetWorksheetCount();
      if (sheetsCount == 0)
        return null;

      if (this.IsCsv)
        return InvokeHelper.GetProperty(this._worksheets, "Item", 1);

      for (int i = 1; i <= sheetsCount; i++)
      {
        object item = InvokeHelper.GetProperty(this._worksheets, "Item", i);
        if (item == null)
          return null;

        if (string.Equals(sheetName, InvokeHelper.GetProperty(item, "Name") as string, StringComparison.CurrentCultureIgnoreCase))
          return item;
      }
      return null;
    }
    public Int32 GetSheetIndex(String sheetName)
    {
      int sheetsCount = this.GetWorksheetCount();
      if (sheetsCount == 0)
        return -1;

      if (this.IsCsv)
        return 1;

      for (int i = 1; i <= sheetsCount; i++)
      {
        object item = InvokeHelper.GetProperty(this._worksheets, "Item", i);
        if (item == null)
        {
          item = null;
          return -1;
        }
        if (string.Equals(sheetName, InvokeHelper.GetProperty(item, "Name") as string, StringComparison.CurrentCultureIgnoreCase))
        {
          item = null;
          return i;
        }
        item = null;
      }
      return -1;
    }

    public String GetSheetName(int sheetIndex)
    {
      int sheetsCount = this.GetWorksheetCount();
      if (sheetIndex > sheetsCount)
        return null;

      var sheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
      return GetSheetName(sheet);
    }
    public String GetSheetName()
    {
      return GetSheetName(_worksheet);
    }
    public static String GetSheetName(object sheet)
    {
      if (sheet == null)
        return null;
      var ret = InvokeHelper.GetProperty(sheet, "Name");
      if (ret == null)
        return null;
      return ret.ToString();
    }
    public void SetSheetName(String sheetName)
    {
      SetSheetName(_worksheet, sheetName);
    }
    public static void SetSheetName(object sheet, String sheetName)
    {
      try
      {
        if (sheet == null)
          return;
        InvokeHelper.SetProperty(sheet, "Name", sheetName);
      }
      catch (Exception ex)
      {
        //ExceptionHandler.ThrowException(ex);
      }
    }

    public Int32 CopySheetFrom(int sheetIndex, String sheetName)
    {
      _worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
      InvokeHelper.CallMethod(_worksheet, "Copy", missing, _worksheet);
      _worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex + 1);
      SetSheetName(_worksheet, sheetName);

      return sheetIndex + 1;
    }
    public Int32 CopySheet(String sheetName)
    {
      _worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", this.GetWorksheetCount());
      InvokeHelper.CallMethod(_worksheet, "Copy", missing, _worksheet);
      SetSheetName(_worksheet, sheetName);

      return this.GetWorksheetCount() - 1;
    }
    public object AddSheet(String sheetName)
    {
      _worksheet = InvokeHelper.CallMethod(_worksheets, "Add", Type.Missing, _worksheet);
      InvokeHelper.SetProperty(_worksheet, "Name", sheetName);
      return _worksheet;
    }
    public void RemoveLastSheet()
    {
      int sheetsCount = this.GetWorksheetCount();
      if (sheetsCount > 1)
      {
        object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetsCount);
        InvokeHelper.CallMethod(worksheet, "Delete");
      }
    }
    public void SetActiveSheet(int sheetIndex)
    {
      int sheetsCount = this.GetWorksheetCount();
      sheetIndex = Math.Min(sheetsCount, sheetIndex);

      object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
      InvokeHelper.CallMethod(worksheet, "Activate");
    }
    public void SaveWorkBook(string outputFile)
    {
      InvokeHelper.CallMethod(_workbook, "SaveCopyAs", outputFile);
    }
    public void SaveWorkBook()
    {
      if (System.IO.File.Exists(this._excelFileName))
        InvokeHelper.CallMethod(_workbook, "Save");
      else
        this.SaveWorkBook(this._excelFileName);
    }
    public bool CloseExcel()
    {
      try
      {
        if (this._application == null)
          return true;

        if (!String.IsNullOrEmpty(_excelFileName))
          InvokeHelper.CallMethod(_workbook, "Close", true, _excelFileName);
        InvokeHelper.CallMethod(_application, "Quit");

        System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheet);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheets);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbooks);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);

        return true;
      }
      catch (Exception)
      {
        Process[] myProcesses = Process.GetProcessesByName("Excel");

        foreach (Process myProcess in myProcesses)
        {
          DateTime startTime = myProcess.StartTime;
          if (startTime >= _beforeTime && startTime <= _afterTime)
          {
            myProcess.Kill();
          }
        }

        return false;
      }
      finally
      {
        // release object
        _worksheet = null;
        _worksheets = null;
        _workbook = null;
        _workbooks = null;
        _application = null;

        GcCollect();
      }
    }
    private void GcCollect()
    {
      // release memory
      GC.Collect();
      GC.WaitForPendingFinalizers();
      GC.Collect();
      GC.WaitForPendingFinalizers();
    }
    public int GetWorksheetCount()
    {
      if (_worksheets == null)
        return 0;
      return (int)InvokeHelper.GetProperty(_worksheets, "Count");
    }

    public void InsertRow(int sheetIndex, int baseRow, bool copyContent = true)
    {
      object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
      InvokeHelper.CallMethod(worksheet, "Activate");
      object rowRange = InvokeHelper.GetProperty(worksheet, "Rows", baseRow);
      InvokeHelper.CallMethod(rowRange, "Select");
      if (copyContent)
        InvokeHelper.CallMethod(rowRange, "Copy");
      InvokeHelper.CallMethod(rowRange, "Insert");
    }
    public void InsertColumn(int sheetIndex, int baseColumn, bool copyContent = true)
    {
      object worksheet = InvokeHelper.GetProperty(this._worksheets, "Item", sheetIndex);
      InvokeHelper.CallMethod(worksheet, "Activate");
      object rowRange = InvokeHelper.GetProperty(worksheet, "Columns", baseColumn);
      InvokeHelper.CallMethod(rowRange, "Select");
      if (copyContent)
        InvokeHelper.CallMethod(rowRange, "Copy");
      InvokeHelper.CallMethod(rowRange, "Insert");
    }

    #region Copy
    public static void Copy(object[,] values)
    {
      DateTime beforeTime = DateTime.Now;
      Type oExcel = Type.GetTypeFromProgID("Excel.Application");
      if (oExcel == null)
        throw new Exception("Office excel has not been installed!");

      object application = Activator.CreateInstance(oExcel);
      object workbooks = InvokeHelper.GetProperty(application, "Workbooks");
      object workbook = InvokeHelper.CallMethod(workbooks, "Add", true);
      object worksheet = InvokeHelper.GetProperty(workbook, "ActiveSheet");
      object cell1 = InvokeHelper.GetProperty(worksheet, "Cells", 1, 1);
      object cell2 = InvokeHelper.GetProperty(worksheet, "Cells", values.GetLength(0), values.GetLength(1));
      object range = InvokeHelper.GetProperty(worksheet, "Range", cell1, cell2);
      InvokeHelper.SetProperty(range, "Value", values);
      InvokeHelper.CallMethod(range, "Select");
      InvokeHelper.CallMethod(range, "Copy");
      DateTime afterTime = DateTime.Now;

      InvokeHelper.CallMethod(application, "Quit");

      System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
      System.Runtime.InteropServices.Marshal.ReleaseComObject(application);

      // release object
      worksheet = null;
      workbook = null;
      workbooks = null;
      application = null;
    }
    #endregion

    public bool IsCsv
    {
      get
      {
        return string.Equals(System.IO.Path.GetExtension(this._excelFileName), ".csv", StringComparison.CurrentCultureIgnoreCase);
      }
    }

    private object missing = Missing.Value;
    private object _application;
    private object _workbooks;
    private object _workbook;
    private object _worksheets;
    private object _worksheet;

    private DateTime _beforeTime;
    private DateTime _afterTime;
    private string _excelFileName = null;
    private int _decimalPlaces = 3;
    public string NumberFormatString
    {
      get
      {
        string sNumberFormat = "#,";
        for (int i = 0; i < _decimalPlaces; i++)
          sNumberFormat += "#";
        sNumberFormat += "0.";
        for (int i = 0; i < _decimalPlaces; i++)
          sNumberFormat += "0";
        return sNumberFormat;
      }
    }
  }
  public enum XlHAlign
  {
    xlHAlignRight = -4152,
    xlHAlignLeft = -4131,
    xlHAlignJustify = -4130,
    xlHAlignDistributed = -4117,
    xlHAlignCenter = -4108,
    xlHAlignGeneral = 1,
    xlHAlignFill = 5,
    xlHAlignCenterAcrossSelection = 7,
  }
  public enum XlVAlign
  {
    xlVAlignBottom = -4107,
    xlVAlignCenter = -4108,
    xlVAlignDistributed = -4117,
    xlVAlignJustify = -4130,
    xlVAlignTop = -4160,
  }
  public enum XlSheetVisibility
  {
    xlSheetVisible = -1,
    xlSheetHidden = 0,
    xlSheetVeryHidden = 2,
  }
  public enum XlCellErrorValue
  {
    ErrDiv0 = -2146826281, //#Div/0!
    ErrNA = -2146826246, //#N/A
    ErrName = -2146826259, //#Name?
    ErrNull = -2146826288, //#Null!
    ErrNum = -2146826252, //#Num!
    ErrRef = -2146826265, //#Ref!
    ErrValue = -2146826273 //#Value!
  }
  public class Int32Size
  {
    public Int32Size(Int32 width, Int32 height)
    {
      _width = width;
      _height = height;
    }
    public bool IsEmpty
    {
      get
      {
        return _width < 0;
      }
    }

    public Int32 Width
    {
      get
      {
        return _width;
      }
      set
      {
        if (value < 0)
        {
          throw new System.ArgumentException("Width And Height Cannot Be Negative");
        }
        _width = value;
      }
    }

    public Int32 Height
    {
      get
      {
        return _height;
      }
      set
      {
        if (value < 0)
        {
          throw new System.ArgumentException("Width And Height Cannot Be Negative");
        }

        _height = value;
      }
    }
    private Int32 _width;
    private Int32 _height;

    public Int32 MaxDimension()
    {
      return Math.Abs(this.Width) > Math.Abs(this.Height) ? this.Width : this.Height;
    }
  }
  public class ExcelRangeFormat
  {
    public ExcelRangeFormat()
    {
      RowStartIndex = 1;
      ColumnStartIndex = 1;
      RowCount = 1;
      ColumnCount = 1;
      Foreground = Colors.Black;
      Background = Brushes.White;
      Bold = false;
      NeedNumberFormat = false;
      TextAlign = null;
      TextVerticalAlign = null;
    }
    public static Int32 ColorToColorIndex(Brush brush)
    {
      return ColorToColorIndex(DPApplication.BrushToColor(brush));
    }
    public static Int32 ColorToColorIndex(Color color)
    {
      Int32 colorIndex = 1;
      if (color == Colors.Black) // 1: Black, #000000, 黑色
        colorIndex = 1;
      else if (color == Colors.White || color == Colors.Transparent || color.A == 0) // 1: White, #FFFFFF, 白色
        colorIndex = 2;
      else if (color == Colors.Red || color == DPApplication.Red) // 3: Red, #FF0000, 红色
        colorIndex = 3;
      else if (color == Colors.SpringGreen || color == DPApplication.Green || color.ToString().EndsWith("00FF00")) // 4: BrightGreen, #00FF00, 鲜绿色
        colorIndex = 4;
      else if (color == Colors.Blue || color == DPApplication.Purple) // 5: Blue, #0000FF, 蓝色
        colorIndex = 5;
      else if (color == Colors.Yellow || color == DPApplication.Yellow) // 6: Yellow, #FFFF00, 黄色
        colorIndex = 6;
      else if (color == Colors.Pink || color.ToString().EndsWith("FF00FF")) // 7: Pink, #FF00FF, 粉红色
        colorIndex = 7;
      else if (color == Colors.Turquoise || color.ToString().EndsWith("00FFFF")) // 8: Turquoise, #00FFFF, 青绿色
        colorIndex = 8;
      else if (color == Colors.DarkRed || color.ToString().EndsWith("800000")) // 9: DarkRed, #800000, 深红色
        colorIndex = 9;
      else if (color == Colors.Green || color.ToString().EndsWith("008000")) // 10: Green, #008000, 绿色
        colorIndex = 10;
      else if (color == Colors.DarkBlue || color.ToString().EndsWith("000080")) // 11: Dark Blue, #000080, 深蓝色
        colorIndex = 11;
      else if (color.ToString().EndsWith("808000"))// 12: DarkYellow, #808000, 深黄色
        colorIndex = 12;
      else if (color == Colors.Violet || color.ToString().EndsWith("800080")) // 13: Violet, #800080, 紫罗兰
        colorIndex = 13;
      else if (color == Colors.Teal) // 14: Teal, #008080, 青色
        colorIndex = 14;
      else if (color == Colors.LightGray) // 15: Gray-25%, #C0C0C0, 灰－25％
        colorIndex = 15;
      else if (color == Colors.Gray) // 16: Gray-50%, #808080, 灰－50％
        colorIndex = 16;
      else if (color.ToString().EndsWith("9999FF")) // 17: Periwinkle, #9999FF, 海螺色
        colorIndex = 17;
      else if (color == Colors.Plum || color.ToString().EndsWith("993366")) // 18: Plum+, #993366, 梅红色
        colorIndex = 18;
      else if (color == Colors.Ivory || color.ToString().EndsWith("FFFFCC")) // 19: Ivory, #FFFFCC, 象牙色
        colorIndex = 19;
      else if (color.ToString().EndsWith("CCFFFF")) // 20: LiteTurquoise, #CCFFFF, 浅青绿
        colorIndex = 20;
      else if (color.ToString().EndsWith("660066")) // 21: DarkPurple, #660066, 深紫色
        colorIndex = 21;
      else if (color == Colors.Coral || color.ToString().EndsWith("FF8080")) // 22: Coral, #FF8080, 珊瑚红
        colorIndex = 22;
      else if (color.ToString().EndsWith("0066CC")) // 23: OceanBlue, #0066CC, 海蓝色
        colorIndex = 23;
      else if (color.ToString().EndsWith("CCCCFF")) // 24: IceBlue, #CCCCFF, 冰蓝
        colorIndex = 24;
      else if (color == Colors.DarkBlue || color.ToString().EndsWith("000080")) // 25: DarkBlue+, #000080, 深蓝色
        colorIndex = 25;
      else if (color == Colors.HotPink) // 26: Pink+, #FF00FF, 粉红色
        colorIndex = 26;
      else if (color == Colors.Yellow || color.ToString().EndsWith("FFFF00")) // 27: Yellow+, #FFFF00, 黄色
        colorIndex = 27;
      else if (color == Colors.DarkTurquoise || color.ToString().EndsWith("00FFFF")) // 28: Turquoise+, #00FFFF, 青绿色
        colorIndex = 28;
      else if (color == Colors.DarkViolet) // 29: Violet+, #800080, 紫罗兰
        colorIndex = 29;
      else if (color == Colors.DarkRed || color.ToString().EndsWith("800000")) // 30: DarkRed+, #800000, 深红色
        colorIndex = 30;
      else if (color.ToString().EndsWith("008080")) // 31: Teal+, #008080, 青色
        colorIndex = 31;
      else if (color == Colors.DarkBlue || color.ToString().EndsWith("008000")) // 32: Blue+, #0000FF, 蓝色
        colorIndex = 32;
      else if (color == Colors.SkyBlue || color.ToString().EndsWith("00CCFF")) // 33: SkyBlue, #00CCFF, 天蓝色
        colorIndex = 33;
      else if (color == Colors.MediumTurquoise) // 34: LightTurquoise, #CCFFFF, 浅青绿
        colorIndex = 34;
      else if (color == Colors.LightGreen) // 35: LightGreen, #CCFFCC, 浅绿色
        colorIndex = 35;
      else if (color == Colors.LightYellow) // 36: LightYellow, #FFFF99, 浅黄色
        colorIndex = 36;
      else if (color == Colors.LightSkyBlue || color == Colors.LightBlue) // 37: PaleBlue, #99CCFF, 淡蓝色
        colorIndex = 37;
      else if (color.ToString().EndsWith("FF99CC")) // 38: Rose, #FF99CC, 玫瑰红
        colorIndex = 38;
      else if (color == Colors.Lavender) // 39: Lavender, #CC99FF, 淡紫色
        colorIndex = 39;
      else if (color == Colors.Tan) // 40: Tan, #FFCC99, 茶色
        colorIndex = 40;
      else if (color == Colors.LightBlue || color.ToString().EndsWith("3366FF")) // 41: LightBlue, #3366FF, 浅蓝色
        colorIndex = 41;
      else if (color == Colors.Aqua || color.ToString().EndsWith("33CCCC")) // 42: Aqua, #33CCCC, 水绿色
        colorIndex = 42;
      else if (color == Colors.YellowGreen || color == Colors.Lime || color.ToString().EndsWith("99CC00")) // 43: Lime, #99CC00, 酸橙色
        colorIndex = 43;
      else if (color == Colors.Gold || color.ToString().EndsWith("FFCC00")) // 44: Gold, #FFCC00, 金色
        colorIndex = 44;
      else if (color == Colors.Gold || color.ToString().EndsWith("FF9900")) // 45: LightOrange, #FF9900, 浅橙色
        colorIndex = 45;
      else if (color == Colors.Orange || color.ToString().EndsWith("FF6600")) // 46: Orange, #FF6600, 橙色
        colorIndex = 46;
      else if (color.ToString().EndsWith("666699")) // 47: Blue-Gray, #666699, 蓝－灰
        colorIndex = 47;
      else if (color.ToString().EndsWith("969696")) // 48: Gray-40%, #969696, 灰－40％
        colorIndex = 48;
      else if (color.ToString().EndsWith("003366")) // 49: DarkTeal, #003366, 深青
        colorIndex = 49;
      else if (color == Colors.SeaGreen) // 50: SeaGreen, #339966, 海绿
        colorIndex = 50;
      else if (color == Colors.DarkGreen) // 51: DarkGreen, #003300, 深绿
        colorIndex = 51;
      else if (color == Colors.DarkOliveGreen) // 52: OliveGreen, #333300, 橄榄色
        colorIndex = 52;
      else if (color == Colors.Brown || color.ToString().EndsWith("993300")) // 53: Brown, #993300, 褐色
        colorIndex = 53;
      else if (color == Colors.Plum || color.ToString().EndsWith("993366")) // 54: Plum, #993366, 梅红色
        colorIndex = 54;
      else if (color == Colors.Indigo || color.ToString().EndsWith("333399")) // 55: Indigo, #333399, 靛蓝
        colorIndex = 54;
      else if (color == Colors.DarkGray) // 56: Gray-80%, #333333, 灰－80％
        colorIndex = 56;

      return colorIndex;
    }

    public object ValueObject { get; set; }
    public Int32 RowStartIndex { get; set; } // Based on 1
    public Int32 ColumnStartIndex { get; set; } // Based on 1
    public Int32 RowCount { get; set; }
    public Int32 ColumnCount { get; set; }
    public Color Foreground { get; set; }
    public Brush Background { get; set; }
    public Boolean Bold { get; set; }
    public Boolean NeedNumberFormat { get; set; }
    public VerticalAlignment? TextVerticalAlign { get; set; }
    public TextAlignment? TextAlign { get; set; }

    public Int32 RowEndIndex { get { return RowStartIndex + RowCount - 1; } }
    public Int32 ColumnEndIndex { get { return ColumnStartIndex + ColumnCount - 1; } }
    public Int32 ForegroundColorIndex { get { return ExcelRangeFormat.ColorToColorIndex(Foreground); } }
    public Int32 BackgroundColorIndex { get { return ExcelRangeFormat.ColorToColorIndex(Background); } }
    public XlHAlign? HAlign
    {
      get
      {
        if (this.TextAlign == null)
          return null;

        switch (this.TextAlign.Value)
        {
          case TextAlignment.Left:
            return XlHAlign.xlHAlignLeft;
          case TextAlignment.Right:
            return XlHAlign.xlHAlignRight;
          default:
            return XlHAlign.xlHAlignCenter;
        }
      }
    }
    public XlVAlign? VAlign
    {
      get
      {
        if (this.TextVerticalAlign == null)
          return null;

        switch (this.TextVerticalAlign.Value)
        {
          case VerticalAlignment.Top:
            return XlVAlign.xlVAlignTop;
          case VerticalAlignment.Bottom:
            return XlVAlign.xlVAlignBottom;
          default:
            return XlVAlign.xlVAlignCenter;
        }
      }
    }
  }
}
