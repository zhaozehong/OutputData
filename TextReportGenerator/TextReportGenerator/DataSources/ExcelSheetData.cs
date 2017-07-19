using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TextReportGenerator
{
  public class ExcelSheetData
  {
    public ExcelSheetData(object[,] dataArray, List<ExcelRangeFormat> rangeFormats = null, String sheetName = "Sheet1")
    {
      _dataArray = dataArray;
      if (rangeFormats != null)
        _rangeFormats.AddRange(rangeFormats);
      SheetName = sheetName;
    }

    public object[,] DataArray { get { return _dataArray; } }
    public List<ExcelRangeFormat> RangeFormats { get { return _rangeFormats; } }

    private object[,] _dataArray = null;
    private List<ExcelRangeFormat> _rangeFormats = new List<ExcelRangeFormat>();

    public String SheetName { get; set; }
  }
}
