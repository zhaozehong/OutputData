using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media;

namespace TextReportGenerator
{
  public class DPApplication
  {
    public static String GetValidFilePath(String filePath)
    {
      if (String.IsNullOrWhiteSpace(filePath))
        return String.Empty;

      var directoryName = Path.GetDirectoryName(filePath);
      var r = new Regex(String.Format("[{0}]", Regex.Escape(new String(Path.GetInvalidPathChars()))));
      directoryName = r.Replace(directoryName, "");

      var fileName = Path.GetFileName(filePath);
      r = new Regex(String.Format("[{0}]", Regex.Escape(new String(Path.GetInvalidFileNameChars()))));
      fileName = r.Replace(fileName, "");

      return Path.Combine(directoryName, fileName);
    }
    public static Color BrushToColor(Brush brush)
    {
      var solidBrush = brush as SolidColorBrush;
      if (solidBrush == null)
      {
        var gradientBrush = brush as GradientBrush;
        if (gradientBrush == null)
          return Colors.Transparent;

        var collection = (GradientStopCollection)gradientBrush.SafeGetValue(GradientBrush.GradientStopsProperty);
        if (collection != null)
        {
          GradientStop gs = collection.SafeGetAt(collection.SafeGetCount() - 1);
          if (gs != null)
            return (Color)gs.SafeGetValue(GradientStop.ColorProperty);
        }
      }
      return (Color)solidBrush.SafeGetValue(SolidColorBrush.ColorProperty);
    }


    public static readonly Color Red = ("#9FFF0000").SafeConvertInvariantStringTo<Color>();
    public static readonly Color Yellow = ("#9FFFFF00").SafeConvertInvariantStringTo<Color>();
    public static readonly Color Green = ("#9F00FF00").SafeConvertInvariantStringTo<Color>();
    public static readonly Color Purple = ("#9F0000FF").SafeConvertInvariantStringTo<Color>();
    public static readonly Brush BrushRed = ("#9FFF0000").SafeConvertInvariantStringTo<Brush>();
    public static readonly Brush BrushYellow = ("#9FFFFF00").SafeConvertInvariantStringTo<Brush>();
    public static readonly Brush BrushGreen = ("#9F00FF00").SafeConvertInvariantStringTo<Brush>();
    public static readonly Brush BrushPurple = ("#9F0000FF").SafeConvertInvariantStringTo<Brush>();
  }
}
