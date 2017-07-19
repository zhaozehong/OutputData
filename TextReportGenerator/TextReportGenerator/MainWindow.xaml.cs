using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TextReportGenerator
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    public MainWindow()
    {
      InitializeComponent();
    }

    private void OutputExcelFile()
    {
      var header = new Dictionary<String, object>();
      header.Add("Report Type", "Text Report");
      header.Add("Chart Type", "Details Chart");
      header.Add("Company", "WAI");
      //2017-07-18 16:58
      header.Add("Start Date", new DateTime(2017, 6, 30).ToString("yyyy-MM-dd hh:mm"));
      header.Add("End Date", new DateTime(2017, 7, 14));
      header.Add("Description", "This is a test data, not a real data.");

      var dataArray = new object[5, 10];
      dataArray[0, 0] = 5.0;
      dataArray[0, 1] = 5.1;
      dataArray[0, 2] = 5.2;
      dataArray[0, 3] = 5.3;
      dataArray[0, 4] = 5.4;
      dataArray[0, 5] = 5.5;
      dataArray[0, 6] = 5.6;
      dataArray[0, 7] = 5.7;
      dataArray[0, 8] = 5.8;
      dataArray[0, 9] = 5.9;

      dataArray[1, 0] = 15.0;
      dataArray[1, 1] = 15.1;
      dataArray[1, 2] = 15.2;
      dataArray[1, 3] = 15.3;
      dataArray[1, 4] = 15.4;
      dataArray[1, 5] = 15.5;
      dataArray[1, 6] = 15.6;
      dataArray[1, 7] = 15.7;
      dataArray[1, 8] = 15.8;
      dataArray[1, 9] = 15.9;

      dataArray[2, 0] = 25.0;
      dataArray[2, 1] = 25.1;
      dataArray[2, 2] = 25.2;
      dataArray[2, 3] = 25.3;
      dataArray[2, 4] = 25.4;
      dataArray[2, 5] = 25.5;
      dataArray[2, 6] = 25.6;
      dataArray[2, 7] = 25.7;
      dataArray[2, 8] = 25.8;
      dataArray[2, 9] = 25.9;

      dataArray[3, 0] = 35.0;
      dataArray[3, 1] = 35.1;
      dataArray[3, 2] = 35.2;
      dataArray[3, 3] = 35.3;
      dataArray[3, 4] = 35.4;
      dataArray[3, 5] = 35.5;
      dataArray[3, 6] = 35.6;
      dataArray[3, 7] = 35.7;
      dataArray[3, 8] = 35.8;
      dataArray[3, 9] = 35.9;

      dataArray[4, 0] = 45.0;
      dataArray[4, 1] = 45.1;
      dataArray[4, 2] = 45.2;
      dataArray[4, 3] = 45.3;
      dataArray[4, 4] = 45.4;
      dataArray[4, 5] = 45.5;
      dataArray[4, 6] = 45.6;
      dataArray[4, 7] = 45.7;
      dataArray[4, 8] = 45.8;
      dataArray[4, 9] = 45.9;

      var toExcel = new DataToExcel();
      toExcel.Save(@"C:\Users\zehong.zhao\Desktop\Fake.xlsx", header, dataArray);
    }

    private void btnGenerate_Click(object sender, RoutedEventArgs e)
    {
      OutputExcelFile();
    }
  }
}
