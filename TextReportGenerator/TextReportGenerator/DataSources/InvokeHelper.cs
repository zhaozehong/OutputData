using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TextReportGenerator
{
  public class InvokeHelper
  {
    public static bool HasProperty(object reference, string propertyName)
    {
      var pi = reference.GetType().GetProperty(propertyName);
      return pi != null;
    }
    public static void SetProperty(object reference, string propertyName, object value)
    {
      var parameter = new object[1];
      parameter[0] = value;
      reference.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, reference, parameter);
    }
    public static object GetProperty(object reference, string propertyName, params object[] parameters)
    {
      return reference.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, reference, parameters);
    }
    public static object CallMethod(object reference, string propertyName, params object[] parameters)
    {
      return reference.GetType().InvokeMember(propertyName, BindingFlags.InvokeMethod, null, reference, parameters);
    }

    public static bool HasPropertyIncludeNonPublic(object reference, string propertyName)
    {
      var pi = reference.GetType().GetProperty(propertyName, BindingFlags.NonPublic | BindingFlags.Instance);
      return pi != null;
    }
    public static void SetPropertyIncludeNonPublic(object reference, string propertyName, object value)
    {
      var parameter = new object[1];
      parameter[0] = value;
      reference.GetType().InvokeMember(propertyName, BindingFlags.SetProperty | BindingFlags.NonPublic | BindingFlags.Instance, null, reference, parameter);
    }
    public static object GetPropertyIncludeNonPublic(object reference, string propertyName, params object[] parameters)
    {
      return reference.GetType().InvokeMember(propertyName, BindingFlags.GetProperty | BindingFlags.NonPublic | BindingFlags.Instance, null, reference, parameters);
    }
    public static object CallMethodIncludeNonPublic(object reference, string propertyName, params object[] parameters)
    {
      return reference.GetType().InvokeMember(propertyName, BindingFlags.InvokeMethod | BindingFlags.NonPublic | BindingFlags.Instance, null, reference, parameters);
    }
  }
}
