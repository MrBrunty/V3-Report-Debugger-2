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
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = System.Windows.Window;
using Microsoft.Win32;

namespace Custom_Report_Debugger_2._0
{
    class TestingFields
    {
        public static (List<string>, bool) TestData(string testData)
        {
            var testFields = new List<string>();
            bool exceptionThrown;
            var index = 0;

            try
            {
                var testFieldArray = testData.Split(' ');

                foreach (var field in testFieldArray)
                {
                    if (field == "MERGEFIELD")
                    {
                        testFields.Add("{" + field.Trim() + " " + testFieldArray[index + 1].Trim() + "}");
                    }
                    index++;
                }
                exceptionThrown = false;
                return (testFields, exceptionThrown);
            }

            catch (Exception)
            {                
                exceptionThrown = true;
                return (testFields, exceptionThrown);
            }
        }
    }
}