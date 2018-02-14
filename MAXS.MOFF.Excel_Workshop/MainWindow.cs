using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace MAXS.MOFF.Excel_Workshop
{
    public partial class MainWindow : Form
    {
        public ExcelHandler myHandler = new ExcelHandler();
        public MainWindow()
        {
            InitializeComponent();
        }

        private void TestBtn_Click(object sender, EventArgs e)
        {
            //init excel doc
            myHandler.CreateWorkbook();
            Range myActiveRange = myHandler.workbook.ActiveSheet.Cells(1, 1);
            myActiveRange.Value = "Hello World";
            myHandler.workbook.Activate();

        }
    }
}

namespace Microsoft.Office.Interop.Excel
{
    public class ExcelHandler
    {
        public Application application { get; }
        public Workbook workbook { get; set; }
        public Worksheet active_Worksheet { get; set; }

        public ExcelHandler()
        {
            this.application = new Application();

        }

        public ExcelHandler(Workbook defaultWorkbook) : this()
        {
            this.workbook = defaultWorkbook;
        }
        public ExcelHandler(Workbook defaultWorkbook, Worksheet defaultWorksheet) : this()
        {
            this.workbook = defaultWorkbook;
            this.active_Worksheet = defaultWorksheet;
        }
        public void CreateWorkbook()
        {
            this.workbook = this.application.Workbooks.Add();
        }


        /*
         * Microsoft Visual Studio 17.0 .NET Framework Version v4.6.1
         * Automatically generated ..........
         *                         Equals
         *                         GetHasCode
         *                         ==
         *                         !=
         */

        public override bool Equals(object obj)
        {
            var handler = obj as ExcelHandler;
            return handler != null &&
                   EqualityComparer<Application>.Default.Equals(application, handler.application) &&
                   EqualityComparer<Workbook>.Default.Equals(workbook, handler.workbook) &&
                   EqualityComparer<Worksheet>.Default.Equals(active_Worksheet, handler.active_Worksheet);
        }

        public override int GetHashCode()
        {
            var hashCode = 168020149;
            hashCode = hashCode * -1521134295 + EqualityComparer<Application>.Default.GetHashCode(application);
            hashCode = hashCode * -1521134295 + EqualityComparer<Workbook>.Default.GetHashCode(workbook);
            hashCode = hashCode * -1521134295 + EqualityComparer<Worksheet>.Default.GetHashCode(active_Worksheet);
            return hashCode;
        }

        public static bool operator ==(ExcelHandler handler1, ExcelHandler handler2)
        {
            return EqualityComparer<ExcelHandler>.Default.Equals(handler1, handler2);
        }

        public static bool operator !=(ExcelHandler handler1, ExcelHandler handler2)
        {
            return !(handler1 == handler2);
        }
    }

}
