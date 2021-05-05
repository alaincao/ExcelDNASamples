using ExcelDna.Integration;

namespace CustomTaskPane
{
    public class AddIn : IExcelAddIn
    {
		internal static string XllPath;

        public void AutoOpen()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
			XllPath = (string)ExcelDna.Integration.XlCall.Excel( ExcelDna.Integration.XlCall.xlGetName );
        }

        public void AutoClose()
        {
        }

    }
}
