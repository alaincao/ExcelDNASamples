using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace CustomTaskPane
{
    [ComVisible(true)]
    public partial class ContentControl : UserControl
    {
        public ContentControl()
        {
            //InitializeComponent();
			Foo();
		}

		private static Microsoft.Web.WebView2.Core.CoreWebView2Environment Env = null;

		private async void Foo()
		{
			var dirPath = System.IO.Path.GetDirectoryName( AddIn.XllPath );
			if( Env == null )
			{
				Env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync( null, dirPath );
			}

			var wv = new Microsoft.Web.WebView2.WinForms.WebView2{ Dock=DockStyle.Fill };
			await wv.EnsureCoreWebView2Async( Env );
			this.Controls.Add( wv );

			wv.CoreWebView2.AddHostObjectToScript( "excel", ExcelDna.Integration.ExcelDnaUtil.Application );
			wv.CoreWebView2.SetVirtualHostNameToFolderMapping( "foobar.net", dirPath, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow );
			wv.CoreWebView2.Navigate( "https://foobar.net/barfoo.html" );
        }

		internal static void RunInDotnet()
		{
			dynamic application = ExcelDna.Integration.ExcelDnaUtil.Application;
			var cell = application.ActiveCell;
			cell.Value = 12;
			var cell2 = cell.Cells( 2, 2 );
			cell2.Value = 34;
		}
    }
}
