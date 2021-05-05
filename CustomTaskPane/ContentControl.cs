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
			if( Env == null )
			{
				var envPath = System.IO.Path.GetDirectoryName( AddIn.XllPath );
				Env = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync( null, envPath );
			}

			var wv = new Microsoft.Web.WebView2.WinForms.WebView2{ Dock=DockStyle.Fill };
			await wv.EnsureCoreWebView2Async( Env );
			this.Controls.Add( wv );
			wv.CoreWebView2.Navigate( "https://github.com/Excel-DNA/Samples/tree/master/CustomTaskPane" );
        }
    }
}
