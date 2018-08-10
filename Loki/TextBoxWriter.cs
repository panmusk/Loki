/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: RL2570
 * Data: 06.08.2018
 * Godzina: 08:13
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Loki
{
	/// <summary>
	/// Appends text to Forms.TextBox
	/// </summary>
	public class TextBoxWriter : TextWriter
	{
		private TextBox textbox;
		[DllImport("user32.dll")]
		public static extern int SendMessage(IntPtr hWnd, Int32 wMsg, bool wParam, Int32 lParam);
		public TextBoxWriter(TextBox textbox)
		{
			this.textbox = textbox;
		}
	
		public override void Write(char value)
		{
	        
			if (value == '\r' || value == '\n') {
				//textbox.AppendText(value.ToString());
				//textbox.Invoke((Action)(()=>textbox.ResumeLayout()));
				textbox.Invoke((Action)(() => textbox.AppendText(value.ToString())));
				textbox.Invoke((Action)(() => SendMessage(textbox.Handle, 0x000B, true, 0)));
			} else {
				//textbox.Text += value;
				//textbox.Invoke((Action)(()=>textbox.Text += value));
				textbox.Invoke((Action)(() => SendMessage(textbox.Handle, 0x000B, false, 0)));
				//textbox.Invoke((Action)(()=>textbox.SuspendLayout()));
				textbox.Invoke((Action)(() => textbox.AppendText(value.ToString())));
			}
		}
	
		public override void Write(string value)
		{
			//textbox.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + value + "\r\n");
			//textbox.Text += value;
			//textbox.SelectionStart= textbox.Text.Length;
			//textbox.ScrollToCaret();
			textbox.Invoke((Action)(() => textbox.AppendText(string.Format("{0:yyyy-MM-dd HH:mm:ss }{1}", DateTime.Now, value))));
		}
		public override void WriteLine(string value)
		{
			//textbox.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss ") + value + "\r\n");
			//textbox.Text += value;
			//textbox.SelectionStart= textbox.Text.Length;
			//textbox.ScrollToCaret();
			textbox.Invoke((Action)(() => textbox.AppendText(string.Format("{0:yyyy-MM-dd HH:mm:ss }{1}\r\n", DateTime.Now, value))));
		}
	
		public override Encoding Encoding {
			get { return Encoding.UTF8; }
		}
	}
}
