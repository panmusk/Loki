/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: RL2570
 * Data: 28.06.2018
 * Godzina: 09:47
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Linq;


namespace Loki
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public static WordTools WdTools;
		private System.Drawing.Font footerFont;
		public MainForm()
		{
			// The InitializeComponent() call is required for Windows Forms designer support.
			InitializeComponent();
			Console.SetOut(new TextBoxWriter(LogTxb));
		}

		void RptExecuteBtnClick(object sender, EventArgs e)
		{
			if (CheckPrefix()) {
				MessageBox.Show("Niepoprawne znaki w prefiksie");
				DestFilenamePrefixTbx.Focus();
				return;
			}
			if (OldTextb.Text.Length > 0 && NewTextb.Text.Length > 0 && DestFolderTxb.Text.Length > 0) {
				if (FileRadio.Checked && File.Exists(FilePathTxb.Text)) {
					RptTools.ReplaceTextsRpt(FilePathTxb.Text, OldTextb.Text, NewTextb.Text, DestFolderTxb.Text, DestFilenamePrefixTbx.Text);
				} else {
					if (!FileRadio.Checked && Directory.Exists(FilePathTxb.Text)) {
						string[] Files;
						Files = RecursiveChkb.Checked ? Directory.GetFiles(FilePathTxb.Text, "*.rpt", SearchOption.AllDirectories) : Directory.GetFiles(FilePathTxb.Text, "*.rpt");
						foreach (string FileName in Files) {
							RptTools.ReplaceTextsRpt(FileName, OldTextb.Text, NewTextb.Text, DestFolderTxb.Text, DestFilenamePrefixTbx.Text);
						}
						Console.WriteLine("Done.");
						
					}
				}
			} else {
				MessageBox.Show("Brak wypełnionych parametrów");
			}
			//
		}
		void DirRadioCheckedChanged(object sender, EventArgs e)
		{
			RecursiveChkb.Enabled = DirRadio.Checked ? true : false;
			FilePathTxb.Text = "";
		}
		void Button2Click(object sender, EventArgs e)
		{
			if (FileRadio.Checked) {
				OpenFileDialog fd = new OpenFileDialog();
				fd.Filter = "Crystal Reports file (*.rpt), Word docx files (*.docx)|*.rpt;*.docx";
				fd.Multiselect = false;
				fd.ShowDialog();
				if (fd.FileNames.Length > 0)
					FilePathTxb.Text = fd.FileNames[0];
			} else {
				FolderBrowserDialog fd = new FolderBrowserDialog();
				fd.ShowDialog();
				if (fd.SelectedPath != "")
					FilePathTxb.Text = fd.SelectedPath;
			}
		}

		void Button3Click(object sender, EventArgs e)
		{
			LogTxb.Clear();
		}
		void DestFolderBrowseBtnClick(object sender, EventArgs e)
		{
			FolderBrowserDialog fd = new FolderBrowserDialog();
			fd.ShowDialog();
			if (fd.SelectedPath != "")
				DestFolderTxb.Text = fd.SelectedPath;	
		}
		bool CheckPrefix()
		{
			Regex R = new Regex("[\\*\\.\"\\/\\\\[\\]\\:\\;\\|\\=\\,]");
			MatchCollection M = R.Matches(DestFilenamePrefixTbx.Text);
			return M.Count > 0;
		}
		void MainFormKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
				this.Close();
		}
		async void FtrExecuteBtnClick(object sender, EventArgs e)
		{
			if (FileRadio.Checked && File.Exists(FilePathTxb.Text)) {
				WdTools = WdTools ?? new WordTools();
				int result = await WdTools.InsertFooter(FilePathTxb.Text, footerContentTxb.Text, footerFont);
				Console.WriteLine("Done.");
			} else {
				if (!FileRadio.Checked && Directory.Exists(FilePathTxb.Text)) {
					string[] Files;
					Files = RecursiveChkb.Checked ? Directory.GetFiles(FilePathTxb.Text, "*.docx", SearchOption.AllDirectories) : Directory.GetFiles(FilePathTxb.Text, "*.docx");
					foreach (string FileName in Files) {
						WdTools = WdTools ?? new WordTools();
						int result = await WdTools.InsertFooter(FileName, footerContentTxb.Text, footerFont);
					}
					Console.WriteLine("Done.");
				}
			}
		}
		private async void StatsExecuteBtnClick(object sender, EventArgs e)
		{
			if (FileRadio.Checked && File.Exists(FilePathTxb.Text)) {
				WdTools = WdTools ?? new WordTools();
				int result = await WdTools.PrintDocxStats(FilePathTxb.Text, this.statsXpathChkb.Checked, this.statsSectionsChkb.Checked, this.statsFooterHeadersChkb.Checked);
			} else {
				if (!FileRadio.Checked && Directory.Exists(FilePathTxb.Text)) {
					string[] Files;
					Files = RecursiveChkb.Checked ? Directory.GetFiles(FilePathTxb.Text, "*.docx", SearchOption.AllDirectories) : Directory.GetFiles(FilePathTxb.Text, "*.docx");
					foreach (string FileName in Files) {
						WdTools = WdTools ?? new WordTools();
						int result = await WdTools.PrintDocxStats(FileName, this.statsXpathChkb.Checked, this.statsSectionsChkb.Checked, this.statsFooterHeadersChkb.Checked);
					}
				}
			}
			Console.WriteLine("Done.");
		}

		void MainFormFormClosed(object sender, FormClosedEventArgs e)
		{
			if (WdTools != null) {
				try {
					WdTools.Dispose();
				} catch {
					Console.WriteLine("Closing Word application failed. Kill winword.exe process manualy.");
				}
			}
				
		}
		private async void WdExecuteBtnClick(object sender, EventArgs e)
		{
			string Action = "";
			if (this.XpathFragmRadio.Checked) {
				Action = "XpathReplaceFragment";
			}
			if (this.XpathExactRadio.Checked) {
				Action = "XpathReplaceExact";
			}
			if (this.TextReplRadio.Checked) {
				Action = "TextReplace";
			}
			WdTools = WdTools ?? new WordTools();
			if (FileRadio.Checked && File.Exists(FilePathTxb.Text)) {
				var result = (System.Threading.Tasks.Task<int>)WdTools.GetType().GetMethod(Action).Invoke(WdTools, new object[] {
					FilePathTxb.Text,
					wdOldTextb.Text,
					wdNewTextb.Text
				});
				await result;
				Console.WriteLine("Done.");
			} else {
				if (!FileRadio.Checked && Directory.Exists(FilePathTxb.Text)) {
					string[] Files;
					Files = RecursiveChkb.Checked ? Directory.GetFiles(FilePathTxb.Text, "*.docx", SearchOption.AllDirectories) : Directory.GetFiles(FilePathTxb.Text, "*.docx");
					foreach (string FileName in Files) {
						var result = (System.Threading.Tasks.Task<int>)WdTools.GetType().GetMethod(Action).Invoke(WdTools, new object[] {
							FileName,
							wdOldTextb.Text,
							wdNewTextb.Text
						});
						await result;
					}
					Console.WriteLine("Done.");
				}
			}
		}
		void MainFormResize(object sender, EventArgs e)
		{
			if (WindowState == FormWindowState.Minimized) {
				TextBoxWriter.SendMessage(LogTxb.Handle, 0x000B, false, 0);
			} else {
				TextBoxWriter.SendMessage(LogTxb.Handle, 0x000B, true, 0);
			}
		}
		void FontCmbxClick(object sender, EventArgs e)
		{
			if (FontCmbx.DataSource == null) {
				IList<string> fontNames = FontFamily.Families.Select(f => f.Name).ToList();
				FontCmbx.DataSource = fontNames;
			}
		}
		void FontCmbxSelectedValueChanged(object sender, EventArgs e)
		{
			footerFont = new System.Drawing.Font(FontCmbx.Text, float.Parse(fontSizeTxb.Text));
		}
		void FontSizeTxbTextChanged(object sender, EventArgs e)
		{
			footerFont = new System.Drawing.Font(FontCmbx.Text, float.Parse(fontSizeTxb.Text));
		}
	
	}
}