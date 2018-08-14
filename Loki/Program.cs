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
using System.Windows.Forms;
//using CrystalDecisions.CrystalReports;
//using CrystalDecisions.Enterprise;
//using CrystalDecisions.ReportAppServer.ReportDefModel;
//using CrystalDecisions.ReportAppServer.CommLayer;
//using CRAXDDRT;
//using CrystalDecisions.ReportSource;
//using CrystalDecisions.Shared;
//using CrystalDecisions.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace Loki
{
	/// <summary>
	/// Class with program entry point.
	/// </summary>
	internal sealed class Program
	{
		public static MainForm Frm;
		/// <summary>
		/// Program entry point.
		/// </summary>
		[STAThread]
		private static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Frm = new MainForm();
			Frm.ShowDialog();
		}
		public static void GetFiles(string dir, string pattern, ref List<string> files){
			try {
				for (int i = 0, maxLength = Directory.GetDirectories(dir).Length; i < maxLength; i++) {
					var subdir = Directory.GetDirectories(dir)[i];
					GetFiles(subdir, pattern, ref files);
				}
				for (int i = 0, maxLength = Directory.GetFiles(dir, pattern).Length; i < maxLength; i++) {
					var file = Directory.GetFiles(dir, pattern)[i];
					files.Add(file);
				}
			} catch (Exception) {
				Console.WriteLine(string.Format("skipping {0}, list size: {1}", dir, files.Count));
			}
		}
	}
}
