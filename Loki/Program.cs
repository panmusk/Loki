/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: RL2570
 * Data: 28.06.2018
 * Godzina: 09:47
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
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
	}
}
