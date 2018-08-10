/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: RL2570
 * Data: 06.08.2018
 * Godzina: 08:56
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.IO;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer.ClientDoc;
namespace Loki
{
	/// <summary>
	/// Description of RptTools.
	/// </summary>
	public class RptTools
	{
		private RptTools()
		{
		}
		public static void ReplaceTextsRpt(String RptFilePath, String OldText, String NewText, String DestFolder, String DestFilePrefix){
			var CrDoc = new ReportDocument();
			Console.WriteLine(String.Format("Loading file {0}",RptFilePath));
			try {
				CrDoc.Load(RptFilePath);
				Console.WriteLine(String.Format("File {0} loaded successfully, report title: \"{1}\"", RptFilePath, CrDoc.SummaryInfo.ReportTitle));
			} catch (Exception e) {
				Console.WriteLine(String.Format("Error loading {0}: {1}", RptFilePath, e.Message));
				return;
			}
			ISCDReportClientDocument ClientDoc = CrDoc.ReportClientDocument;
			int ChangesCount = 0;
			foreach (var RepObj in CrDoc.ReportDefinition.ReportObjects) {
				switch(RepObj.GetType().ToString())
				{
				case "CrystalDecisions.CrystalReports.Engine.TextObject":
					TextObject TextObj = (TextObject)RepObj;
					if(TextObj.Text.Contains(OldText)){
						String TextBefore = TextObj.Text;
						TextObj.Text = TextObj.Text.Replace(OldText, NewText);
						Console.WriteLine("Replacing static text:");
						Console.WriteLine(TextBefore);
						Console.WriteLine("With:");
						Console.WriteLine(TextObj.Text);
						ChangesCount++;
				   }
					break;
				case "CrystalDecisions.CrystalReports.Engine.FieldObject":
						FieldObject FldObj = (FieldObject)RepObj;
						if(FldObj.DataSource.GetType().ToString() == "CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition"){
							FormulaFieldDefinition FldDef = (FormulaFieldDefinition)FldObj.DataSource;
							if(FldDef.Text.Contains(OldText)){
								String TextBefore = FldDef.Text;
								FldDef.Text = FldDef.Text.Replace(OldText, NewText);
								Console.WriteLine("Replacing formula:");
								Console.WriteLine(Regex.Escape(TextBefore));
								Console.WriteLine("With:");
								Console.WriteLine(Regex.Escape(FldDef.Text));
								ChangesCount++;		
						   }
						}
						break;
				}
			}
			if(ChangesCount > 0){
				Console.WriteLine(String.Format("No of changes in {0}: {1}", RptFilePath, ChangesCount));
		        FileInfo oldRptFileInfo = new FileInfo(RptFilePath);
		        String NewFileName = oldRptFileInfo.FullName.Replace(oldRptFileInfo.DirectoryName, DestFolder).Replace(@"\" + oldRptFileInfo.Name,@"\" + DestFilePrefix + oldRptFileInfo.Name);
				try {
					CrDoc.SaveAs(NewFileName);
					Console.WriteLine(String.Format("Saved as {0}", NewFileName));
				} catch (Exception e) {
					Console.WriteLine(String.Format("Error saving {0}: {1}", NewFileName, e.Message));
				}
				
			}else{
				Console.WriteLine(String.Format("Number \"New text\" occurrences in {0}, skipping", RptFilePath));
			}
		}
	}
}
