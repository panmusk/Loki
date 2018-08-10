/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: RL2570
 * Data: 06.08.2018
 * Godzina: 08:03
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;


namespace Loki
{
	/// <summary>
	/// Set of Microsoft office Word tools
	/// </summary>
	public class WordTools:IDisposable
	{
		#region IDisposable implementation

		public void Dispose()
		{
			//throw new NotImplementedException();
			try {
				Console.WriteLine("Closing Word application");
				WordApp.Quit(SaveChanges: false);
			} catch (Exception) {
				
				Console.WriteLine("Error closing Word application. Kill winword.exe manually.");
			}
			
		}
		private static string TagParents;

		#endregion

		public static Microsoft.Office.Interop.Word.Application WordApp;
		public WordTools()
		{
			if (WordApp == null) {
				Console.WriteLine("Starting Word app");
				WordApp = new ApplicationClass();
			}
			System.Diagnostics.Process[] WordProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
			if (WordProcesses.Length == 0) {
				WordApp = new ApplicationClass();
			}
			WordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
			Console.WriteLine("Word app started");
		}
		public static void StartWordApp()
		{
			
			if (WordApp == null) {
				Console.WriteLine("Starting Word app");
				WordApp = new ApplicationClass();
			}
			System.Diagnostics.Process[] WordProcesses = System.Diagnostics.Process.GetProcessesByName("winword");
			if (WordProcesses.Length == 0) {
				WordApp = new ApplicationClass();
			}
			WordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
			Console.WriteLine("Word app started");
		}
		public async System.Threading.Tasks.Task<int> InsertFooter(string DocFileName, string ftrContent, System.Drawing.Font ftrFont)
		{
			return await System.Threading.Tasks.Task.Run(() => {
				Document doc;
				StartWordApp();
				Console.WriteLine(string.Format("Loading file {0}", DocFileName));
				try {
					doc = WordApp.Documents.Open(FileName: DocFileName);
				} catch (Exception e) {
					Console.WriteLine(string.Format("Error loading File {0}: {1}", DocFileName, e.Message));
					return 0;
				}
				Console.WriteLine(string.Format("File {0} loaded", DocFileName));
				HeaderFooter footer = doc.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
				Paragraph par1;
				Paragraph par2;
				footer.Range.Text = "";
				footer.Range.Paragraphs.Add();
				footer.Range.Paragraphs.Add();
				par1 = footer.Range.Paragraphs[1];
				par2 = footer.Range.Paragraphs[footer.Range.Paragraphs.Count];
				par1.Range.Text = ftrContent;
				par2.Range.Text = "Strona x z y";
				par2.Range.Fields.Add(Range: par2.Range.Characters[8], Type: WdFieldType.wdFieldPage, PreserveFormatting: false);
				par2.Range.Fields.Add(Range: par2.Range.Characters[12], Type: WdFieldType.wdFieldNumPages, PreserveFormatting: false);
				par1.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
				par2.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
				footer.Range.Fields.Update();
				footer.Range.Font.Name = ftrFont.Name;
				footer.Range.Font.Size = ftrFont.Size;
				Console.WriteLine(string.Format("Footer after change in {0}: {1}", DocFileName, Regex.Escape(footer.Range.Text)));
				ForceSaveDocx(DocFileName);
				Console.WriteLine(string.Format("File {0} saved", DocFileName));
				return 0;
			});
		}
		public async System.Threading.Tasks.Task<int> PrintDocxStats(string DocFileName, bool Xpath, bool Sections, bool FootersHeaders)
		{
			return await System.Threading.Tasks.Task.Run(() => {
				Document doc;
				StartWordApp();
				Console.WriteLine(string.Format("Loading file {0}", DocFileName));
				try {
					doc = WordApp.Documents.Open(FileName: DocFileName, ReadOnly: true);
				} catch (Exception e) {
					Console.WriteLine(string.Format("Error loading File {0}: {1}", DocFileName, e.Message));
					return 1;
				}
				Console.WriteLine(string.Format("File {0} loaded", DocFileName));
				if (Xpath) {
					foreach (ContentControl Tag in doc.ContentControls) {
						if (Tag.Range.ContentControls.Count == 0) {
							TagParents = "";
							TagParents = Tag.Tag + GetParentTagList(Tag);
							var TagParentsArr = TagParents.Split('<');
							Array.Reverse(TagParentsArr);
							TagParents = String.Join(">", TagParentsArr);
							string tagRangeText;
							try {
								tagRangeText = Tag.Range.Text.Contains("\r") || Tag.Range.Text.Contains("\n") ? Regex.Escape(Tag.Range.Text) : Tag.Range.Text;
							} catch (NullReferenceException) {
								tagRangeText = "";
							} 
							tagRangeText = tagRangeText.Contains(";") ? string.Format("\"{0}\"", tagRangeText) : tagRangeText;
							string tagTitle;
							try {
								tagTitle = Tag.Title.Contains(";") ? string.Format("\"{0}\"", Tag.Title) : Tag.Title;
							} catch (Exception) {
								tagTitle = "";
							}	
							Console.WriteLine(string.Format("Tag;{0};{1};{2};{3}", tagTitle, tagRangeText, DocFileName, TagParents));
						}
					}
				}
				if (Sections) {
					Console.WriteLine(string.Format("Number of sections in {0}): {1}", DocFileName, doc.Sections.Count));
				}
				if (FootersHeaders) {
					int i = 1;
					foreach (Section sec in doc.Sections) {
						for (var j = 1; j < 3; j++) {
							WdHeaderFooterIndex hdrType = (WdHeaderFooterIndex)j;
							HeaderFooter hdrFtr;
							hdrFtr = sec.Headers[hdrType];
							Console.WriteLine(string.Format("{5} {4} header in section {0} (Words: {1}, Paragraphs: {2})Text: {3}", 
								i, hdrFtr.Range.Words.Count, hdrFtr.Range.Paragraphs.Count, Regex.Escape(hdrFtr.Range.Text), hdrType, Path.GetFileName(DocFileName)));
							hdrFtr = sec.Headers[hdrType];
							hdrFtr = sec.Footers[hdrType];
							Console.WriteLine(string.Format("{5} {4} footer in section {0} (Words: {1}, Paragraphs: {2})Text: {3}", 
								i, hdrFtr.Range.Words.Count, hdrFtr.Range.Paragraphs.Count, Regex.Escape(hdrFtr.Range.Text), hdrType, Path.GetFileName(DocFileName)));
						}
						i++;
					}
				}
				doc.Close(SaveChanges: false);
				return 0;                                       	
			});

		}
		public async System.Threading.Tasks.Task<int> XpathReplaceExact(string FileName, string OldText, string NewText)
		{
			return await System.Threading.Tasks.Task.Run(() => {
				StartWordApp();
				Console.WriteLine(string.Format("Loading document {0}", FileName));
				Document doc;
				try {
					doc = WordApp.Documents.Open(FileName: FileName);
				} catch (Exception e) {
					Console.WriteLine(string.Format("Error loading File {0}: {1}", FileName, e.Message));
					return 0;
				}
				Console.WriteLine(string.Format("Document {0} loaded", FileName));
				bool changed = false;
				string BeforeChange;
				foreach (ContentControl Tag in doc.ContentControls) {
					if (Tag.Range.ContentControls.Count == 0) {
						try {
							if (Tag.Range.Text == OldText) {
								BeforeChange = Tag.Range.Text;
								Tag.Range.Text = NewText;
								changed = true;
								Console.WriteLine(string.Format("Replaced {0} with {1}, document: {2}", BeforeChange, Tag.Range.Text, FileName));
							}
						} catch (NullReferenceException) {
							Console.WriteLine(string.Format("Empty tag ID: {0}, Title: {1}, Document: {2}", Tag.ID, Tag.Title, FileName));
						}
						if (Tag.Title == OldText) {
							BeforeChange = Tag.Title;
							Tag.Title = NewText;
							changed = true;
							Console.WriteLine(string.Format("Replaced {0} with {1}, document: {2}", BeforeChange, Tag.Title, FileName));
						}
					}
				}
				Console.WriteLine(string.Format("Closing {0}, saved: {1}", FileName, changed));
				if (changed) {
					ForceSaveDocx(FileName);
				} else {
					doc.Close(SaveChanges: false);
				}
				return 0;
			});
		}
		public async System.Threading.Tasks.Task<int> XpathReplaceFragment(string FileName, string OldText, string NewText)
		{
			return await System.Threading.Tasks.Task.Run(() => {
				StartWordApp();
				Console.WriteLine(string.Format("Loading document {0}", FileName));
				Document doc;
				try {
					doc = WordApp.Documents.Open(FileName: FileName);
				} catch (Exception e) {
					Console.WriteLine(string.Format("Error loading File {0}: {1}", FileName, e.Message));
					return 0;
				}
				Console.WriteLine(string.Format("Document {0} loaded", FileName));
				bool changed = false;
				string BeforeChange;
				foreach (ContentControl Tag in doc.ContentControls) {
					if (Tag.Range.ContentControls.Count == 0) {
						try {
							if (Tag.Range.Text.Contains(OldText)) {
								BeforeChange = Tag.Range.Text;
								Tag.Range.Text = Tag.Range.Text.Replace(OldText, NewText);
								changed = true;
								Console.WriteLine(string.Format("Replaced {0} with {1}, document: {2}", BeforeChange, Tag.Range.Text, FileName));
							}
						} catch (NullReferenceException) {
							Console.WriteLine(string.Format("Skippinkg possibly empty tag ID: {0}, Title: {1}, Document: {2}", Tag.ID, Tag.Title, FileName));
						}
						if (Tag.Title.Contains(OldText)) {
							BeforeChange = Tag.Title;
							Tag.Title = Tag.Title.Replace(OldText, NewText);
							changed = true;
							Console.WriteLine(string.Format("Replaced {0} with {1}, document: {2}", BeforeChange, Tag.Title, FileName));
						}
					}
				}
				Console.WriteLine(string.Format("Closing {0}, saved: {1}", FileName, changed));
				if (changed) {
					ForceSaveDocx(FileName);
				} else {
					doc.Close(SaveChanges: false);
				}
				return 0;
			});
		}
		public void ForceSaveDocx(string FileName)
		{
			string TempFileName = string.Format("{0}{1}{2}.docx", Environment.GetEnvironmentVariable("TEMP"), "\\", Guid.NewGuid().ToString());
			Document doc = WordApp.Documents[FileName];
			doc.SaveAs(FileName: TempFileName);
			doc.Close(SaveChanges: false);
			File.Copy(TempFileName, FileName, true);
			File.Delete(TempFileName);
		}
		public async System.Threading.Tasks.Task<int> TextReplace(string FileName, string OldText, string NewText)
		{
			return await System.Threading.Tasks.Task.Run(() => {
				StartWordApp();
				Document doc;
				try {
					Console.WriteLine(string.Format("Loading document {0}", FileName));
					doc = WordApp.Documents.Open(FileName: FileName);
				} catch (Exception e) {
					Console.WriteLine(string.Format("Error loading File {0}: {1}", FileName, e.Message));
					return 0;
				}			                                             	
				Console.WriteLine(string.Format("Document {0} loaded", FileName));
				doc.ToggleFormsDesign();
				bool changed = doc.Content.Find.Execute(FindText: OldText, MatchCase: true, MatchWholeWord: true, ReplaceWith: NewText, Replace: WdReplace.wdReplaceAll);
				if (changed) {
					Console.WriteLine(string.Format("No changes in document: {0}", FileName));
				} else {
					Console.WriteLine(string.Format("Chanes made in {0}, document saved", FileName));
				}
				if (changed) {
					ForceSaveDocx(FileName);
				} else {
					doc.Close(SaveChanges: false);
				}
				return 0;
			});
		}
		private string GetParentTagList(ContentControl Tag)
		{
			if (Tag.ParentContentControl != null) {
				TagParents += "<" + Tag.ParentContentControl.Tag;
				GetParentTagList(Tag.ParentContentControl);
			}
			return TagParents;
		}
	}
}
