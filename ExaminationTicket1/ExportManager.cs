using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace ExaminationTicket1
{
    public class ExportManager
    {
        public async Task ToWordAsync(string patternWordFileName, Dictionary<string, string> data, string newFileName)
        {
            FileInfo fileInfo = null;
            Word.Application app = null;

            if (File.Exists(patternWordFileName))
            {
                fileInfo = new FileInfo(patternWordFileName);
            }
            else
            {
                throw new Exception("Шаблон файла не найден");
            }
            await Task.Run(() =>
            {
                try
                {
                    app = new Word.Application();
                    var file = fileInfo.FullName;
                    app.Documents.Open(file);

                    foreach (var item in data)
                    {
                        Word.Find find = app.Selection.Find;
                        find.Text = item.Key;
                        find.Replacement.Text = item.Value;

                        var wrap = Word.WdFindWrap.wdFindContinue;
                        var replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing,
                            MatchCase: false,
                            MatchWildcards: false,
                            MatchSoundsLike: false,
                            MatchAllWordForms: Type.Missing,
                            Forward: true,
                            Wrap: wrap,
                            Format: false,
                            Replace: replace);
                    }
                    Thread.Sleep(10000);
                    app.ActiveDocument.SaveAs2(newFileName);
                    app.ActiveDocument.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (app != null)
                    {
                        if (app.ActiveDocument != null)
                            app.ActiveDocument.Close();
                        app.Quit();
                    }
                }
            });
        }
    }
}
