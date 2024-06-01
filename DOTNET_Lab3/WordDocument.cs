using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace DOTNET_Lab3
{
    public class WordDocument : IDisposable
    {
        Word.Application? app = null;
        Word.Document? doc = null;
        //Word.Paragraph? p = null;
        

        public WordDocument() 
        {
            app = new Word.Application();
            doc = app.Documents.Add();
        }

        public WordDocument(string fileName)
        {
            app = new Word.Application();
            doc = app.Documents.Open(fileName);
        }

        public void SaveAs(string fileName)
        {
            doc?.SaveAs(fileName);
        }

        public string? this[int row]
        {
            get => doc?.Paragraphs[row].Range.Text.ToString();
            set
            {
                //if (p is not null)
                    doc.Paragraphs[row].Range.Text = value;
            }
        }
        public void AddParagraph(string text, int fontSize, Word.WdColor fontColor, Word.WdParagraphAlignment alignment)
        {
            var para = doc.Paragraphs.Add();
            
            para.Range.Font.Size = fontSize;
            para.Range.Font.Color = fontColor;           
            para.Range.Text = text;
            para.Alignment = alignment;

            para.Range.InsertParagraphAfter();
        }

        public void Dispose()
        {
            doc?.Close();
            app.Quit();
            Release(app); app = null;
            Release(doc); doc = null;
            //Release(p); p = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void Release(object obj)
        {
            if (obj != null)
            {
                Marshal.FinalReleaseComObject(obj);
            }
        }
    }
}
