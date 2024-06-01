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
        

        public WordDocument() 
        {
            app = new Word.Application();
        }

        public WordDocument(string fileName)
        {
            app = new Word.Application();
        }
        public void Dispose()
        {
            app.Quit();
            if (app != null)
            {
                Marshal.FinalReleaseComObject(app);
            }

            app = null;
        }
    }
}
