using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// using Word = Microsoft.Office.Interop.Word;

namespace IncludeImage
{
    class Program
    {
        static void Main(string[] args)
        {   
            // Word = Microsoft.Office.Interop.Word
            // Microsoft.Office.Interop.Word.Application app = new .Application();
            // using Word = Microsoft.Office.Interop.Word;

            object missing = System.Reflection.Missing.Value;
            object name = args[0];
            object Range = System.Reflection.Missing.Value;

            //Start Word and create a new document.  
            Microsoft.Office.Interop.Word.Application oWord;
            Microsoft.Office.Interop.Word.Document oDoc;
            oWord = new Microsoft.Office.Interop.Word.Application();
            oDoc = oWord.Documents.Open(ref name, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            //bool t = oWord.Selection.Find.Execute("You Can");  
            object start = 0;
            object end = 1;
            Microsoft.Office.Interop.Word.Range range = oDoc.Range(ref start, ref end);

            // Microsoft.Office.Interop.Word.Range range = oDoc.Paragraphs[3].Range;
            //range.SetRange(range.Start, range.End - 10);

            // Microsoft.Office.Interop.Word.InlineShape autoScaledInlineShape = range.InlineShapes.AddOLEObject(FileName: @"C:\Users\Al\Desktop\gg.jpg");
            Microsoft.Office.Interop.Word.InlineShape autoScaledInlineShape = range.InlineShapes.AddPicture(FileName: args[1]);
            autoScaledInlineShape.Width = 100;
            autoScaledInlineShape.Height = 200;

            oDoc.Save();
            oDoc.Close();
            oWord.Quit();
            oDoc = null;
            oWord = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
