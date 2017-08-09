using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Reflection;






namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Application appl = new Application();
            Object templatePathObj = Environment.CurrentDirectory + "\\Shablom.dot";
            Microsoft.Office.Interop.Word.Document doc = appl.Documents.Add(templatePathObj);

            doc.Bookmarks["TYPE"].Range.Text = "ТВВ-320";
            doc.Bookmarks["POWER"].Range.Text = "320";
            doc.Bookmarks["VOLTAGE"].Range.Text = "20";
            doc.Bookmarks["IMAGE"].Range.InlineShapes.AddPicture(Environment.CurrentDirectory + "\\Koala.jpg");
            doc.SaveAs(FileName: Environment.CurrentDirectory + "\\For_print.docx");
            appl.Quit();



        }
    }
}

