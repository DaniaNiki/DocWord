using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace DocWord
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            Document doc = app.Documents.Add(Visible: true);
            Microsoft.Office.Interop.Word.Range r = doc.Range();
            r.Text = "Nikitin D";
            //r.Bold = 20;
            Table t = doc.Tables.Add(r, 5, 5);
            t.Borders.Enable = 1;
            foreach(Row row in t.Rows)
            {
                foreach(Cell cell in row.Cells)
                {
                    if(cell.RowIndex == 1)
                    {
                        cell.Range.Text = "Row" + cell.ColumnIndex.ToString();
                        cell.Range.Bold = 1;
                        cell.Range.Font.Name = "Times new roman";
                        cell.Range.Font.Size = 14;

                        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                        cell.Range.Text = "Hi!";
                    }
                }
            }
            doc.Save();
            app.Documents.Open(@"C:\Users\Даня\Desktop\doc1.docx");
            Console.ReadKey();
            try
            { 
                doc.Close();
                app.Quit(); 
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadKey();

        }
    }
}
