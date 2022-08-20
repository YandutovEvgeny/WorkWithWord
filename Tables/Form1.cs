using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Tables
{
    public partial class Form1 : Form
    {
        Word.Document doc; 
        Word.Application app = new Word.Application();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                doc = app.Documents.Open("C:\\Users\\Admin\\Desktop\\TablesC#.docx");
                doc.Activate();
                Word.Table table = doc.Tables[1];
                //Range - то место куда мы ставим курсор мыши
                //Word.Range range = table.Cell(2, 2).Range;
                //range.Text = "Привет мир!";
                //table.Rows.Add(range);
                Word.Range range = table.Cell(1, 1).Range;
                for (int i = 0; i < 8; i++)
                {
                    table.Rows.Add(range);
                    range = table.Cell(i + 2, 1).Range;
                }

                for (int i = 1; i <= 9; i++)
                {
                    for (int j = 1; j <= 9; j++)
                    {
                        range = table.Cell(i, j).Range;
                        range.Text = $"{i} * {j} = {i * j}";
                    }
                }
            }
            catch(Exception ex)
            {
                //MessageBox.Show(ex.Message);
                doc.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object obj = System.Reflection.Missing.Value;
            doc = app.Documents.Add(obj,obj,obj,obj);
            Word.Range range = doc.Range(obj, obj);
            Word.Table table = doc.Tables.Add(range, 9, 9, obj,obj);
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            for (int i = 1; i <= 9; i++)
            {
                for (int j = 1; j <= 9; j++)
                {
                    range = table.Cell(i, j).Range;
                    if (i == 1 || j == 1)
                    {
                        range.Font.Color = Word.WdColor.wdColorRed;
                        table.Cell(i, j).Shading.BackgroundPatternColor = Word.WdColor.wdColorLightYellow;
                        
                    }
                    range.Text = $"{i} * {j} = {i * j}";
                }
            }
            doc.Activate();
            doc.SaveAs2("C:\\1\\2.docx");
            doc.Close();
        }
    }
}
