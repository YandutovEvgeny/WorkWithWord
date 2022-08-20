using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace Word
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Document doc = null;
            try
            {
                string path = "c:\\Users\\Admin\\Desktop\\WordC#.docx";
                Microsoft.Office.Interop.Word.Application application =
                    new Microsoft.Office.Interop.Word.Application();
                doc = application.Documents.Open(path);

                doc.Activate();
                Bookmarks bookmarks = doc.Bookmarks;
                int n = bookmarks.Count;
                List<string> text = new List<string>();
                text.Add(textBox2.Text);
                text.Add(monthCalendar1.SelectionStart.ToString());
                text.Add(textBox1.Text);
                int i = 0;
                Range range;
                foreach (Bookmark bookmark in bookmarks)
                {
                    range = bookmark.Range;
                    range.Text = text[i++];
                }   
                doc.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                doc.Close();
            }
        }
    }
}
