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


namespace Marriage_certificate
{
    public partial class Form1 : Form
    {
        Word.Document doc;
        Word.Application app;
        public Form1()
        {
            InitializeComponent();
            string path = "C:\\1\\Свидетельство о заключении брака.docx";
            app = new Word.Application();
            doc = app.Documents.Open(path);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                doc.Activate();
                Word.Bookmarks bookmarks = doc.Bookmarks;
                List<string> insertValues = new List<string>();
                insertValues.Add(textBox1.Text);    //Фамилия с кем заключают a
                insertValues.Add(textBox2.Text);    //Имя с кем заключают b
                insertValues.Add(textBox3.Text);    //Отчество с кем заключают c
                insertValues.Add(monthCalendar1.SelectionStart.Day.ToString());  //День рождения с кем заключают d 
                insertValues.Add(monthCalendar1.SelectionStart.Month.ToString()); //Месяц рождения с кем заключают e 
                insertValues.Add(monthCalendar1.SelectionStart.Year.ToString()); //Год рождения с кем заключают f

                insertValues.Add(textBox4.Text);    //Отчество кто заключает g
                insertValues.Add(textBox5.Text);    //Имя кто заключает h 
                insertValues.Add(textBox6.Text);    //Фамилия кто заключает i              
                insertValues.Add(monthCalendar2.SelectionStart.Day.ToString()); //день рождения кто заключает j
                insertValues.Add(monthCalendar2.SelectionStart.Month.ToString()); //месяц рождения кто заключает k
                insertValues.Add(monthCalendar2.SelectionStart.Year.ToString());    //год рождения кто заключает l

                insertValues.Add(textBox7.Text);    //Кол-во детей не достигших совершеннолетия m
                insertValues.Add(textBox8.Text);    //Фамилия после брака n 

                insertValues.Add(monthCalendar3.SelectionStart.Day.ToString()); //День заключения договора o
                insertValues.Add(monthCalendar3.SelectionStart.Month.ToString()); //Месяц заключения договора p
                insertValues.Add(monthCalendar3.SelectionStart.Year.ToString());  //Год заключения договора q

                insertValues.Add(textBox9.Text);    //Подпись r

                int i = 0;
                Word.Range range;
                foreach(Word.Bookmark bookmark in bookmarks)
                {
                    range = bookmark.Range;
                    range.Text = insertValues[i++];
                }
                doc.Close();
            }
            catch (Exception)
            {

                doc.Close();
            }
        }
    }
}
