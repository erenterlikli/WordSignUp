using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using wordgetir = Microsoft.Office.Interop.Word;
using System.Reflection;
namespace WordKayıt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object omissing = System.Reflection.Missing.Value;
            object dokumansonu = "\\endofdoc";
            wordgetir.Application olustur;
            wordgetir.Document icerik;

            olustur = new wordgetir.Application();
            olustur.Visible = true;
            icerik = olustur.Documents.Add(ref omissing);

            wordgetir.Paragraph ad;
            ad = icerik.Content.Paragraphs.Add(ref omissing);
            ad.Range.Text = "Ad: " +textBox1.Text;
            ad.Range.Font.Bold = 1;
            ad.Format.SpaceAfter = 10;
            ad.Range.InsertParagraphAfter();

            wordgetir.Paragraph soyad;
            object ob = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            soyad = icerik.Content.Paragraphs.Add(ref omissing);
            soyad.Range.Text = "Soyad: " + textBox2.Text;
            soyad.Range.Font.Bold = 1;
            soyad.Format.SpaceAfter = 10;
            soyad.Range.InsertParagraphAfter();

            wordgetir.Paragraph yas;
            ob = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            yas = icerik.Content.Paragraphs.Add(ref omissing);
            yas.Range.Text = "Yaş: " + textBox3.Text;
            yas.Range.Font.Bold = 1;
            yas.Format.SpaceAfter = 10;
            yas.Range.InsertParagraphAfter();

            wordgetir.Paragraph sehir;
            ob = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            sehir = icerik.Content.Paragraphs.Add(ref omissing);
            sehir.Range.Text = "Şehir: " + comboBox1.Text;
            sehir.Range.Font.Bold = 1;
            sehir.Format.SpaceAfter = 10;
            sehir.Range.InsertParagraphAfter();

            wordgetir.Paragraph tel;
            ob = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            tel = icerik.Content.Paragraphs.Add(ref omissing);
            tel.Range.Text = "Telefon: " + textBox4.Text;
            tel.Range.Font.Bold = 1;
            tel.Format.SpaceAfter = 10;
            tel.Range.InsertParagraphAfter();

            wordgetir.Paragraph meslek;
            ob = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            meslek = icerik.Content.Paragraphs.Add(ref omissing);
            meslek.Range.Text = "Meslek: " + textBox5.Text;
            meslek.Range.Font.Bold = 1;
            meslek.Format.SpaceAfter = 10;
            meslek.Range.InsertParagraphAfter();

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            comboBox1.Text = " ";

        
        }
    }
}
