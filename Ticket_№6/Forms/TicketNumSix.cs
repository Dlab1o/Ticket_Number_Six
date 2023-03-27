using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace Ticket__6
{
    public partial class TicketNumSix : Form
    {
        public TicketNumSix()
        {
            InitializeComponent();
        }

        //По нажатию на чекбокс (Фотопечать) появляется место для картинки
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
            {
                pictureBox4.Visible = false;
                label2.Visible = false;
            }
            else
            {
                pictureBox4.Visible = true;
                label2.Visible = true;
            }
        }

        private void TicketNumSix_Load(object sender, EventArgs e)
        {

        }

        // Добавление картинки
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            OpenFileDialog opFile = new OpenFileDialog();
            opFile.Title = "Выбирите изображение услуги...";
            opFile.Filter = "All files (*.*)|*.*|png files (*.png)|*.png";

            string appPath = Path.GetDirectoryName(Application.ExecutablePath) + @"\ServicesImages\";
            if (Directory.Exists(appPath) == false)
            {
                Directory.CreateDirectory(appPath);
            }

            if (opFile.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string iName = opFile.SafeFileName;
                    string filepath = opFile.FileName;
                    File.Copy(filepath, appPath + iName);
                    pictureBox4.Image = new Bitmap(opFile.OpenFile());
                    Classes.PriceFunct.ImagePath = iName;
                    Classes.PriceFunct.IsUpdate = true;
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Невозможно открыть файл!\n" + exp.Message);
                }
            }
            else
            {
                opFile.Dispose();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Создаём объект документа
            try
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                {
                    return;
                }
                string filename = saveFileDialog1.FileName;

                var time = DateTime.Now;

                Word.Document doc = null;
                // Создаём объект приложения
                Word.Application app = new Word.Application();
                // Путь до шаблона документа
                string source = Path.Combine(Directory.GetCurrentDirectory(), "Чек.docx");
                // Открываем
                doc = app.Documents.Open(source);
                doc.Activate();

                // Добавляем информацию
                // Bookmarks содержит все закладки
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
                Random r = new Random();

                string[] data = new string[3] { time.ToString(), Classes.BillEdit.Price.ToString(), Classes.BillEdit.ID.ToString()};


                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }
                doc.SaveAs2(filename + ".docx");


                doc.Close();
                MessageBox.Show($@"Чек был успешно сформирован!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show($@"{ex.Message}", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            panel1.Capture = false;
            Message m = Message.Create(Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            WndProc(ref m);
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            label1.Capture = false;
            Message m = Message.Create(Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            WndProc(ref m);
        }

        private void pictureBox2_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox2.Capture = false;
            Message m = Message.Create(Handle, 0xa1, new IntPtr(2), IntPtr.Zero);
            WndProc(ref m);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Classes.PriceFunct.CheckAllSpace(textBox1.Text) == true)
            {
                MessageBox.Show("Пожалуйста, укажите ширину потолка!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                if (Classes.PriceFunct.CheckAllSpace(textBox2.Text) == true)
                {
                    MessageBox.Show("Пожалуйста, укажите длинну потолка!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    if (checkBox2.Checked == true && Classes.PriceFunct.IsUpdate == false)
                    { 
                        MessageBox.Show("Пожалуйста, загрузите фотографию потолка!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    else
                    {
                        if (radioButton1.Checked == true)
                        {
                            if (checkBox1.Checked == false && checkBox2.Checked == false)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double Price = Convert.ToDouble(SquareMeters) * Math.Round(213.15, 2);


                                label6.Text = ($"Цена за кв.м: 213.15\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 213.15 = {Price}\nИтоговая стоимость: {Math.Round(Price, 2)} руб.");

                                Random rnd = new Random();

                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = Math.Round(Price, 2);
                                Classes.PriceFunct.PrintVariant = 1;
                            }
                            else if (checkBox1.Checked == true && checkBox2.Checked == false)
                            {
                                
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(213.15, 2);
                                double Price = Math.Round(TPrice, 2);
                                double TDoublePrice = Price * 30 / 100;
                                double DoublePrice = Math.Round(TDoublePrice, 2);
                                double TEndPrice = Price + DoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 213.15\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 213.15 = {Price}\nДобавочная стоимость за многоуровневость: 30% ({DoublePrice})\nИтоговая стоимость: {EndPrice} руб.");

                                Random rnd = new Random();

                                Classes.BillEdit.ID = rnd.Next();

                                Classes.PriceFunct.PrintVariant = 2;
                            }
                            else if (checkBox1.Checked == false && checkBox2.Checked == true)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(213.15, 2);
                                double Price = Math.Round(TPrice, 2);
                                double TDoublePrice = Price * 26 / 100;
                                double DoublePrice = Math.Round(TDoublePrice, 2);
                                double TEndPrice = Price + DoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 213.15\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 213.15 = {Price}\nДобавочная стоимость за фотопечать: 26% ({DoublePrice})\nИтоговая стоимость: {EndPrice} руб.");

                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = EndPrice;
                                Classes.PriceFunct.PrintVariant = 3;
                            }
                            else if (checkBox1.Checked == true && checkBox2.Checked == true)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(213.15, 2);
                                double Price = Math.Round(TPrice, 2);
                                double PDoublePrice = Price * 26 / 100;
                                double LDoublePrice = Price * 30 / 100;
                                double DoublePrice = Math.Round(LDoublePrice, 2);
                                double SecDoublePrice = Math.Round(PDoublePrice, 2);

                                double TEndPrice = Price + DoublePrice + SecDoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 213.15\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 213.15 = {Price}\nДобавочная стоимость за многоуровневость: 30% ({DoublePrice})\nДобавочная стоимость за фотопечать: 26% ({SecDoublePrice})\nИтоговая стоимость: {EndPrice} руб.");

                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = EndPrice;
                                Classes.PriceFunct.PrintVariant = 4;
                            }

                        }
                        else if (radioButton2.Checked == true)
                        {
                            if (checkBox1.Checked == false && checkBox2.Checked == false)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double Price = Convert.ToDouble(SquareMeters) * Math.Round(265.8, 2);
                                Math.Round(Price, 2);


                                label6.Text = ($"Цена за кв.м: 265.8\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 265.8 = {Price}\nИтоговая стоимость: {Price} руб.");

                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = Math.Round(Price, 2);

                                Classes.PriceFunct.PrintVariant = 5;
                            }
                            else if (checkBox1.Checked == true && checkBox2.Checked == false)
                            {

                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(265.8, 2);
                                double Price = Math.Round(TPrice, 2);
                                double TDoublePrice = Price * 30 / 100;
                                double DoublePrice = Math.Round(TDoublePrice, 2);
                                double TEndPrice = Price + DoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 265.8\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 265.8 = {Price}\nДобавочная стоимость за многоуровневость: 30% ({DoublePrice})\nИтоговая стоимость: {EndPrice} руб.");
                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = EndPrice;
                                Classes.PriceFunct.PrintVariant = 6;
                            }
                            else if (checkBox1.Checked == false && checkBox2.Checked == true)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(265.8, 2);
                                double Price = Math.Round(TPrice, 2);
                                double TDoublePrice = Price * 26 / 100;
                                double DoublePrice = Math.Round(TDoublePrice, 2);
                                double TEndPrice = Price + DoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 265.8\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 265.8 = {Price}\nДобавочная стоимость за фотопечать: 26% ({DoublePrice})\nИтоговая стоимость: {EndPrice} руб.");

                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = EndPrice;
                                Classes.PriceFunct.PrintVariant = 7;
                            }
                            else if (checkBox1.Checked == true && checkBox2.Checked == true)
                            {
                                int Width = Convert.ToInt32(textBox1.Text);
                                int Height = Convert.ToInt32(textBox2.Text);
                                int SquareMeters = Width * Height;

                                double TPrice = Convert.ToDouble(SquareMeters) * Math.Round(265.8, 2);
                                double Price = Math.Round(TPrice, 2);
                                double PDoublePrice = Price * 26 / 100;
                                double LDoublePrice = Price * 30 / 100;
                                double DoublePrice = Math.Round(LDoublePrice, 2);
                                double SecDoublePrice = Math.Round(PDoublePrice, 2);

                                double TEndPrice = Price + DoublePrice + SecDoublePrice;
                                double EndPrice = Math.Round(TEndPrice, 2);

                                label6.Text = ($"Цена за кв.м: 265.8\n{Width} м. * {Height} м. = {SquareMeters} кв.м.\n{SquareMeters} кв.м. * 265.8 = {Price}\nДобавочная стоимость за многоуровневость: 30% ({DoublePrice})\nДобавочная стоимость за фотопечать: 26% ({SecDoublePrice})\nИтоговая стоимость: {EndPrice} руб.");
                                
                                Random rnd = new Random();
                                Classes.BillEdit.ID = rnd.Next();
                                Classes.BillEdit.Price = EndPrice;
                                Classes.PriceFunct.PrintVariant = 8;
                            }
                        }
                    }
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0') && (e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
            {
                return;
            }
            e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0') && (e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
            {
                return;
            }
            e.Handled = true;
        }
    }
}
