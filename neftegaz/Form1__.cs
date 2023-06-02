using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace neftegaz
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            readExcel();
        }

        private void readExcel()
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);

            //ИНФОРМАЦИЯ

            //ИТОГО (ГАЗ)
            ws = wb.Worksheets[1];
            label2.Text = (ws.Cells[2, 1].Text);
            label3.Text = (ws.Cells[2, 2].Text);
            label7.Text = (ws.Cells[2, 3].Text);
            label4.Text = (ws.Cells[3, 3].Text);
            label5.Text = (ws.Cells[3, 4].Text);
            label6.Text = (ws.Cells[3, 5].Text);
            string a = "";
            string b1 = "";
            string b2 = "";
            string b3 = "";
            string c = "";
            for (int i = 4; i < 15; i++)
            {
                a = a + ws.Cells[i, 1].Text;
                a += "\n";
                a += "    -----------------------------------    ";
                a += "\n";

                b1 += ws.Cells[i, 3].Text;
                b1 += "\n";
                b1 += "    -----    ";
                b1 += "\n";

                b2 += ws.Cells[i, 4].Text;
                b2 += "\n";
                b2 += "    -----    ";
                b2 += "\n";

                b3 += ws.Cells[i, 5].Text;
                b3 += "\n";
                b3 += "    -----    ";
                b3 += "\n";

                c += ws.Cells[i, 2].Text;
                c += "\n";
                c += "    -----    ";
                c += "\n";
            }

            richTextBox1.Text = a;
            richTextBox2.Text = b1;
            richTextBox3.Text = b2;
            richTextBox4.Text = b3;
            richTextBox5.Text = c;

            //ws.Cells[30, 30] = "Кто этот нефтегаз ваш придумал";
            //MessageBox.Show(ws.Cells[30, 30].Text);



        }

        private void button2_Click(object sender, EventArgs e)
        {
            string a = "Технологические потери природного газа" + "\n" + "тут короче копипаст из методички";
            MessageBox.Show(a);
        }

        private void readExcel2() //Итого конденсат
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[2];

            label9.Text = (ws.Cells[2, 1].Text);
            label10.Text = (ws.Cells[2, 2].Text);
            label11.Text = (ws.Cells[2, 3].Text);
            label12.Text = (ws.Cells[3, 3].Text);
            label13.Text = (ws.Cells[3, 4].Text);
            label14.Text = (ws.Cells[3, 5].Text);
            string aa = "";
            string bb1 = "";
            string bb2 = "";
            string bb3 = "";
            string cc = "";
            for (int i = 4; i < 13; i++)
            {
                aa += ws.Cells[i, 1].Text;
                aa += "\n";
                aa += "    -----------------------------------    ";
                aa += "\n";

                bb1 += ws.Cells[i, 3].Text;
                bb1 += "\n";
                bb1 += "    -----    ";
                bb1 += "\n";

                bb2 += ws.Cells[i, 4].Text;
                bb2 += "\n";
                bb2 += "    -----    ";
                bb2 += "\n";

                bb3 += ws.Cells[i, 5].Text;
                bb3 += "\n";
                bb3 += "    -----    ";
                bb3 += "\n";

                cc += ws.Cells[i, 2].Text;
                cc += "\n";
                cc += "    -----    ";
                cc += "\n";
            }
            richTextBox6.Text = aa;
            richTextBox7.Text = bb1;
            richTextBox8.Text = bb2;
            richTextBox9.Text = bb3;
            richTextBox10.Text = cc;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            readExcel2();
        }

        private void readExcel3() //гди + гки
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[3];
            string a3 = "";
            string b3 = "";
            string c3 = "";
            string d3 = "";
            string e3 = "";
            string f3 = "";

            for (int i = 5; i <= 13; i++)
            {
                a3 += (" " + ws.Cells[i, 1].Text);
                a3 += "\n";
            }
            for (int i = 5; i <= 13; i++)
            {
                b3 += (" " + ws.Cells[i, 2].Text);
                b3 += "\n";
            }
            for (int i = 5; i <= 13; i++)
            {
                c3 += (" " + ws.Cells[i, 3].Text);
                c3 += "\n";
            }
            for (int i = 5; i <= 13; i++)
            {
                d3 += (" " + ws.Cells[i, 4].Text);
                d3 += "\n";
            }
            for (int i = 5; i <= 13; i++)
            {
                e3 += (" " + ws.Cells[i, 4].Text);
                e3 += "\n";
            }
            for (int i = 18; i <= 26; i++)
            {
                f3 += (" " + ws.Cells[i, 5].Text);
                f3 += "\n";
            }
            richTextBox18.Text = a3;
            richTextBox27.Text = a3;
            richTextBox19.Text = b3;
            richTextBox26.Text = b3;
            richTextBox20.Text = c3;
            richTextBox25.Text = c3;
            richTextBox21.Text = d3;
            richTextBox24.Text = e3;
            richTextBox22.Text = f3;
            richTextBox23.Text = f3;
            label24.Text = ws.Cells[4, 1].Text;
            label25.Text = ws.Cells[4, 1].Text;
            label26.Text = ws.Cells[4, 2].Text;
            label27.Text = ws.Cells[4, 2].Text;
            label28.Text = ws.Cells[4, 3].Text;
            label29.Text = ws.Cells[4, 3].Text;
            label30.Text = ws.Cells[4, 4].Text;
            label31.Text = ws.Cells[4, 4].Text;
            label32.Text = ws.Cells[4, 5].Text;
            label33.Text = ws.Cells[4, 5].Text;
        }
        private void readExcel4() //опорож (шлейфы)
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[4];
        }
        private void readExcel5() //опорож (оборудование)
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[5];
        }
        private void readExcel6() //Дегазация жидкостей
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[6];
            string a6 = "";
            string b6 = "";
            string c6 = "";
            string d6 = "";
            string e6 = "";
            string f6 = "";
            string g6 = "";
            string h6 = "";
            string k6 = "";
            string t6 = "";
            for (int i = 5; i <= 13; i++)
            {
                a6 += (" " + ws.Cells[i, 1].Text);
                a6 += "\n";
                b6 += (" " + ws.Cells[i, 2].Text);
                b6 += "\n";
                c6 += (" " + ws.Cells[i, 3].Text);
                c6 += "\n";
                d6 += (" " + ws.Cells[i, 4].Text);
                d6 += "\n";
                e6 += (" " + ws.Cells[i, 5].Text);
                e6 += "\n";
            }
            for (int i = 18; i <= 26; i++)
            {
                f6 += (" " + ws.Cells[i, 2].Text);
                f6 += "\n";
                g6 += (" " + ws.Cells[i, 3].Text);
                g6 += "\n";
                h6 += (" " + ws.Cells[i, 4].Text);
                h6 += "\n";
                k6 += (" " + ws.Cells[i, 5].Text);
                k6 += "\n";
            }
            richTextBox28.Text = a6;
            richTextBox37.Text = a6;
            richTextBox29.Text = b6;
            richTextBox30.Text = c6;
            richTextBox31.Text = d6;
            richTextBox32.Text = e6;
            richTextBox36.Text = f6;
            richTextBox35.Text = g6;
            richTextBox34.Text = h6;
            richTextBox33.Text = k6;
            label34.Text = ws.Cells[3, 1].Text;
            label35.Text = ws.Cells[16, 1].Text;
            label36.Text = ws.Cells[3, 2].Text;
            label37.Text = ws.Cells[16, 2].Text;
            label38.Text = ws.Cells[3, 3].Text;
            label39.Text = ws.Cells[16, 3].Text;
            label40.Text = ws.Cells[3, 4].Text;
            label41.Text = ws.Cells[16, 4].Text;
            label42.Text = ws.Cells[4, 4].Text;
            label43.Text = ws.Cells[17, 4].Text;
            label44.Text = ws.Cells[3, 5].Text;
            label45.Text = ws.Cells[16, 5].Text;
            label46.Text = ws.Cells[1, 1].Text;
            label47.Text = ws.Cells[2, 1].Text;
            label48.Text = ws.Cells[15, 1].Text;

        }
        private void readExcel7() //Хим. реагенты
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];
            string a7 = "";
            string b7 = "";
            string c7 = "";
            string d7 = "";
            string e7 = "";
            for (int i = 4; i <= 11; i++)
            {
                a7 += (" " + ws.Cells[i, 1].Text);
                a7 += "\n";
                b7 += (" " + ws.Cells[i, 2].Text);
                b7 += "\n";
                c7 += (" " + ws.Cells[i, 3].Text);
                c7 += "\n";
                d7 += (" " + ws.Cells[i, 4].Text);
                d7 += "\n";
                e7 += (" " + ws.Cells[i, 5].Text);
                e7 += "\n";
            }
            richTextBox38.Text = a7;
            richTextBox39.Text = b7;
            richTextBox40.Text = c7;
            richTextBox41.Text = d7;
            richTextBox42.Text = e7;
            label49.Text = ws.Cells[3, 1].Text;
            label50.Text = ws.Cells[3, 2].Text;
            label51.Text = ws.Cells[3, 3].Text;
            label52.Text = ws.Cells[3, 4].Text;
            label53.Text = ws.Cells[3, 5].Text;
            label54.Text = ws.Cells[1, 1].Text;
        }
        private void readExcel8() //Отбор проб (газ)
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[8];
            string a8 = "";
            string b8 = "";
            string c8 = "";
            string d8 = "";
            string e8 = "";
            for (int i = 6; i <= 16; i++)
            {
                a8 += (" " + ws.Cells[i, 1].Text);
                a8 += "\n";
                b8 += (" " + ws.Cells[i, 2].Text);
                b8 += "\n";
                c8 += (" " + ws.Cells[i, 3].Text);
                c8 += "\n";
                d8 += (" " + ws.Cells[i, 4].Text);
                d8 += "\n";
                e8 += (" " + ws.Cells[i, 5].Text);
                e8 += "\n";
            }
            richTextBox43.Text = a8;
            richTextBox44.Text = b8;
            richTextBox45.Text = c8;
            richTextBox46.Text = d8;
            richTextBox47.Text = e8;
            label55.Text = ws.Cells[3, 1].Text;
            label56.Text = ws.Cells[3, 2].Text;
            label57.Text = ws.Cells[3, 3].Text;
            label58.Text = ws.Cells[3, 4].Text;
            label59.Text = ws.Cells[4, 4].Text;
            label60.Text = ws.Cells[3, 5].Text;
            label61.Text = ws.Cells[1, 1].Text;
            label62.Text = ws.Cells[5, 1].Text;
        }
        private void readExcel9() //Отбор проб (конденсат)
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[9];
            string a9 = "";
            string b9 = "";
            string c9 = "";
            string d9 = "";
            string e9 = "";
            for (int i = 6; i <= 11; i++)
            {
                a9 += (" " + ws.Cells[i, 1].Text);
                a9 += "\n";
                b9 += (" " + ws.Cells[i, 2].Text);
                b9 += "\n";
                c9 += (" " + ws.Cells[i, 3].Text);
                c9 += "\n";
                d9 += (" " + ws.Cells[i, 4].Text);
                d9 += "\n";
                e9 += (" " + ws.Cells[i, 5].Text);
                e9 += "\n";
            }
            richTextBox48.Text = a9;
            richTextBox49.Text = b9;
            richTextBox50.Text = c9;
            richTextBox51.Text = d9;
            richTextBox52.Text = e9;
            label63.Text = ws.Cells[3, 1].Text;
            label64.Text = ws.Cells[3, 2].Text;
            label65.Text = ws.Cells[3, 3].Text;
            label66.Text = ws.Cells[3, 4].Text;
            label67.Text = ws.Cells[4, 4].Text;
            label68.Text = ws.Cells[3, 5].Text;
            label69.Text = ws.Cells[1, 1].Text;
            label70.Text = ws.Cells[5, 1].Text;
        }
        private void readExcel10() //клапана
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[10];
            string a10 = "";
            string b10 = "";
            string c10 = "";
            string d10 = "";
            string e10 = "";
            for (int i = 4; i <= 14; i++)
            {
                a10 += (" " + ws.Cells[i, 1].Text);
                a10 += "\n";
                b10 += (" " + ws.Cells[i, 2].Text);
                b10 += "\n";
                c10 += (" " + ws.Cells[i, 3].Text);
                c10 += "\n";
                d10 += (" " + ws.Cells[i, 4].Text);
                d10 += "\n";
                e10 += (" " + ws.Cells[i, 5].Text);
                e10 += "\n";
            }
            richTextBox53.Text = a10;
            richTextBox54.Text = b10;
            richTextBox55.Text = c10;
            richTextBox56.Text = d10;
            richTextBox57.Text = e10;
            label71.Text = ws.Cells[3, 1].Text;
            label72.Text = ws.Cells[3, 2].Text;
            label73.Text = ws.Cells[3, 3].Text;
            label74.Text = ws.Cells[3, 4].Text;
            label75.Text = ws.Cells[3, 5].Text;
            label76.Text = ws.Cells[1, 1].Text;
        }
        private void readExcel11() //Унос с жидкостью
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[11];
            string a11 = "";
            string b11 = "";
            string c11 = "";
            string d11 = "";
            string e11 = "";
            for (int i = 4; i <= 11; i++)
            {
                a11 += (" " + ws.Cells[i, 1].Text);
                a11 += "\n";
                b11 += (" " + ws.Cells[i, 2].Text);
                b11 += "\n";
                c11 += (" " + ws.Cells[i, 3].Text);
                c11 += "\n";
                d11 += (" " + ws.Cells[i, 4].Text);
                d11 += "\n";
                e11 += (" " + ws.Cells[i, 5].Text);
                e11 += "\n";
            }
            richTextBox58.Text = a11;
            richTextBox59.Text = b11;
            richTextBox60.Text = c11;
            richTextBox61.Text = d11;
            richTextBox62.Text = e11;
            label77.Text = ws.Cells[3, 1].Text;
            label78.Text = ws.Cells[3, 2].Text;
            label79.Text = ws.Cells[3, 3].Text;
            label80.Text = ws.Cells[3, 4].Text;
            label81.Text = ws.Cells[3, 5].Text;
            label82.Text = ws.Cells[1, 1].Text;
        }
        private void readExcel12() //расчет плотности
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[12];
        }
        private void readExcel13() //Расчет Z
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[13];
            string a13 = "";
            string b13 = "";
            string c13 = "";
            string d13 = "";

            for (int i = 1; i <= 7; i++)
            {
                a13 += (" " + ws.Cells[i, 1].Text);
                a13 += "\n";
            }
            for (int i = 1; i <= 7; i++)
            {
                b13 += (" " + ws.Cells[i, 2].Text);
                b13 += "\n";
            }
            for (int i = 10; i <= 39; i++)
            {
                c13 += (" " + ws.Cells[i, 1].Text);
                c13 += "\n";
            }
            for (int i = 10; i <= 39; i++)
            {
                d13 += (" " + ws.Cells[i, 2].Text);
                d13 += "\n";
            }
            richTextBox14.Text = a13;
            richTextBox15.Text = b13;
            richTextBox16.Text = c13;
            richTextBox17.Text = d13;
        }

        private void readExcel14() //растворимость газа
        {
            string filePath = "C:\\file.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[14];
            string a14 = "";
            string b14 = "";
            string c14 = "";
            label20.Text = ws.Cells[1, 10].Text;
            label23.Text = ws.Cells[18, 2].Text;
            label21.Text = ws.Cells[17, 1].Text;
            label22.Text = ws.Cells[18, 1].Text;
            for (int i = 2; i <= 4; i++)
            {
                a14 += (" " + ws.Cells[i, 10].Text);
                a14 += "  |  ";
                a14 += ws.Cells[i, 11].Text;
                a14 += "\n";

            }
            a14 += ws.Cells[5, 10].Text + " |  " + ws.Cells[5, 11].Text;
            richTextBox11.Text = a14;
            c14 += ("  0" + "   |  " + ws.Cells[19, 2].Text + "\n");
            for (int i = 20; i <= 23; i++)
            {
                c14 += (" " + ws.Cells[i, 1].Text);
                c14 += "  | ";
                c14 += (" " + ws.Cells[i, 2].Text);
                c14 += "\n";
            }
            c14 += (ws.Cells[24, 1].Text + " |  " + ws.Cells[24, 2].Text);
            richTextBox12.Text = c14;
            for (int i = 12; i <= 13; i++)
            {
                b14 += (ws.Cells[i, 2].Text + " | " + ws.Cells[i, 3].Text + " | " + ws.Cells[i, 4].Text + " | " + ws.Cells[i, 5].Text);
                b14 += "\n";

            }
            richTextBox13.Text = b14;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            readExcel14();
        }
        private void richTextBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            readExcel13();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            readExcel3();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            readExcel6();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            readExcel7();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            readExcel8();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            readExcel9();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            readExcel10();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            readExcel11();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            readExcel4();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            readExcel5();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            readExcel12();
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void дегазация_Click(object sender, EventArgs e)
        {

        }

        private void химреагенты_Click(object sender, EventArgs e)
        {

        }
    }
}