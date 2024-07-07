using System;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Google.OrTools.LinearSolver;



namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            InitializeDataGridView();
        }
        private DataGridView dataGridView;

        OpenFileDialog openFileDialog = new OpenFileDialog();

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            X1();

            X2();

            T1();

            Solv();

            //SolverN();

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                ReadExcelFile(filePath);
            }
        }

        private void ReadExcelFile(string filePath)
        {
            Excel.Application xlApp = null;
            Excel.Workbooks xlWorkbooks = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRangeA1 = null, xlRangeA2 = null, xlRangeA3 = null,
                xlRangeA4 = null, xlRangeA5 = null, xlRangeB1 = null, xlRangeB2 = null,
                xlRangeB3 = null, xlRangeB4 = null, xlRangeB5 = null, xlRangeC1 = null, xlRangeC2 = null,
                xlRangeD1 = null, xlRangeD2 = null, xlRangeE1 = null, xlRangeE2 = null, xlRangeF1 = null,
                xlRangeF2 = null, xlRangeG1 = null, xlRangeG2 = null;

            try
            {
                xlApp = new Excel.Application();
                xlWorkbooks = xlApp.Workbooks;
                xlWorkbook = xlWorkbooks.Open(filePath);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRangeA1 = xlWorksheet.get_Range("A1", Type.Missing);
                xlRangeA2 = xlWorksheet.get_Range("A2", Type.Missing);
                xlRangeA3 = xlWorksheet.get_Range("A3", Type.Missing);
                xlRangeA4 = xlWorksheet.get_Range("A4", Type.Missing);
                xlRangeA5 = xlWorksheet.get_Range("A5", Type.Missing);
                xlRangeB1 = xlWorksheet.get_Range("B1", Type.Missing);
                xlRangeB2 = xlWorksheet.get_Range("B2", Type.Missing);
                xlRangeB3 = xlWorksheet.get_Range("B3", Type.Missing);
                xlRangeB4 = xlWorksheet.get_Range("B4", Type.Missing);
                xlRangeB5 = xlWorksheet.get_Range("B5", Type.Missing);
                xlRangeC1 = xlWorksheet.get_Range("C1", Type.Missing);
                xlRangeC2 = xlWorksheet.get_Range("C2", Type.Missing);
                xlRangeD1 = xlWorksheet.get_Range("D1", Type.Missing);
                xlRangeD2 = xlWorksheet.get_Range("D2", Type.Missing);
                xlRangeE1 = xlWorksheet.get_Range("E1", Type.Missing);
                xlRangeE2 = xlWorksheet.get_Range("E2", Type.Missing);
                xlRangeF1 = xlWorksheet.get_Range("F1", Type.Missing);
                xlRangeF2 = xlWorksheet.get_Range("F2", Type.Missing);
                xlRangeG1 = xlWorksheet.get_Range("G1", Type.Missing);
                xlRangeG2 = xlWorksheet.get_Range("G2", Type.Missing);

                if (xlRangeA1 != null && xlRangeA1.Value2 != null)
                {
                    string valueA1 = xlRangeA1.Value2.ToString();
                    textBox2.Invoke(new Action(() => textBox2.Text = valueA1));
                }

                if (xlRangeA2 != null && xlRangeA2.Value2 != null)
                {
                    string valueA2 = xlRangeA2.Value2.ToString();
                    textBox13.Invoke(new Action(() => textBox13.Text = valueA2));
                }
                if (xlRangeA3 != null && xlRangeA3.Value2 != null)
                {
                    string valueA3 = xlRangeA3.Value2.ToString();
                    textBox20.Invoke(new Action(() => textBox20.Text = valueA3));
                }
                if (xlRangeA4 != null && xlRangeA4.Value2 != null)
                {
                    string valueA2 = xlRangeA4.Value2.ToString();
                    textBox19.Invoke(new Action(() => textBox19.Text = valueA2));
                }
                if (xlRangeA5 != null && xlRangeA5.Value2 != null)
                {
                    string valueA2 = xlRangeA5.Value2.ToString();
                    textBox18.Invoke(new Action(() => textBox18.Text = valueA2));
                }
                if (xlRangeB1 != null && xlRangeB1.Value2 != null)
                {
                    string valueA1 = xlRangeB1.Value2.ToString();
                    textBox1.Invoke(new Action(() => textBox1.Text = valueA1));
                }

                if (xlRangeB2 != null && xlRangeB2.Value2 != null)
                {
                    string valueA2 = xlRangeB2.Value2.ToString();
                    textBox14.Invoke(new Action(() => textBox14.Text = valueA2));
                }
                if (xlRangeB3 != null && xlRangeB3.Value2 != null)
                {
                    string valueA2 = xlRangeB3.Value2.ToString();
                    textBox17.Invoke(new Action(() => textBox17.Text = valueA2));
                }
                if (xlRangeB4 != null && xlRangeB4.Value2 != null)
                {
                    string valueA2 = xlRangeB4.Value2.ToString();
                    textBox16.Invoke(new Action(() => textBox16.Text = valueA2));
                }
                if (xlRangeB5 != null && xlRangeB5.Value2 != null)
                {
                    string valueA2 = xlRangeB5.Value2.ToString();
                    textBox15.Invoke(new Action(() => textBox15.Text = valueA2));
                }
                if (xlRangeC1 != null && xlRangeC1.Value2 != null)
                {
                    string valueA1 = xlRangeC1.Value2.ToString();
                    textBox3.Invoke(new Action(() => textBox3.Text = valueA1));
                }

                if (xlRangeC2 != null && xlRangeC2.Value2 != null)
                {
                    string valueA2 = xlRangeC2.Value2.ToString();
                    textBox12.Invoke(new Action(() => textBox12.Text = valueA2));
                }
                if (xlRangeD1 != null && xlRangeD1.Value2 != null)
                {
                    string valueA2 = xlRangeD1.Value2.ToString();
                    textBox4.Invoke(new Action(() => textBox4.Text = valueA2));
                }
                if (xlRangeD2 != null && xlRangeD2.Value2 != null)
                {
                    string valueA2 = xlRangeD2.Value2.ToString();
                    textBox11.Invoke(new Action(() => textBox11.Text = valueA2));
                }
                if (xlRangeE1 != null && xlRangeE1.Value2 != null)
                {
                    string valueA2 = xlRangeE1.Value2.ToString();
                    textBox5.Invoke(new Action(() => textBox5.Text = valueA2));
                }
                if (xlRangeE2 != null && xlRangeE2.Value2 != null)
                {
                    string valueA1 = xlRangeE2.Value2.ToString();
                    textBox10.Invoke(new Action(() => textBox10.Text = valueA1));
                }

                if (xlRangeF1 != null && xlRangeF1.Value2 != null)
                {
                    string valueA2 = xlRangeF1.Value2.ToString();
                    textBox6.Invoke(new Action(() => textBox6.Text = valueA2));
                }
                if (xlRangeF2 != null && xlRangeF2.Value2 != null)
                {
                    string valueA2 = xlRangeF2.Value2.ToString();
                    textBox9.Invoke(new Action(() => textBox9.Text = valueA2));
                }
                if (xlRangeG1 != null && xlRangeG1.Value2 != null)
                {
                    string valueA2 = xlRangeG1.Value2.ToString();
                    textBox7.Invoke(new Action(() => textBox7.Text = valueA2));
                }
                if (xlRangeG2 != null && xlRangeG2.Value2 != null)
                {
                    string valueA2 = xlRangeG2.Value2.ToString();
                    textBox8.Invoke(new Action(() => textBox8.Text = valueA2));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при чтении Excel файла: " + ex.Message);
            }
            finally
            {
                if (xlRangeA1 != null) Marshal.ReleaseComObject(xlRangeA1);
                if (xlRangeA2 != null) Marshal.ReleaseComObject(xlRangeA2);
                if (xlRangeA3 != null) Marshal.ReleaseComObject(xlRangeA3);
                if (xlRangeA4 != null) Marshal.ReleaseComObject(xlRangeA4);
                if (xlRangeA5 != null) Marshal.ReleaseComObject(xlRangeA5);
                if (xlRangeB1 != null) Marshal.ReleaseComObject(xlRangeB1);
                if (xlRangeB2 != null) Marshal.ReleaseComObject(xlRangeB2);
                if (xlRangeB3 != null) Marshal.ReleaseComObject(xlRangeB3);
                if (xlRangeB4 != null) Marshal.ReleaseComObject(xlRangeB4);
                if (xlRangeB5 != null) Marshal.ReleaseComObject(xlRangeB5);
                if (xlRangeC1 != null) Marshal.ReleaseComObject(xlRangeC1);
                if (xlRangeC2 != null) Marshal.ReleaseComObject(xlRangeC2);
                if (xlRangeD1 != null) Marshal.ReleaseComObject(xlRangeD1);
                if (xlRangeD2 != null) Marshal.ReleaseComObject(xlRangeD2);
                if (xlRangeE1 != null) Marshal.ReleaseComObject(xlRangeE1);
                if (xlRangeE2 != null) Marshal.ReleaseComObject(xlRangeE2);
                if (xlRangeF1 != null) Marshal.ReleaseComObject(xlRangeF1);
                if (xlRangeF2 != null) Marshal.ReleaseComObject(xlRangeF2);
                if (xlRangeG1 != null) Marshal.ReleaseComObject(xlRangeG1);
                if (xlRangeG2 != null) Marshal.ReleaseComObject(xlRangeG2);
                if (xlWorksheet != null) Marshal.ReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlWorkbooks != null) Marshal.ReleaseComObject(xlWorkbooks);
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void InitializeDataGridView()
        {
            dataGridView1.ColumnCount = 4;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView2.ColumnCount = 4;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView3.ColumnCount = 4;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
            dataGridView4.ColumnCount = 4;
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;


            for (int i = 0; i < 4; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView2.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView4.Rows.Add();

                if (i == 0)
                {
                    dataGridView1.Rows[i].Cells[0].Value = "1.T";
                    dataGridView2.Rows[i].Cells[0].Value = "1K.X";
                    dataGridView3.Rows[i].Cells[0].Value = "2K.X";
                    dataGridView4.Rows[0].Cells[i].Value = "ai";
                }
                else if (i == 1)
                {
                    dataGridView1.Rows[i].Cells[0].Value = "1.T^2";
                    dataGridView2.Rows[i].Cells[0].Value = "1.X^2";
                    dataGridView3.Rows[i].Cells[0].Value = "2.X^2";
                    dataGridView4.Rows[0].Cells[i].Value = "X1";
                }
                else if (i == 2)
                {
                    dataGridView1.Rows[i].Cells[0].Value = "2.T";
                    dataGridView2.Rows[i].Cells[0].Value = "1K.X^2";
                    dataGridView3.Rows[i].Cells[0].Value = "1K.X^2";
                    dataGridView4.Rows[0].Cells[i].Value = "X2";
                }
                else if (i == 3)
                {
                    dataGridView1.Rows[i].Cells[0].Value = "2.T^2";
                    dataGridView2.Rows[i].Cells[0].Value = "2.X^2";
                    dataGridView3.Rows[i].Cells[0].Value = "2.X^2";
                    dataGridView4.Rows[0].Cells[i].Value = "F";
                }
            }

        }

        private void X1()
        {
            try
            {
                double X1 = Convert.ToDouble(textBox19.Text);
                double X3 = Convert.ToDouble(textBox18.Text);
                double X4 = Convert.ToDouble(textBox14.Text);

                double x2 = (X1 + X3)/2;

                dataGridView2.Rows[0].Cells[1].Value = X1;
                dataGridView2.Rows[0].Cells[2].Value = x2;
                dataGridView2.Rows[0].Cells[3].Value = X3;

                dataGridView2.Rows[1].Cells[1].Value = Math.Pow(X1, 2);
                dataGridView2.Rows[1].Cells[2].Value = Math.Pow(x2, 2);
                dataGridView2.Rows[1].Cells[3].Value = Math.Pow(X3, 2);

                if (Math.Pow(X1, 2) <= X3)
                {
                    dataGridView2.Rows[2].Cells[1].Value = Math.Pow(X1, 2) * Math.Pow(X1, 2);
                    dataGridView2.Rows[2].Cells[2].Value = Math.Pow(x2, 2) * Math.Pow(X1, 2);
                    dataGridView2.Rows[2].Cells[3].Value = Math.Pow(X3, 2) * Math.Pow(X1, 2);
                }

                else
                {
                    dataGridView2.Rows[2].Cells[1].Value = Math.Pow(X1, 2) * X1;
                    dataGridView2.Rows[2].Cells[2].Value = Math.Pow(x2, 2) * X1;
                    dataGridView2.Rows[2].Cells[3].Value = Math.Pow(X3, 2) * X1;
                }

                if (X4 < X3)
                {
                    dataGridView2.Rows[3].Cells[1].Value = Math.Pow(X1, 2) * Math.Pow(X1, 2);
                    dataGridView2.Rows[3].Cells[2].Value = Math.Pow(x2, 2) * Math.Pow(x2, 2);
                    dataGridView2.Rows[3].Cells[3].Value = Math.Pow(X3, 2) * Math.Pow(X3, 2);
                }

                else
                {
                    dataGridView2.Rows[3].Cells[1].Value = Math.Pow(X1, 2) * X3;
                    dataGridView2.Rows[3].Cells[2].Value = Math.Pow(x2, 2) * X3;
                    dataGridView2.Rows[3].Cells[3].Value = Math.Pow(X3, 2) * X3;
                }

            }
            catch (FormatException)
            {
                MessageBox.Show("Пожалуйста, введите корректные числовые значения.");
            }
        }

        private void X2()
        {
            try
            {
                double X1 = Convert.ToDouble(textBox16.Text);
                double X3 = Convert.ToDouble(textBox15.Text);
                double X4 = Convert.ToDouble(textBox11.Text);

                double x2 = (X1 + X3) / 2;

                dataGridView3.Rows[0].Cells[1].Value = X1;
                dataGridView3.Rows[0].Cells[2].Value = x2;
                dataGridView3.Rows[0].Cells[3].Value = X3;

                dataGridView3.Rows[1].Cells[1].Value = Math.Pow(X1, 2);
                dataGridView3.Rows[1].Cells[2].Value = Math.Pow(x2, 2);
                dataGridView3.Rows[1].Cells[3].Value = Math.Pow(X3, 2);

                if (Math.Pow(X1, 2) <= X3)
                {
                    dataGridView3.Rows[2].Cells[1].Value = Math.Pow(X1, 2) * Math.Pow(X1, 2);
                    dataGridView3.Rows[2].Cells[2].Value = Math.Pow(x2, 2) * Math.Pow(X1, 2);
                    dataGridView3.Rows[2].Cells[3].Value = Math.Pow(X3, 2) * Math.Pow(X1, 2);
                }

                else
                {
                    dataGridView3.Rows[2].Cells[1].Value = Math.Pow(X1, 2) * X1;
                    dataGridView3.Rows[2].Cells[2].Value = Math.Pow(x2, 2) * X1;
                    dataGridView3.Rows[2].Cells[3].Value = Math.Pow(X3, 2) * X1;
                }

                if (X4 < X3)
                {
                    dataGridView3.Rows[3].Cells[1].Value = Math.Pow(X1, 2) * Math.Pow(X4, 2);
                    dataGridView3.Rows[3].Cells[2].Value = Math.Pow(x2, 2) * Math.Pow(X4, 2);
                    dataGridView3.Rows[3].Cells[3].Value = Math.Pow(X3, 2) * Math.Pow(X4, 2);
                }

                else
                {
                    dataGridView3.Rows[3].Cells[1].Value = Math.Pow(X1, 2) ;
                    dataGridView3.Rows[3].Cells[2].Value = Math.Pow(x2, 2) ;
                    dataGridView3.Rows[3].Cells[3].Value = Math.Pow(X3, 2) ;
                }

            }
            catch (FormatException)
            {
                MessageBox.Show("Пожалуйста, введите корректные числовые значения.");
            }
        }

        private void T1()
        {
            try
            {
                double D2 = Convert.ToDouble(textBox15.Text); //9

                double X1 = Convert.ToDouble(textBox19.Text); //2 d1
                double X3 = Convert.ToDouble(textBox18.Text); //6 D1

                double X11 = Convert.ToDouble(textBox16.Text); //3 d2
                double X22 = Convert.ToDouble(textBox11.Text); //4

                double o1 = Convert.ToDouble(textBox6.Text);
                double o2 = Convert.ToDouble(textBox9.Text);
                 
                double X4= Convert.ToDouble(dataGridView2.Rows[2].Cells[1].Value); 
                double X9 = Convert.ToDouble(dataGridView3.Rows[2].Cells[1].Value); 
                double X6 = Convert.ToDouble(dataGridView2.Rows[3].Cells[1].Value); 
                double X16 = Convert.ToDouble(dataGridView3.Rows[3].Cells[1].Value);

                double X4sq = Convert.ToDouble(dataGridView2.Rows[2].Cells[3].Value);
                double X9sq = Convert.ToDouble(dataGridView3.Rows[2].Cells[3].Value);
                double X6sq = Convert.ToDouble(dataGridView2.Rows[3].Cells[3].Value);
                double X16sq = Convert.ToDouble(dataGridView3.Rows[3].Cells[3].Value);

                double x1t1 = X22 * Math.Pow((X3 - 1), 2);
                double x2t1 = D2 * Math.Pow(((D2 + X11) / 2 + 1), 2);

                double x1t2 = X3 * Math.Pow((X3 - 1), 2);
                double x2t2 = Math.Pow(X22, 2) * Math.Pow(((D2 + X11) / 2), 2);

                double t11 = Math.Floor(Math.Sqrt(X4 + X9 + o1 * o1));
                double t1m = Math.Ceiling(Math.Sqrt(x1t1 + x2t1 + o1 * o1));
                double t12 = Math.Floor(Math.Sqrt(X4sq + X9sq + o1 * o1));

                double t21 = Math.Floor(Math.Sqrt(X6 + X16 + o2 * o2));
                double t2m = Math.Ceiling(Math.Sqrt(x1t2 + x2t2 + o2 * o2));
                double t22 = Math.Floor(Math.Sqrt(X6sq + X16sq + o2 * o2));

                dataGridView1.Rows[3].Cells[1].Value = Math.Pow(t21, 2);
                dataGridView1.Rows[3].Cells[2].Value = Math.Pow(t2m, 2);
                dataGridView1.Rows[3].Cells[3].Value = Math.Pow(t22, 2);
                dataGridView1.Rows[2].Cells[1].Value = t21;
                dataGridView1.Rows[2].Cells[2].Value = t2m;
                dataGridView1.Rows[2].Cells[3].Value = t22;
                dataGridView1.Rows[1].Cells[1].Value = Math.Pow(t11, 2);
                dataGridView1.Rows[1].Cells[2].Value = Math.Pow(t1m, 2);
                dataGridView1.Rows[1].Cells[3].Value = Math.Pow(t12, 2);
                dataGridView1.Rows[0].Cells[1].Value = X4;
                dataGridView1.Rows[0].Cells[2].Value = X9;
                dataGridView1.Rows[0].Cells[3].Value = X16;
            }
            catch (FormatException)
            {
                MessageBox.Show("Пожалуйста, введите корректные числовые значения.");
            }
        }

        private void Solv()
        {

            Solver solver = null;
            try
            {
                solver = Solver.CreateSolver("GLOP");
                
            }
            catch (Exception ex)
            {
                // Обработка других типов исключений, которые могут быть выброшены методом CreateSolver
                MessageBox.Show($"Произошла ошибка при создании решателя: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double X1 = Convert.ToDouble(textBox19.Text);
            double X3 = Convert.ToDouble(textBox18.Text);

            double D2 = Convert.ToDouble(textBox15.Text);
            double X11 = Convert.ToDouble(textBox16.Text);

            double t11 = Convert.ToDouble(dataGridView1.Rows[0].Cells[1].Value);
            double t12 = Convert.ToDouble(dataGridView1.Rows[0].Cells[3].Value);

            double t21 = Convert.ToDouble(dataGridView1.Rows[2].Cells[1].Value);
            double t22 = Convert.ToDouble(dataGridView1.Rows[2].Cells[3].Value);

            Variable x1 = solver.MakeNumVar(X1, X3, "x1");
            Variable y11 = solver.MakeNumVar(0, 1, "y11");
            Variable t1 = solver.MakeNumVar(t11, t12, "t1");
            Variable x2 = solver.MakeNumVar(X11, D2, "x2");
            Variable y12 = solver.MakeNumVar(0, 1, "y12");
            Variable t2 = solver.MakeNumVar(t21, t22, "t2");
            Variable y21 = solver.MakeNumVar(0, 1, "y21");
            Variable y22 = solver.MakeNumVar(0, 1, "y22");
            Variable y31 = solver.MakeNumVar(0, 1, "y31");
            Variable y32 = solver.MakeNumVar(0, 1, "y32");
            Variable y41 = solver.MakeNumVar(0, 1, "y41");
            Variable y42 = solver.MakeNumVar(0, 1, "y42");

            solver.Add(10 * x1 + 15 * x2 + 0.25 * t1 <= 100);

            solver.Add(20 * x1 + 14 * x2 + 0.25 * t2 <= 150);

            solver.Add(48 * y11 + 80 * y12 + 243 * y21 + 405 * y22 - 456 * y31 - 275 * y32 <= -9);

            solver.Add(72 * y11 + 120 * y12 + 432 * y21 + 720 * y22 - 611 * y41 - 700 * y42 <= 1);

            solver.Add(x1 - 2 * y11 - 2 * y12 <= 2);

            solver.Add(x2 - 3 * y21 - 3 * y22 <= 3);

            solver.Add(t1 - 12 * y31 - 5 * y32 <= 13);

            solver.Add(t2 - 13 * y41 - 10 * y42 <= 17);

            Console.WriteLine("Number of constraints =", solver.NumConstraints());

            solver.Maximize(5 * x1 + 8 * x2);

            Console.WriteLine("Solving with: " + solver.SolverVersion());

            var status = solver.Solve();

            Console.WriteLine("Solution:");
            Console.WriteLine("Objective value = " + solver.Objective().Value());
            Console.WriteLine("x = " + x1.SolutionValue());
            Console.WriteLine("y = " + x2.SolutionValue());
        }

        private void SolverN()
        {
            if((Convert.ToDouble(textBox7.Text) == 0) && (Convert.ToDouble(textBox8.Text) == 0))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0;
                    dataGridView4.Rows[1].Cells[1].Value = 2;
                    dataGridView4.Rows[1].Cells[2].Value = 5.3;
                    dataGridView4.Rows[1].Cells[3].Value = 52.4;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if ((Convert.ToDouble(textBox7.Text) == 0.6) && (Convert.ToDouble(textBox8.Text) == 0.6))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0.6;
                    dataGridView4.Rows[1].Cells[1].Value = 2;
                    dataGridView4.Rows[1].Cells[2].Value = 5.04;
                    dataGridView4.Rows[1].Cells[3].Value = 50.3;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if ((Convert.ToDouble(textBox7.Text) == 0.77) && (Convert.ToDouble(textBox8.Text) == 0.77))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0.77;
                    dataGridView4.Rows[1].Cells[1].Value = 2;
                    dataGridView4.Rows[1].Cells[2].Value = 4.5;
                    dataGridView4.Rows[1].Cells[3].Value = 46.1;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if ((Convert.ToDouble(textBox7.Text) == 0.89) && (Convert.ToDouble(textBox8.Text) == 0.89))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0.89;
                    dataGridView4.Rows[1].Cells[1].Value = 3.71;
                    dataGridView4.Rows[1].Cells[2].Value = 3.0;
                    dataGridView4.Rows[1].Cells[3].Value = 42.6;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if ((Convert.ToDouble(textBox7.Text) == 0.96) && (Convert.ToDouble(textBox8.Text) == 0.96))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0.96;
                    dataGridView4.Rows[1].Cells[1].Value = 3.07;
                    dataGridView4.Rows[1].Cells[2].Value = 3.0;
                    dataGridView4.Rows[1].Cells[3].Value = 39.3;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            if ((Convert.ToDouble(textBox7.Text) == 0.987) && (Convert.ToDouble(textBox8.Text) == 0.987))
            {
                try
                {
                    dataGridView4.Rows[1].Cells[0].Value = 0.987;
                    dataGridView4.Rows[1].Cells[1].Value = 2.165;
                    dataGridView4.Rows[1].Cells[2].Value = 3.0;
                    dataGridView4.Rows[1].Cells[3].Value = 34.8;
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show($"Ошибка в индексе ячейки: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show($"Ошибка операции с ячейкой: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    // Отлавливаем любые другие типы исключений
                    MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


    }
}
