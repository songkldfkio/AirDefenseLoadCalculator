using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

namespace MID多窗体
{
    /// <summary>
    /// form1类
    /// </summary>
    public partial class Form1 : Form
    {

        double[] p1 = new double[100];//数组p1，最多100组
        double[] p2 = new double[100];//数组p2，最多100组
        double[] p1c = new double[100];//乘以需要系数
        double[] p2c = new double[100];//乘以需要系数
        double[] q1c = new double[100];//
        double[] q2c = new double[100];//
        double[] k = new double[100];//数组k，最多100组
        double[] cos = new double[100];//数组cos，最多100组
        double[] S1 = new double[100];//
        double[] S2 = new double[100];//
        double p1zc;
        double p2zc;
        double q1zc;
        double q2zc;
        double kp;
        double kq;
        double kz;
        double cosz;
        double Sc;
        double m = 0;
        double n = 0;

        public Form1()//初始化form1
        {
            InitializeComponent();
            Size = new Size(432, 526);
        }


        /// <summary>
        /// 点击按钮增加一行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {



            int x = dataGridView1.Rows.Count;//获得行数x
            if (x < 100)
            {

                dataGridView1.Rows.Add();
                dataGridView2.Rows.Add();
                dataGridView2.Rows.Add();

                x = dataGridView1.Rows.Count;//获得行数x
                dataGridView1.Rows[x - 1].Cells[0].Value = "防护单元" + x;
                dataGridView1.Rows[x - 1].Cells[1].Value = "7.5";
                dataGridView1.Rows[x - 1].Cells[2].Value = "10.5";
                dataGridView1.Rows[x - 1].Cells[3].Value = "0.9";
                dataGridView1.Rows[x - 1].Cells[4].Value = "0.8";
            }
            else
            {
                MessageBox.Show("无法增加，仅支持100组数据！");
            }
            //强制生产数据，button3先点击一下
            button3.PerformClick();






        }


        /// <summary>
        /// 点击按钮删除一行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {




            int x = dataGridView1.Rows.Count;//获得行数x
            int y = dataGridView2.Rows.Count;//获得行数y
            if (x >= 1)
            {
                //int i = dataGridView1.CurrentRow.Index;//移除当前所在行
                //DataGridViewRow row1 = dataGridView1.Rows[i];
                //DataGridViewRow row2 = dataGridView2.Rows[2 * i];
                //DataGridViewRow row3 = dataGridView2.Rows[2 * i + 1];
                //dataGridView2.Rows.Remove(row1);
                //dataGridView2.Rows.Remove(row2);
                //dataGridView2.Rows.Remove(row3);


                dataGridView1.Rows.Remove(dataGridView1.Rows[x - 1]);
                dataGridView2.Rows.Remove(dataGridView2.Rows[y - 1]);
                dataGridView2.Rows.Remove(dataGridView2.Rows[y - 2]);

            }
            //强制生产数据，button3先点击一下
            button3.PerformClick();





        }


        /// <summary>
        /// 数据全部重置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {



            Array.Clear(p1, 0, p1.Length);//数重新置零，防止求和错误
            Array.Clear(p2, 0, p2.Length);//数重新置零，防止求和错误
            Array.Clear(p1c, 0, p1c.Length);//数重新置零，防止求和错误
            Array.Clear(p2c, 0, p2c.Length);//数重新置零，防止求和错误
            Array.Clear(q1c, 0, q1c.Length);//数重新置零，防止求和错误
            //Array.Clear(k, 0, k.Length);//数重新置零，防止求和错误
            //Array.Clear(cos, 0, cos.Length);//数重新置零，防止求和错误
            Array.Clear(S1, 0, S1.Length);//数重新置零，防止求和错误
            Array.Clear(S2, 0, S2.Length);//数重新置零，防止求和错误

            int x = dataGridView1.Rows.Count;//获得行数x
            for (int i = 1; i <= x; i++)
            {
                dataGridView1.Rows[i - 1].Cells[0].Value = "防护单元" + i;
                dataGridView1.Rows[i - 1].Cells[1].Value = "7.5";
                dataGridView1.Rows[i - 1].Cells[2].Value = "10.5";
                dataGridView1.Rows[i - 1].Cells[3].Value = "0.9";
                dataGridView1.Rows[i - 1].Cells[4].Value = "0.8";
            }

            ankp.Value = 0.9M;
            ankq.Value = 0.9M;

            //强制生产数据，button3先点击一下
            button3.PerformClick();


        }


        /// <summary>
        /// 不放大窗口计算,最核心代码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {


            int x = dataGridView1.Rows.Count;//获得行数x
            int y = dataGridView2.Rows.Count;//获得行数y
            int z = dataGridView3.Rows.Count;//获得行数z
            Array.Clear(p1, 0, p1.Length);//数重新置零，防止求和错误
            Array.Clear(p2, 0, p2.Length);//数重新置零，防止求和错误
            Array.Clear(p1c, 0, p1c.Length);//数重新置零，防止求和错误
            Array.Clear(p2c, 0, p2c.Length);//数重新置零，防止求和错误
            Array.Clear(q1c, 0, q1c.Length);//数重新置零，防止求和错误
            Array.Clear(q2c, 0, q2c.Length);//数重新置零，防止求和错误
            Array.Clear(k, 0, k.Length);//数重新置零，防止求和错误
            Array.Clear(cos, 0, cos.Length);//数重新置零，防止求和错误
            Array.Clear(S1, 0, S1.Length);//数重新置零，防止求和错误
            Array.Clear(S2, 0, S2.Length);//数重新置零，防止求和错误




            for (int i = 1; i <= x; i++)//表格和数据赋值给数组，转成double，可能出错，用输入的不是数字
            {
                try//转换出错时提示用户，输入数字
                {
                    p1[i - 1] = Convert.ToDouble(dataGridView1.Rows[i - 1].Cells[1].Value);
                    p2[i - 1] = Convert.ToDouble(dataGridView1.Rows[i - 1].Cells[2].Value);
                    k[i - 1] = Convert.ToDouble(dataGridView1.Rows[i - 1].Cells[3].Value);
                    cos[i - 1] = Convert.ToDouble(dataGridView1.Rows[i - 1].Cells[4].Value);
                }
                catch
                {
                    MessageBox.Show("请检查一二级负荷、需要系数、功率因数，应输入纯数字类型！");
                }
                kp = Convert.ToDouble(ankp.Value);
                kq = Convert.ToDouble(ankq.Value);




                p1c[i - 1] = p1[i - 1] * k[i - 1];//乘以系数后的数组
                p2c[i - 1] = p2[i - 1] * k[i - 1];//乘以系数后的数组

                q1c[i - 1] = p1c[i - 1] * Math.Tan(Math.Acos(cos[i - 1]));
                q2c[i - 1] = p2c[i - 1] * Math.Tan(Math.Acos(cos[i - 1]));
                S1[i - 1] = p1c[i - 1] / cos[i - 1];
                S2[i - 1] = p2c[i - 1] / cos[i - 1];
                p1zc = p1c.Sum();
                p2zc = p2c.Sum();
                q1zc = q1c.Sum();
                q2zc = q2c.Sum();


            }

            //Sc = Math.Sqrt((p1zc + p2zc) * (p1zc + p2zc) * kp * kp + (q1zc + q2zc) * (q1zc + q2zc) * kq * kq);
            kz = (p1zc * kp + p2zc * kp) / (p1.Sum() + p2.Sum());
            cosz = (p1zc * kp + p2zc * kp) / Sc;








            if (2 * x > y)
            {
                for (int i = 1; i <= 2 * x - y; i++)//设置表2行数，是表一二倍
                {
                    dataGridView2.Rows.Add();

                }

            }


            //向表格2写数据
            int j = 0;
            for (int i = 0; i <= 2 * x - 1; i = i + 2)
            {
                j = j + 1;

                dataGridView2.Rows[i].Cells[0].Value = dataGridView1.Rows[j - 1].Cells[0].Value;

                dataGridView2.Rows[i].Cells[1].Value = "一级";
                dataGridView2.Rows[i + 1].Cells[1].Value = "二级";

                dataGridView2.Rows[i].Cells[2].Value = p1[j - 1].ToString("0.00");
                dataGridView2.Rows[i + 1].Cells[2].Value = p2[j - 1].ToString("0.00");

                dataGridView2.Rows[i].Cells[3].Value = p1c[j - 1].ToString("0.00");
                dataGridView2.Rows[i + 1].Cells[3].Value = p2c[j - 1].ToString("0.00");

                dataGridView2.Rows[i].Cells[4].Value = q1c[j - 1].ToString("0.00");
                dataGridView2.Rows[i + 1].Cells[4].Value = q2c[j - 1].ToString("0.00");

                dataGridView2.Rows[i].Cells[5].Value = S1[j - 1].ToString("0.00");
                dataGridView2.Rows[i + 1].Cells[5].Value = S2[j - 1].ToString("0.00");

                dataGridView2.Rows[i].Cells[6].Value = "k=" + k[j - 1];
                dataGridView2.Rows[i + 1].Cells[6].Value = "cos=" + cos[j - 1];

            }

            if (z < 6)//表格3增加6行
            {
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
                dataGridView3.Rows.Add();
            }



            m = 0;
            for (int i = 0; i < y; i++)
            {

                m = m + Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
            }

            n = 0;
            for (int i = 0; i < y; i++)
            {

                n = n + Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value);
            }

            Sc = Math.Sqrt(m * m * kp * kp + n * n * kq * kq);
            cosz = (m * kp) / Sc;

            dataGridView3.Rows[0].Cells[0].Value = "合计";
            dataGridView3.Rows[0].Cells[1].Value = "一级";
            dataGridView3.Rows[0].Cells[2].Value = p1.Sum().ToString("0.00");
            dataGridView3.Rows[0].Cells[3].Value = p1c.Sum().ToString("0.00");
            dataGridView3.Rows[0].Cells[4].Value = q1c.Sum().ToString("0.00");
            dataGridView3.Rows[0].Cells[5].Value = S1.Sum().ToString("0.00");

            dataGridView3.Rows[1].Cells[0].Value = "";
            dataGridView3.Rows[1].Cells[1].Value = "二级";
            dataGridView3.Rows[1].Cells[2].Value = p2.Sum().ToString("0.00");
            dataGridView3.Rows[1].Cells[3].Value = p2c.Sum().ToString("0.00");
            dataGridView3.Rows[1].Cells[4].Value = q2c.Sum().ToString("0.00");
            dataGridView3.Rows[1].Cells[5].Value = S2.Sum().ToString("0.00");

            dataGridView3.Rows[2].Cells[0].Value = "";
            dataGridView3.Rows[2].Cells[1].Value = "一二级";
            dataGridView3.Rows[2].Cells[2].Value = (p1.Sum() + p2.Sum()).ToString("0.00");
            dataGridView3.Rows[2].Cells[3].Value = (p1c.Sum() + p2c.Sum()).ToString("0.00");
            dataGridView3.Rows[2].Cells[4].Value = (q1c.Sum() + q2c.Sum()).ToString("0.00");
            dataGridView3.Rows[2].Cells[5].Value = (S1.Sum() + S2.Sum()).ToString("0.00");

            dataGridView3.Rows[3].Cells[0].Value = "";
            dataGridView3.Rows[3].Cells[1].Value = "K∑p= " + kp;
            dataGridView3.Rows[3].Cells[2].Value = "K∑q= " + kq;
            dataGridView3.Rows[3].Cells[3].Value = "";
            dataGridView3.Rows[3].Cells[4].Value = "";
            dataGridView3.Rows[3].Cells[5].Value = "";

            dataGridView3.Rows[4].Cells[0].Value = "总计";
            dataGridView3.Rows[4].Cells[1].Value = "Pn= " + (p1.Sum() + p2.Sum()).ToString("0.00" + " kW");
            dataGridView3.Rows[4].Cells[2].Value = "Pc= " + (m * kp).ToString("0.00")+" kW";
            dataGridView3.Rows[4].Cells[3].Value = "Sc= " + Sc.ToString("0.00") + " kVA";
            dataGridView3.Rows[4].Cells[4].Value = "";
            dataGridView3.Rows[4].Cells[5].Value = "";
            dataGridView3.Rows[4].Cells[6].Value = "";

            dataGridView3.Rows[5].Cells[0].Value = ""; 
            dataGridView3.Rows[5].Cells[1].Value = "K∑= " + kz.ToString("0.000");
            dataGridView3.Rows[5].Cells[2].Value = "cos∑= " + cosz.ToString("0.000");
            dataGridView3.Rows[5].Cells[3].Value = "Ic= " + (Sc / (1.732 * 0.38)).ToString("0.00") + " A";
            dataGridView3.Rows[5].Cells[4].Value = "";
            dataGridView3.Rows[5].Cells[5].Value = "";
            dataGridView3.Rows[5].Cells[6].Value = "";





        }


        /// <summary>
        /// 生成规范表格，窗口放大
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //强制生产数据，button3先点击一下
            button3.PerformClick();


            //将窗口放大
            Size = new Size(1138, 526);



        }


        /// <summary>
        /// 打开新程序
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();
            frm2.Show();
        }

        /// <summary>
        /// 重置单元序号
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            int x = dataGridView1.Rows.Count;//获得行数x
            for (int i = 0; i < x; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = "防护单元" + (i + 1);
            }

            //强制生产数据，button3先点击一下
            button3.PerformClick();


        }
        /// <summary>
        /// 导出为excel，并隐藏原界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {

            HeBing();



            //开始打印dataGridView4至excel

            if (dataGridView2.Rows.Count == 0)
            {

                MessageBox.Show("当前无数据可导出!");
            }
            else
            {
                //DataGridViewToCsv(dataGridView4);//导csv格式

                try//保存文件名称已被打开时，提示失败
                {
                    //DataToExcel(dataGridView4);//导excel,标准格式
                    EEfangfa();//完美方法导出

                }
                catch 
                {
                    MessageBox.Show("此文件已被打开，无法保存！");

                }


                Size = new Size(726, 724);//隐藏原界面

                panel1.Visible = false;//隐藏原界面
            }



        }



        /// <summary>
        /// 导出旧版excel的方法
        /// </summary>
        /// <param name="m_DataView"></param>

        public void DataToExcel(DataGridView m_DataView)

        {

            SaveFileDialog kk = new SaveFileDialog();

            kk.Title = "保存EXECL文件";

            kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*";

            kk.FilterIndex = 1;

            if (kk.ShowDialog() == DialogResult.OK)

            {

                string FileName = kk.FileName /*+ ".xls"*/;

                if (File.Exists(FileName))

                    File.Delete(FileName);

                FileStream objFileStream;

                StreamWriter objStreamWriter;

                string strLine = "";

                objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);

                objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);

                for (int i = 0; i < m_DataView.Columns.Count; i++)

                {

                    if (m_DataView.Columns[i].Visible == true)

                    {

                        strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9);

                    }

                }

                objStreamWriter.WriteLine(strLine);

                strLine = "";

                for (int i = 0; i < m_DataView.Rows.Count; i++)

                {

                    if (m_DataView.Columns[0].Visible == true)

                    {

                        if (m_DataView.Rows[i].Cells[0].Value == null)

                            strLine = strLine + " " + Convert.ToChar(9);

                        else

                            strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9);

                    }

                    for (int j = 1; j < m_DataView.Columns.Count; j++)

                    {

                        if (m_DataView.Columns[j].Visible == true)

                        {

                            if (m_DataView.Rows[i].Cells[j].Value == null)

                                strLine = strLine + " " + Convert.ToChar(9);

                            else

                            {

                                string rowstr = "";

                                rowstr = m_DataView.Rows[i].Cells[j].Value.ToString();

                                if (rowstr.IndexOf("\r\n") > 0)

                                    rowstr = rowstr.Replace("\r\n", " ");

                                if (rowstr.IndexOf("\t") > 0)

                                    rowstr = rowstr.Replace("\t", " ");

                                strLine = strLine + rowstr + Convert.ToChar(9);

                            }

                        }

                    }

                    objStreamWriter.WriteLine(strLine);

                    strLine = "";

                }

                objStreamWriter.Close();

                objFileStream.Close();

                //MessageBox.Show(this, "保存EXCEL成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }





        //使用类
        public class ExportToExcel
        {
            public Excel.Application m_xlApp = null;
            public void OutputAsExcelFile(DataGridView dataGridView_result)
            {
                if (dataGridView_result.Rows.Count <= 0)
                {
                    MessageBox.Show("无数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                }
                string filePath = "";
                SaveFileDialog s = new SaveFileDialog
                {
                    Title = "保存Excel文件",
                    Filter = "Excel文件(*.xlsx)|*.xlsx",
                    FilterIndex = 1
                };
                if (s.ShowDialog() == DialogResult.OK)
                    filePath = s.FileName;
                else
                    return;

                //第一步：将dataGridView转化为dataTable,这样能够过滤掉dataGridView中的隐藏列

                DataTable tmpDataTable = new DataTable("tmpDataTable");
                DataTable modelTable = new DataTable("ModelTable");
                for (int column = 0; column < dataGridView_result.Columns.Count; column++)
                {
                    if (dataGridView_result.Columns[column].Visible == true)
                    {
                        DataColumn tempColumn = new DataColumn(dataGridView_result.Columns[column].HeaderText, typeof(string));
                        tmpDataTable.Columns.Add(tempColumn);
                        DataColumn modelColumn = new DataColumn(dataGridView_result.Columns[column].Name, typeof(string));
                        modelTable.Columns.Add(modelColumn);
                    }
                }
                for (int row = 0; row < dataGridView_result.Rows.Count; row++)
                {
                    if (dataGridView_result.Rows[row].Visible == false)
                        continue;
                    DataRow tempRow = tmpDataTable.NewRow();
                    for (int i = 0; i < tmpDataTable.Columns.Count; i++)
                        tempRow[i] = dataGridView_result.Rows[row].Cells[modelTable.Columns[i].ColumnName].Value;
                    tmpDataTable.Rows.Add(tempRow);
                }
                if (tmpDataTable == null)
                {
                    return;
                }

                //第二步：导出dataTable到Excel
                long rowNum = tmpDataTable.Rows.Count;//行数
                int columnNum = tmpDataTable.Columns.Count;//列数
                Excel.Application m_xlApp = new Excel.Application();
                m_xlApp.DisplayAlerts = false;//不显示更改提示
                m_xlApp.Visible = false;

                Excel.Workbooks workbooks = m_xlApp.Workbooks;
                Excel.Workbook workbook = workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];//取得sheet1

                try
                {
                    string[,] datas = new string[rowNum + 1, columnNum];
                    for (int i = 0; i < columnNum; i++) //写入字段
                        datas[0, i] = tmpDataTable.Columns[i].Caption;
                    Excel.Range range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
                    range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Bold = true;
                    range.Font.Size = 10;

                    int r = 0;
                    for (r = 0; r < rowNum; r++)
                    {
                        for (int i = 0; i < columnNum; i++)
                        {
                            object obj = tmpDataTable.Rows[r][tmpDataTable.Columns[i].ToString()];
                            datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引號是为了防止自己主动转化格式
                        }
                        System.Windows.Forms.Application.DoEvents();       //加入进度条

                    }

                    Excel.Range fchR = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
                    fchR.Value2 = datas;

                    worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。


                    m_xlApp.Visible = false;

                    range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];

                    range.Interior.ColorIndex = 15;//15代表灰色
                    range.Font.Size = 9;
                    range.RowHeight = 14.25;
                    range.Borders.LineStyle = 1;
                    range.HorizontalAlignment = 1;
                    workbook.Saved = true;
                    workbook.SaveCopyAs(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出异常：" + ex.Message, "导出异常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally
                {
                    EndReport();
                }

                m_xlApp.Workbooks.Close();
                m_xlApp.Workbooks.Application.Quit();
                m_xlApp.Application.Quit();
                m_xlApp.Quit();
                MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            private void EndReport()
            {
                try
                {

                }
                catch { }
                finally
                {
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Workbooks);    //如果电脑没有装正版的Execl会提示异常，把.xls格式改成wps的.xlsx格式导出成功
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Application);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp);
                        m_xlApp = null;
                    }
                    catch { }
                    try
                    {
                        //清理垃圾进程
                        this.KillProcessThread();
                    }
                    catch { }
                    GC.Collect();
                }
            }

            private void KillProcessThread()
            {
                ArrayList myProcess = new ArrayList();
                for (int i = 0; i < myProcess.Count; i++)
                {
                    try
                    {
                        System.Diagnostics.Process.GetProcessById(int.Parse((string)myProcess[i])).Kill();
                    }
                    catch { }
                }
            }
        }
        //

        public void EEfangfa()
        {
            ExportToExcel excel = new ExportToExcel();
            excel.OutputAsExcelFile(dataGridView4);
        }
        //









        /// <summary>
        /// 定义导出CSV格式的方法
        /// </summary>
        /// <param name="dgv"></param>
        public void DataGridViewToCSV(DataGridView dgv)
        {
            //程序实例化SaveFileDialog控件，并对该控件相关参数进行设置
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Execl files (*.csv)|*.csv";
            dlg.FilterIndex = 0;
            dlg.RestoreDirectory = true;
            dlg.CreatePrompt = true;
            dlg.Title = "保存为csv文件";
            //以上过程也可以通过添加控件，再设置控件属性完成，此处用程序编写出来了，在移植时就可摆脱控件的限制
            try//保存文件名称已被打开时，提示失败
            {
                if (dlg.ShowDialog() == DialogResult.OK)//打开SaveFileDialog控件，判断返回值结果
                {
                    Stream myStream;//流变量
                    myStream = dlg.OpenFile();//返回SaveFileDialog控件打开的文件，并将所选择的文件转化成流
                    StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));//将选择的文件流生成写入流
                    string columnTitle = "";
                    try
                    {

                        //写入列标题    
                        for (int i = 0; i < dgv.ColumnCount; i++)
                        {
                            if (i > 0)
                            {
                                columnTitle += ",";
                            }
                            columnTitle += dgv.Columns[i].HeaderText;//符号 ， 的添加，在保存为Excel时就以 ， 分成不同的列了
                        }

                        sw.WriteLine(columnTitle);//将内容写入文件流中

                        //写入列内容    
                        for (int j = 0; j < dgv.Rows.Count; j++)
                        {
                            string columnValue = "";
                            for (int k = 0; k < dgv.Columns.Count; k++)
                            {
                                if (k > 0)
                                {
                                    columnValue += ",";
                                }
                                if (dgv.Rows[j].Cells[k].Value == null)
                                    columnValue += "";
                                else if (dgv.Rows[j].Cells[k].Value.ToString().Contains(","))
                                {
                                    columnValue += "\"" + dgv.Rows[j].Cells[k].Value.ToString().Trim() + "\"";//将单元格中的，号转义成文本
                                }
                                else
                                {
                                    columnValue += dgv.Rows[j].Cells[k].Value.ToString().Trim() + "\t";//\t 横向跳格
                                }
                            }//获得写入到列中的值
                            sw.WriteLine(columnValue);//将内容写入文件流中
                        }
                        sw.Close();//关闭写入流
                        myStream.Close();//关闭流变量
                                         //MessageBox.Show("导出成功，可用天正电气“读入excel”导入CAD，多余问号改黑体可消除");
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("导出表格失败！");
                    }
                    finally
                    {
                        sw.Close();
                        myStream.Close();
                    }
                }
                //else
                //{
                //    MessageBox.Show("取消导出表格操作!");
                //}


            }
            catch (Exception e)
            {
                MessageBox.Show("文件已打开，无法保存！");

            }



        }

        /// <summary>
        /// dataGridView1值改变时点击button4计算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //强制button3点击事件执行一次，button3不缩放窗口
            button3.PerformClick();

        }

        /// <summary>
        /// kp改变时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ankp_ValueChanged(object sender, EventArgs e)
        {
            //强制button3点击事件执行一次，button3不缩放窗口
            button3.PerformClick();

        }

        /// <summary>
        /// kq改变时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ankq_ValueChanged(object sender, EventArgs e)
        {
            //强制button3点击事件执行一次，button3不缩放窗口
            button3.PerformClick();

        }

        /// <summary>
        /// 鼠标进入button2区域，红色警告
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_MouseEnter(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(255, 255, 0, 0);//设置背景色为红色，不透明度为100%。
        }
        /// <summary>
        /// 鼠标离开button2区域，恢复颜色
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_MouseLeave(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(0, 0, 0, 0);//设置背景色为白色，不透明度为100%。
        }

        /// <summary>
        /// 展开，收缩窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text == ">>")
            {
                Size = new Size(1138, 526);
                button9.Text = "<<";
            }
            else if (button9.Text == "<<")
            {
                Size = new Size(432, 526);
                button9.Text = ">>";
            }
            HeBing();

        }

        /// <summary>
        /// 返回计算界面，显示原界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button10_Click(object sender, EventArgs e)
        {

            Size = new Size(432, 526);
            panel1.Visible = true;

        }

        /// <summary>
        /// 将表格合并的方法
        /// </summary>
        public void HeBing()
        {
            //强制button3点击事件执行一次
            button3.PerformClick();

            //让dataGridView4行是dataGridView2和3之和
            int y = dataGridView2.Rows.Count;//获得行数
            int w = dataGridView4.Rows.Count;//获得行数
            if (y == 0 && w == 0)
            {
                dataGridView4.Rows.Add();
                dataGridView4.Rows.Add();
                dataGridView4.Rows.Add();
                dataGridView4.Rows.Add();
                dataGridView4.Rows.Add();
                dataGridView4.Rows.Add();
            }
            else if (y > w - 6)
            {
                for (int i = 0; i < y - w + 6; i++)
                {

                    dataGridView4.Rows.Add();

                }
            }
            else if (y < w - 6)
            {
                for (int i = 0; i < w - y - 6; i++)
                {

                    dataGridView4.Rows.Remove(dataGridView4.Rows[0]);

                }
            }
            else
            {

            }

            //向dataGridView4写数据
            y = dataGridView2.Rows.Count;//获得行数
            int z = dataGridView3.Rows.Count;//获得行数
            w = dataGridView4.Rows.Count;//获得行数
            for (int i = 0; i < y; i++)
            {
                dataGridView4.Rows[i].Cells[0].Value = dataGridView2.Rows[i].Cells[0].Value;
                dataGridView4.Rows[i].Cells[1].Value = dataGridView2.Rows[i].Cells[1].Value;
                dataGridView4.Rows[i].Cells[2].Value = dataGridView2.Rows[i].Cells[2].Value;
                dataGridView4.Rows[i].Cells[3].Value = dataGridView2.Rows[i].Cells[3].Value;
                dataGridView4.Rows[i].Cells[4].Value = dataGridView2.Rows[i].Cells[4].Value;
                dataGridView4.Rows[i].Cells[5].Value = dataGridView2.Rows[i].Cells[5].Value;
                dataGridView4.Rows[i].Cells[6].Value = dataGridView2.Rows[i].Cells[6].Value;

            }
            for (int i = 0; i < z; i++)
            {
                dataGridView4.Rows[y + i].Cells[0].Value = dataGridView3.Rows[i].Cells[0].Value;
                dataGridView4.Rows[y + i].Cells[1].Value = dataGridView3.Rows[i].Cells[1].Value;
                dataGridView4.Rows[y + i].Cells[2].Value = dataGridView3.Rows[i].Cells[2].Value;
                dataGridView4.Rows[y + i].Cells[3].Value = dataGridView3.Rows[i].Cells[3].Value;
                dataGridView4.Rows[y + i].Cells[4].Value = dataGridView3.Rows[i].Cells[4].Value;
                dataGridView4.Rows[y + i].Cells[5].Value = dataGridView3.Rows[i].Cells[5].Value;
                dataGridView4.Rows[y + i].Cells[6].Value = dataGridView3.Rows[i].Cells[6].Value;

            }
            //强制button3点击事件执行一次
            button3.PerformClick();
        }
        /// <summary>
        /// 复制最终的按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {

            //SendKeys.Send("^{A}");//模拟键盘ctrl+A
            //SendKeys.Send("^{C}");//模拟键盘ctrl+C
        }

        
    }
}
