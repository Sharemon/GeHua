using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.Office.Interop.Word;

namespace cov
{
    public partial class Form1 : Form
    {
        delegate void UpdateUICallback();
        UpdateUICallback updateUI;

        public Form1()
        {
            updateUI = new UpdateUICallback(UpdateState);
            InitializeComponent();
        }

        private const int LENGTH = 10000;   //excel数据长度

        private Bitmap bmp;
        private Graphics ghp;
        private Pen myPen;
        private double Max = 1;         //计算结果中的最大值的绝对值
        private int MaxNum = -1;        //计算结果中的最大值
        private int MaxMinus = 1;       //计算结果的正负号，1为正，-1为负
        private float scale = 1;        //曲线图横坐标单位

        //更新曲线图
        private void UpdateState()
        {
            //初始化
            ghp = Graphics.FromImage(bmp);
            ghp.Clear(pictureBox1.BackColor);
            myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            myPen.Color = Color.DarkGreen;

            //画主坐标轴
            ghp.DrawLine(myPen,0,(float)pictureBox1.Height/2,pictureBox1.Width,(float)pictureBox1.Height/2);
            myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            ghp.DrawLine(myPen, (float)pictureBox1.Width / 2, 0, (float)pictureBox1.Width / 2, pictureBox1.Height);
            myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            
            //确定横坐标单位scale
            if (!string.IsNullOrEmpty(label4.Text))
            {
                if (label4.Text[label4.Text.Length - 1] == 'M' || label4.Text[label4.Text.Length - 1] == 'm')
                {
                    scale = 1 * 1 / float.Parse(label4.Text.Substring(0, label4.Text.Length - 1));
                }
                if (label4.Text[label4.Text.Length - 1] == 'K' || label4.Text[label4.Text.Length - 1] == 'k')
                {
                    scale = 1000 * 1 / float.Parse(label4.Text.Substring(0, label4.Text.Length - 1));
                }
                if (label4.Text[label4.Text.Length - 1] == 'G' || label4.Text[label4.Text.Length - 1] == 'g')
                {
                    scale = 0.001F * 1 / float.Parse(label4.Text.Substring(0, label4.Text.Length - 1));
                }
                ghp.DrawString("单位：us", new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red), 1, 1);
            }

            //画曲线，纵坐标范围为±|Max|
            myPen.Color = Color.BlueViolet;
            for (int i = 0; i < result.Count - 1; i++)
            {
                ghp.DrawLine(myPen, (float)(i) / (float)(2 * (LENGTH) - 1) * (float)pictureBox1.Width, (float)(1.05 * Max - result[i]) / 2/(float)(1.05 * Max) * (float)pictureBox1.Height,
                    (float)(i + 1) / (float)(2 * (LENGTH) - 1) * (float)pictureBox1.Width, (float)(1.05 * Max - result[i + 1]) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height);
            }

            //标出最大值点
            if (MaxNum != -1)
            {
                //myPen = new Pen(Color.Red);
                ghp.FillEllipse(new SolidBrush(Color.Red), (float)(MaxNum) / (2 * (LENGTH) - 1) * pictureBox1.Width - 4, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height - 4, 8, 8);
                //ghp.DrawEllipse(myPen, (float)MaxNum / int.Parse(Len) * pictureBox1.Width, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height, 2, 2);
                ghp.DrawString("采样点序号差：" + ((MaxNum - (LENGTH-1))).ToString(), new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),
                    (float)MaxNum / (2 * (LENGTH) - 1) * pictureBox1.Width + 8, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height + MaxMinus * 8);
            }

            //画次坐标轴
            myPen.Color = Color.DarkGreen;
            for (int i = 0; i < 5; i++)
            {
                ghp.DrawLine(myPen, (float)pictureBox1.Width / 2 - (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 + 5,
                    (float)pictureBox1.Width / 2 - (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 - 5);
                ghp.DrawLine(myPen, (float)pictureBox1.Width / 2 + (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 + 5,
                    (float)pictureBox1.Width / 2 + (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 - 5);
                if (!string.IsNullOrEmpty(label4.Text))
                {
                    ghp.DrawString(
                        (0 - scale * i * (LENGTH) / 5) - (int)(0 - scale * i * (LENGTH) / 5) == 0 ?
                        (0 - scale * i * (LENGTH) / 5).ToString("0") : (0 - scale * i * (LENGTH) / 5).ToString("0.000"),
                        new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),
                        (float)pictureBox1.Width / 2 - (float)pictureBox1.Width * i / 10 - 20, (float)pictureBox1.Height / 2 + 8);
                    ghp.DrawString(
                        (scale * i * (LENGTH) / 5) - (int)(scale * i * (LENGTH) / 5) == 0 ?
                        (scale * i * (LENGTH) / 5).ToString("0") : (scale * i * (LENGTH) / 5).ToString("0.000"),
                        new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),
                       (float)pictureBox1.Width / 2 + (float)pictureBox1.Width * i / 10 - 20, (float)pictureBox1.Height / 2 + 8);
                }
            }
            pictureBox1.Image = bmp;
        }

        System.Threading.Thread DataCal;            //计算互相关的线程
        private DataTable dt = new DataTable();     //从第一个excel里面得到的数据
        private DataTable dt2 = new DataTable();    //从第二个excel里面得到的数据
        private List<double> data1 = new List<double>();
        private List<double> data2 = new List<double>();
        List<double> result = new List<double>();   //计算的结果
        private int Len = 0;                        //有效计算点数

        private void button1_Click(object sender, EventArgs e)
        {
            //如果计算中点击按钮，中断计算
            if (DataCal != null && DataCal.ThreadState == System.Threading.ThreadState.Running)
            {
                DataCal.Abort();
            }
            dt = GetDataFromExcelByConn(false);     //调用GetDataFromExcelByConn得到数据
            dt2 = GetDataFromExcelByConn(false);
            if (dt != null && dt2 != null)
            {
                //从文件名中得到采样频率
                for(int i = 0; i< label2.Text.Length; i++)
                {
                    if(label2.Text[i] == 'M' || label2.Text[i] == 'K' || label2.Text[i] == 'G' || label2.Text[i] == 'k' || label2.Text[i] == 'g' || label2.Text[i] == 'm' )
                    {
                        label4.Text = label2.Text.Substring(0,i+1);
                        break;
                    }
                }
                label8.Visible = false;             //label8为指针放在曲线图上显示的坐标的显示控件
                MouseDowned = false;
                ScaleFlag = false;
                FirstMove = true;
                Max = 1;                            
                MaxNum = -1;                        
                MaxMinus = 1;
                related = 0;
                this.RShow.Text = "";
                Len = int.Parse(textBox1.Text);
                //计算有效点数不能超过数据点数的一半
                if (Len > (dt.Rows.Count/10*10) / 2)
                {
                    Len = dt.Rows.Count / 10 * 10 / 2;
                    textBox1.Text = Len.ToString();
                    MessageBox.Show("有效计算点数超出范围","提示");
                }
                dt.Columns.Add();
                //将两个表的数据合并到一个表dt

                data1.Clear();
                data2.Clear();
                for (int i = 2; i < dt2.Rows.Count; i++)
                {
                    //dt.Rows[i][2] = dt2.Rows[i][1];
                    data1.Add(Convert.ToDouble(dt.Rows[i][1]));
                    data2.Add(Convert.ToDouble(dt2.Rows[i][1]));
                }
                result.Clear();
                CalComp = false;
                DataCal = new System.Threading.Thread(new System.Threading.ThreadStart(cal));
                DataCal.Start();            //开始计算
            }
        }

        private bool CalComp = false;       //是否计算完成
        private double related = 0;

        //计算子线程
        private void cal()
        {
            this.Invoke(new EventHandler(delegate
                {
                    label6.Text = "计算中....";
                }));
            double[] aver = new double[2];
            aver[0] = 0;
            aver[1] = 0;
            for (int j = 0; j < LENGTH; j++)
            {
                aver[0] = aver[0] + data1[j];
                aver[1] = aver[1] + data2[j];
            }
            aver[0] = aver[0] / LENGTH;
            aver[1] = aver[1] / LENGTH;
            //开始计算
            /*
            for (int j = 0; j < 2 * (LENGTH - Len) + 1; j++)
            {
                result.Add(0);
                if (j < LENGTH - Len + 1)
                {
                    for (int i = (Len) + 1; i > 1; i--)
                    {
                        result[j] += (data2[i] - aver[1]) * (data1[LENGTH + 1 - j - (Len + 1 - i)] - aver[0]);
                    }
                }
                else
                {
                    for (int i = 2; i < (Len) + 2; i++)
                    {
                        result[j] += (data1[i] - aver[0]) * (data2[i + (j - (LENGTH + 1 - Len)) + 1] - aver[1]);
                    }
                }
                //更新最大值
                if (Math.Abs(result[j]) >= Max)
                {
                    Max = Math.Abs(result[j]);
                    MaxNum = j;
                    MaxMinus = result[j] > 0 ? 1 : -1;
                }
                //更新曲线图
                if ((0 == j % 1000) || (j == 2 * (LENGTH - Len)))
                {
                    this.Invoke(updateUI);
                }
            }*/
            for (int i = 0; i < 2 * LENGTH - 1; i++)
            {
                result.Add(0);
                int k = i - (LENGTH - 1);
                for (int j = 0; j < LENGTH; j++)
                {
                    if (j - k >= 0 && j - k < LENGTH)
                        result[i] += (data1[j]-aver[0]) * (data2[j - k]-aver[1]);
                }
                //更新最大值
                if (Math.Abs(result[i]) >= Max)
                {
                    Max = Math.Abs(result[i]);
                    MaxNum = i;
                    MaxMinus = result[i] > 0 ? 1 : -1;
                }
                //更新曲线图
                if ((0 == i % 1000) || (i == 2 * (LENGTH)))
                {
                    this.Invoke(updateUI);
                }
            }
            this.Invoke(updateUI);
            double[] sum = new double[3];
            sum[0] = 0;
            sum[1] = 0;
            sum[2] = 0;
            if (MaxNum <= LENGTH - 1)
            {
                for (int i = 0; i <= MaxNum ; i++)
                {
                    sum[0] += (data1[i] - aver[0]) * (data2[i + LENGTH - 1 - MaxNum] - aver[1]);
                    sum[1] = sum[1] + Math.Pow((data2[i] - aver[1]), 2);
                    sum[2] = sum[2] + Math.Pow((data1[i + LENGTH - 1 - MaxNum] - aver[0]), 2);
                }
            }
            else
            {
                for (int i = 0; i < (2 * LENGTH - 1 -MaxNum); i++)
                {
                    sum[0] += (data2[i] - aver[1]) * (data1[i + (MaxNum - LENGTH) + 1] - aver[0]);
                    sum[1] = sum[1] + Math.Pow((data1[i] - aver[0]), 2);
                    sum[2] = sum[2] + Math.Pow((data2[i + (MaxNum - LENGTH) + 1] - aver[1]), 2);
                }
            }
            related = sum[0] / Math.Sqrt(sum[1]*sum[2]);
            CalComp = true;
            this.Invoke(new EventHandler(delegate
                {
                    label6.Text = "计算完成";
                    RShow.Text = "最高峰相关系数： " + (Math.Floor(related*10000)/10000).ToString("0.0000");
                }));


        }

        private string FileName = "";
        //将exceel中的数据读到DataTable中
        DataTable GetDataFromExcelByConn(bool hasTitle = false)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            if (!string.IsNullOrEmpty(FileName))
            {
                openFile.InitialDirectory = System.IO.Path.GetDirectoryName(FileName);
            }
            else
            {
                openFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            openFile.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
            openFile.Multiselect = false;
            if (openFile.ShowDialog() == DialogResult.Cancel) return null;
            var filePath = openFile.FileName;
            FileName = filePath;
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType)) return null;
            label2.Text = System.IO.Path.GetFileName(openFile.FileName);

            using (DataSet ds = new DataSet())
            {
                string strCon = string.Format("Provider=Microsoft.Jet.OLEDB.{0}.0;" +
                                "Extended Properties=\"Excel {1}.0;HDR={2};IMEX=1;\";" +
                                "data source={3};",
                                (fileType == ".xls" ? 4 : 12), (fileType == ".xls" ? 8 : 12), (hasTitle ? "Yes" : "NO"), filePath);
                using (OleDbConnection myConn = new OleDbConnection(strCon))
                {
                    myConn.Open();
                    string sheetname = "";
                    System.Data.DataTable dt = myConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    //if (dt.Rows.Count == 1)
                    {
                        sheetname = dt.Rows[0]["TABLE_NAME"].ToString();
                    }
                    //else
                    //{
                        ;//多个Sheet处理
                    //}
                    string strCom = " SELECT * FROM [" + sheetname + "]";
                    OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                    myCommand.Fill(ds);
                    myCommand.Dispose();
                }
                if (ds == null || ds.Tables.Count <= 0) return null;
                return ds.Tables[0];
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label8.Parent = pictureBox1;
            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            myPen = new Pen(Color.BlueViolet);
            this.Invoke(updateUI);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(DataCal != null)
            {
                DataCal.Abort();
            }
        }

        //在pictureBox1控件中，鼠标按下拖动一段距离松开，可实现放大功能
        private bool MouseDowned = false;       //是否按下了鼠标左键，若按下才执行pictureBox1_MouseMove中的程序
        private float InitX = 0;                //记录下按下鼠标时的X坐标
        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (CalComp)
            {
                if (!ScaleFlag)
                {
                    MouseDowned = true;
                    float x = e.X;
                    InitX = x;
                    ghp = Graphics.FromImage(bmp);
                    myPen.Color = Color.Black;
                    myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                    ghp.DrawLine(myPen, x, 0, x, (float)pictureBox1.Height);
                    ghp.DrawString(Math.Round(((x - (float)pictureBox1.Width /2) / (float)pictureBox1.Width * 2 * (LENGTH))).ToString(),
                        new System.Drawing.Font("宋体", 8, FontStyle.Bold), new SolidBrush(Color.Black), x > pictureBox1.Width / 2 ? x - 16 : x + 10, e.Y);
                    pictureBox1.Image = bmp;
                }
                else
                {
                    ScaleFlag = false;
                    this.Invoke(updateUI);
                }
            }
            else
            {
                MessageBox.Show("请等待计算完成","提示");
            }
            
        }

        private bool FirstMove = true;          //是否为第一次移动，第一次移动时记下原画面，之后移动的动画都在原画面基础上做修改
        private Bitmap InitBmp;
        private Bitmap TempBmp;
        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            float x = e.X;
            float y = e.Y;
            if (MouseDowned)
            {
                label8.Visible = false;
                if (FirstMove)
                {
                    InitBmp = new Bitmap(pictureBox1.Image);
                    FirstMove = false;
                }
                if (TempBmp != null)
                {
                    TempBmp.Dispose();
                }
                TempBmp = new Bitmap(InitBmp);
                ghp = Graphics.FromImage(TempBmp);
                myPen.Color = Color.Black;
                myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                ghp.DrawLine(myPen, InitX, 1, x, 1);
                ghp.DrawLine(myPen, InitX, pictureBox1.Height - 1, x, pictureBox1.Height - 1);
                ghp.DrawLine(myPen, x, 0, x, (float)pictureBox1.Height);
                ghp.DrawString(Math.Round(((x - (float)pictureBox1.Width / 2) / (float)pictureBox1.Width * 2 * (LENGTH))).ToString(),
                        new System.Drawing.Font("宋体", 8, FontStyle.Bold), new SolidBrush(Color.Black), x > pictureBox1.Width / 2 ? x - 16 : x + 10, 
                        y > pictureBox1.Height / 2 ? y - 12 : y + 16);
                pictureBox1.Image = TempBmp;
            }
            else
            {
                if (CalComp)
                {
                    if (!ScaleFlag)
                    {
                        label8.Text = Math.Round(((x - (float)pictureBox1.Width / 2) / (float)pictureBox1.Width * 2 * (LENGTH))).ToString();
                        label8.Left = (int)(x > pictureBox1.Width / 2 ? x - 16 : x + 10);
                        label8.Top = (int)(y > pictureBox1.Height / 2 ? y - 12 : y + 16);
                        label8.Visible = true;
                    }
                    else
                    {
                        label8.Text = Math.Round((x / (float)pictureBox1.Width * (ceil - floor) + floor - (LENGTH + 1 - 1.5))).ToString();
                        label8.Left = (int)(x > pictureBox1.Width / 2 ? x - 16 : x + 10);
                        label8.Top = (int)(y > pictureBox1.Height / 2 ? y - 12 : y + 16);
                        label8.Visible = true;
                    }
                }
            }
        }

        private bool ScaleFlag = false;         //曲线图是否已经在放大模式下，若在放大模式下点击鼠标回则转到原始尺寸
        private int ceil = 0;                   //放大曲线图的X轴上限
        private int floor = 0;                  //放大曲线图的X轴下限
        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (MouseDowned)
            {
                MouseDowned = false;
                FirstMove = true;
                float x = e.X;

                ceil = (int)Math.Ceiling((double)(e.X > InitX ? e.X : InitX) / pictureBox1.Width * (2 * (LENGTH + 1) - 1));
                if (ceil > 2 * (LENGTH + 1) - 1)
                {
                    ceil = 2 * (LENGTH + 1) - 1;
                }
                floor = (int)Math.Floor((double)(e.X < InitX ? e.X : InitX) / pictureBox1.Width * (2 * (LENGTH + 1) - 1));
                if (floor < 0)
                {
                    floor = 0;
                }
                if (ceil - floor > 1)
                {
                    int width = ceil - floor;
                    ghp = Graphics.FromImage(bmp);
                    ghp.Clear(pictureBox1.BackColor);
                    myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
                    myPen.Color = Color.DarkGreen;
                    ghp.DrawLine(myPen, 0, (float)pictureBox1.Height / 2, pictureBox1.Width, (float)pictureBox1.Height / 2);
                    if (floor <= (LENGTH + 1) - 1 && (LENGTH + 1) - 1 <= ceil)
                    {
                        myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                        ghp.DrawLine(myPen, (float)((LENGTH ) - 1.5 - floor) / (float)width * (float)pictureBox1.Width, 0, (float)((LENGTH + 1) - 1.5 - floor) / (float)width * (float)pictureBox1.Width, pictureBox1.Height);
                        myPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
                    }
                    ghp.DrawString("单位：us", new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red), 1, 1);
                    for (int i = 1; i < 10; i++)
                    {
                        ghp.DrawLine(myPen, (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 + 5,
                            (float)pictureBox1.Width * i / 10, (float)pictureBox1.Height / 2 - 5);
                        if (!string.IsNullOrEmpty(label4.Text))
                        {
                            ghp.DrawString((scale * ((float)(floor - (LENGTH + 1) + 1.5) + i * (float)(width) / 10)).ToString("0.000"), new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),
                                (float)pictureBox1.Width * i / 10 - 20, (float)pictureBox1.Height / 2 + 8);
                        }
                    }
                    myPen.Color = Color.BlueViolet;
                    for (int i = -0; i < width - 1; i++)
                    {
                        ghp.DrawLine(myPen, (float)(i - 0.5) / (float)(width) * (float)pictureBox1.Width, (float)(1.05 * Max - result[floor + i]) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height,
                            (float)(i + 0.5) / (float)(width) * (float)pictureBox1.Width, (float)(1.05 * Max - result[floor + i + 1]) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height);
                    }
                    if (MaxNum != -1 && MaxNum > floor && MaxNum < ceil)
                    {
                        //myPen = new Pen(Color.Red);
                        ghp.FillEllipse(new SolidBrush(Color.Red), (float)(MaxNum - floor -0.5) / (width) * pictureBox1.Width - 4, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height - 4, 8, 8);
                        //ghp.DrawEllipse(myPen, (float)MaxNum / int.Parse(Len) * pictureBox1.Width, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height, 2, 2);
                        ghp.DrawString("采样点序号差：" + ((MaxNum - (LENGTH - 1))).ToString(), new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),
                            (float)(MaxNum - floor - 0.5) / (width) * pictureBox1.Width + 8, (float)(1.05 * Max - MaxMinus * Max) / 2 / (float)(1.05 * Max) * (float)pictureBox1.Height + MaxMinus * 8);
                    }
                    pictureBox1.Image = bmp;
                    ScaleFlag = true;
                }
                else
                {
                    this.Invoke(updateUI);
                }
            }
            
        }

        //存到.doc文件
        private void SaveDoc(object strFileName)
        {
            Microsoft.Office.Interop.Word._Application WordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
            _Document WordDoc;
            string strContent = "";
            Object oMissing = System.Reflection.Missing.Value;
            WordDoc = WordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            pictureBox1.Image.Save(strFileName.ToString().Substring(0, strFileName.ToString().Length - 4) + "bmp.bmp");    //先保存到文件，再把文件保存到Doc，最后删除文件
            Microsoft.Office.Interop.Word.WdGoToItem newLine = WdGoToItem.wdGoToLine;
            #region
            strContent = "日期时间：\t" + DateTime.Now.ToString() + "\r\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;
            WordApp.Selection.GoToNext(newLine);
            strContent = "文件名称：\t" + label2.Text + "\r\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;
            WordApp.Selection.GoToNext(newLine);
            strContent = "有效计算点数：\t" + textBox1.Text + "\r\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;
            WordApp.Selection.GoToNext(newLine);
            strContent = "\r\n";
            WordDoc.Paragraphs.Last.Range.Text = strContent;
            WordApp.Selection.GoToNext(newLine);
            string FileName = strFileName.ToString().Substring(0, strFileName.ToString().Length - 4) + "bmp.bmp";
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = WordDoc.Application.Selection.Range;
            WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
            WordDoc.Application.ActiveDocument.InlineShapes[1].Width = 400f;//图片宽度
            WordDoc.Application.ActiveDocument.InlineShapes[1].Height =400f/pictureBox1.Width * pictureBox1.Height;//图片高度
            #endregion
            System.IO.File.Delete(strFileName.ToString().Substring(0, strFileName.ToString().Length - 4) + "bmp.bmp");
            //将WordDoc文档对象的内容保存为DOC文档   
            WordDoc.SaveAs(ref strFileName, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref   oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
            //关闭WordDoc文档对象   
            WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
            //关闭WordApp组件对象   
            WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            MessageBox.Show("Word已存储！");
        }

        //判断按键按下时间，短按存成.bmp文件;长按.doc文件
        private DateTime dt0;
        private void button2_MouseDown(object sender, MouseEventArgs e)
        {
            dt0 = DateTime.Now;
        }

        private void button2_MouseUp(object sender, MouseEventArgs e)
        {
            TimeSpan ts = DateTime.Now - dt0;
            if(ts.TotalMilliseconds < 1000)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.InitialDirectory = @"C:\";
                sfd.Filter = "bmp文件(*.bmp)|*.bmp";
                sfd.FileName = label2.Text.Substring(0, label2.Text.Length - 4) + ".bmp";
                if (sfd.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(sfd.FileName))
                {
                    //pictureBox1.Image.Save(sfd.FileName);
                    Image imgForSave = pictureBox1.Image;
                    Bitmap bmpForSave = new Bitmap(imgForSave);
                    Graphics ghpForSave = Graphics.FromImage(imgForSave);
                    Pen myPenForSave = new Pen(Color.DarkRed);

                    ghp.DrawString(this.RShow.Text, new System.Drawing.Font("宋体", 10), new SolidBrush(Color.Red),1, pictureBox1.Height-15);
                    imgForSave = bmp;
                    imgForSave.Save(sfd.FileName);
                }
            }
            else
            {
                SaveFileDialog sfd = new SaveFileDialog(); 
                sfd.InitialDirectory = @"C:\";
                sfd.Filter = "doc文件(*.doc)|*.doc";
                sfd.FileName = label2.Text.Substring(0, label2.Text.Length - 4) + ".doc";
                if (sfd.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(sfd.FileName))
                {
                    SaveDoc(sfd.FileName);
                }
            }
        }
    }
}
