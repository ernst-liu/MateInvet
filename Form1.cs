using System.Data;
using System.Speech.Synthesis;
using System.Diagnostics;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace MateInvet
{
    using System.Data.OleDb;
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //盘点人名称字符串
        public String OperName;
        //过滤条件字符串
        public String FilterStr = "";
        //朗读文本方法，text为朗读文本内容
        public static void Speak(string text)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Rate = int.Parse("-1");//语速  介于-10于10之间
            speech.Speak(text);
        }
        //打开导入文件浏览窗口
        private void button9_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
        //导入excel表数据至工具
        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                String strPath = this.openFileDialog1.FileName;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//关键是红色区域
                string strCon2007 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串) 
                //"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                string strSql = "select * from [FWO610310$]";//表名的写法也应注意不同，对应的excel表为sheet1，在这里要在其后加美元符号$，并用中括号
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//建立要执行的命令
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//建立数据适配器
                DataSet ds = new DataSet();//新建数据集
                da.Fill(ds, "shyman");//把数据适配器中的数据读到数据集中的一个表中（此处表名为shyman，可以任取表名）
                                      //指定datagridview1的数据源为数据集ds的第一张表（也就是shyman表），也可以写ds.Table["shyman"]
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                //toolStripStatusLabel1.Text = ex.Message;
                String strPath = this.openFileDialog1.FileName;
                String strCon2007 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                OleDbConnection Con = new OleDbConnection(strCon2007);//建立连接
                string strSql = "select * from [FWO610310$]";//表名的写法也应注意不同，对应的excel表为sheet1，在这里要在其后加美元符号$，并用中括号
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//建立要执行的命令
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//建立数据适配器
                DataSet ds = new DataSet();//新建数据集
                da.Fill(ds, "shyman");//把数据适配器中的数据读到数据集中的一个表中（此处表名为shyman，可以任取表名），指定datagridview1的数据源为数据集ds的第一张表（也就是shyman表），也可以写ds.Table["shyman"]
                dataGridView1.DataSource = ds.Tables[0];
            }
        }
        //资产编码输入后开始匹配
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "";
            DataTable sourceTable = (DataTable)dataGridView1.DataSource;//获取数据源
            sourceTable.DefaultView.RowFilter = FilterStr; //根据前期查询情况设置过滤条件
            DataTable rentTable = sourceTable.DefaultView.ToTable(); //获取筛选后结果表
            int r;
            int icount = rentTable.Rows.Count;
            int icomplete = 0;
            DataView dv = rentTable.Copy().DefaultView;
            dv.RowFilter = "资产盘点单状态='已盘点'";
            icomplete = dv.ToTable().Rows.Count;
            //rentTable.DefaultView.RowFilter= FilterStr;
            if (rentTable.Rows.Count <= 0) return;
            if (int.TryParse(textBox3.Text, out r))
            {
                if (textBox3.Text.ToString().Length != 9)
                    return;
                for (int i = 0; i < rentTable.Rows.Count; i++)
                {
                    if (rentTable.Rows[i]["物资编号"].ToString() == textBox3.Text)
                    {
                        dataGridView1.Rows[i].Selected = true;
                        dataGridView1.FirstDisplayedScrollingRowIndex = i;
                        this.label1.Text = "物资编码 " + this.textBox3.Text + " 资产名称 " + rentTable.Rows[i]["资产名称"].ToString() + "使用人为" + rentTable.Rows[i]["使用人员"].ToString();
                        this.textBox3.Text = "";
                        dataGridView1.Rows[i].Cells["资产盘点单状态"].Value = "已盘点";
                        dataGridView1.Rows[i].Cells["盘点结果"].Value = "正常";
                        dataGridView1.Rows[i].Cells["登记人"].Value = OperName;
                        dataGridView1.Rows[i].Cells["登记日期"].Value = DateTime.Now.ToString("G");
                        if (this.checkBox1.Checked == true)
                        Speak(this.label1.Text);
                        icomplete = icomplete + 1;
                        this.toolStripStatusLabel1.Text = "盘点进度 " + icomplete.ToString() + " / " + icount.ToString();
                        return;
                    }
                    dataGridView1.Rows[i].Selected = false;
                }
                this.label1.Text = "未找到" + this.textBox3.Text+"";
                if (this.checkBox1.Checked == true)
                    Speak(this.label1.Text+"请输入盘盈原因!");
                string splusreason;
                InputDialog.Show("请录入盘盈原因", out splusreason);
                if (splusreason != "") 
                {
                    splusreason = "，" + splusreason + "。";
                    DataRow drinsert = ((DataTable)dataGridView1.DataSource).NewRow(); 
                    drinsert["物资编号"] = this.textBox3.Text;
                    drinsert["登记人"] = OperName;
                    drinsert["资产盘点单状态"] = "已盘点";
                    drinsert["登记日期"]  = DateTime.Now.ToString("G");
                    drinsert["盘点结果"] = "盘盈" + splusreason;
                    ((DataTable)dataGridView1.DataSource).Rows.Add(drinsert);
                    //this.dataGridView1.Rows[dataGridView1.RowCount-1].Selected = true;
                }

                this.textBox3.Text = "";
            }
            else
            {
                this.toolStripStatusLabel1.Text = "请输入物资编号！";
                toolStripStatusLabel1.Text = "盘点进度 " + icomplete.ToString() + " / " + icount.ToString();
            }
        }
        //开始盘点与暂停盘点按钮
        private void button6_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.DataSource != null)
            {
                if (button6.Text == "开始盘点") //盘点资产逻辑
                {
                    button6.Text = "暂停盘点";
                    this.textBox3.Visible = true;
                    this.label2.Visible = true;
                    InputDialog.Show("请录入盘点人姓名", out OperName);
                    DataTable rentTable = (DataTable)dataGridView1.DataSource;//获取数据源
                    for (int i = 0; i < rentTable.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Cells["盘点结果"].Value = "盘亏，";
                        dataGridView1.Rows[i].Cells["登记人"].Value = OperName;
                        dataGridView1.Rows[i].Cells["登记日期"].Value = DateTime.Now.ToString("G");
                    }
                }
                else if (button6.Text == "暂停盘点")
                {
                    button6.Text = "继续盘点";
                    this.textBox3.Visible = false;
                    this.label2.Visible = false;
                }
                else if (button6.Text == "继续盘点")
                {
                    InputDialog.Show("请录入盘点人姓名", out OperName);
                    this.textBox3.Visible = true;
                    this.label2.Visible = true;
                    button6.Text = "暂停盘点";
                }
            }
            else
                MessageBox.Show("盘点表为空，请导入盘点表后再次点击");
        }
        //结束盘点并导出盘点表
        private void button8_Click(object sender, EventArgs e)
        {
            //统计盘点总数信息
            DataTable sourceTable = (DataTable)dataGridView1.DataSource;//获取数据源
            sourceTable.DefaultView.RowFilter = FilterStr; //根据前期查询情况设置过滤条件
            DataTable rentTable = sourceTable.DefaultView.ToTable(); //获取筛选后结果表
            int icount = rentTable.Rows.Count;            
            //统计已盘点数
            DataView dv = rentTable.Copy().DefaultView;
            dv.RowFilter = "资产盘点单状态='已盘点'";
            int icomplete = dv.ToTable().Rows.Count;
            //统计盘盈数
            dv.RowFilter = "盘点结果  like '%盘盈%'";
            int iplus = dv.ToTable().Rows.Count;
            //汇总盘盈原因
            string sPlusreason = "其中  【一】、盘盈的原因说明如下：";
            sPlusreason = sPlusreason+ GroupByDt(dv.ToTable());
            //统计盘库正常数
            dv.RowFilter = "盘点结果='正常'";
            int inormal = dv.ToTable().Rows.Count;            
            //统计盘亏数
            dv.RowFilter = "盘点结果  like '%盘亏%'";
            int iminus = dv.ToTable().Rows.Count;

            //汇总盘亏原因
            string sMinreason = " 【二】、盘亏的原因说明如下：";
            sMinreason = sMinreason + GroupByDt(dv.ToTable());
            
            string ReslutStr = String.Format("本次导入{0}条物资记录，盘点了{1}个资产，其中盘点正常{2}个、盘盈{3}个、盘亏{4}个。", icount,icomplete,inormal,iplus,iminus);
            ReslutStr = ReslutStr + sPlusreason + sMinreason;
            MessageBox.Show(ReslutStr);
            DataRow drinsert = ((DataTable)dataGridView1.DataSource).NewRow();
            drinsert[0] = ReslutStr;
            ((DataTable)dataGridView1.DataSource).Rows.Add(drinsert);
            this.dataGridView1.Rows[dataGridView1.RowCount-1].Selected = true;
            if (dataGridView1.RowCount > 0)
            {
                ExportDataToExcel(this.dataGridView1.DataSource as DataTable, "盘点结果明细表" + System.DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            }
            else MessageBox.Show("表数据为空，无法导出！");
        }

        //将Datatable 数据进行分组合并字符串，返回字符串。
        public string GroupByDt(DataTable dt)
        {
            string resultstr = "";
            int icount = 0;
            DataView dv = dt.DefaultView;
            dv.Sort = "盘点结果";
            dt =dv.ToTable();

            for (int i = 0; i < dt.Rows.Count;)
            {
                  string name = dt.Rows[i]["盘点结果"].ToString();
                  icount = 0;
                //内层也是循环同一个表，当遇到不同的name时跳出内层循环
                for (; i < dt.Rows.Count;)
                {
                    if (name == dt.Rows[i]["盘点结果"].ToString())
                    {
                        string nametemp = name;
                        if (icount == 0)
                            resultstr = resultstr + ";"+name ;// + nametemp.Replace("盘盈，", "").Replace("盘亏，",""); 
                        icount++;
                          i++;
                        if (i == dt.Rows.Count) resultstr = resultstr + " " + icount.ToString() + "个 ";
                    }
                    else
                    {
                        resultstr = resultstr + " "+ icount.ToString() + "个 ";
                        break;
                    }
                }
               
            }
            return resultstr;


        }
        //将datatable 导出为excel并提示是否打开
        //使用了NPOI组件并安装
        public void ExportDataToExcel(DataTable TableName, string FileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //设置文件标题
            saveFileDialog.Title = "导出Excel文件";
            //设置文件类型
            saveFileDialog.Filter = "Excel 工作簿(*.xlsx)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
            //设置默认文件类型显示顺序  
            saveFileDialog.FilterIndex = 1;
            //是否自动在文件名中添加扩展名
            saveFileDialog.AddExtension = true;
            //是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            //设置默认文件名
            saveFileDialog.FileName = FileName;
            //按下确定选择的按钮  
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //获得文件路径 
                string localFilePath = saveFileDialog.FileName.ToString();
                //数据初始化
                int TotalCount;     //总行数
                int RowRead = 0;    //已读行数
                int Percent = 0;    //百分比
                TotalCount = TableName.Rows.Count;
                toolStripStatusLabel1.Text = "共有" + TotalCount + "条数据";
                toolStripStatusLabel1.Visible = true;
                toolStripProgressBar1.Visible = true;

                //NPOI
                IWorkbook workbook;
                string FileExt = Path.GetExtension(localFilePath).ToLower();
                if (FileExt == ".xlsx")
                {
                    workbook = new XSSFWorkbook();
                }
                else if (FileExt == ".xls")
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    workbook = null;
                }
                if (workbook == null)
                {
                    return;
                }
                ISheet sheet = string.IsNullOrEmpty(FileName) ? workbook.CreateSheet("Sheet1") : workbook.CreateSheet(FileName);
                //秒钟
                Stopwatch timer = new Stopwatch();
                timer.Start();
                try
                {
                    //读取标题  
                    IRow rowHeader = sheet.CreateRow(0);
                    for (int i = 0; i < TableName.Columns.Count; i++)
                    {
                        ICell cell = rowHeader.CreateCell(i);
                        cell.SetCellValue(TableName.Columns[i].ColumnName);
                    }

                    //读取数据  
                    for (int i = 0; i < TableName.Rows.Count; i++)
                    {
                        IRow rowData = sheet.CreateRow(i + 1);
                        for (int j = 0; j < TableName.Columns.Count; j++)
                        {
                            ICell cell = rowData.CreateCell(j);
                            cell.SetCellValue(TableName.Rows[i][j].ToString());
                        }
                        //状态栏显示
                        RowRead++;
                        Percent = (int)(100 * RowRead / TotalCount);
                        toolStripProgressBar1.Maximum = TotalCount;
                        toolStripProgressBar1.Value = RowRead;
                        toolStripStatusLabel1.Text = "共有" + TotalCount + "条数据，已读取" + Percent.ToString() + "%的数据。";
                        Application.DoEvents();
                    }

                    //状态栏更改
                    //   toolStripStatusLabel1.Text = "正在生成Excel...";
                    Application.DoEvents();

                    //转为字节数组  
                    MemoryStream stream = new MemoryStream();
                    workbook.Write(stream);
                    var buf = stream.ToArray();

                    //保存为Excel文件  
                    using (FileStream fs = new FileStream(localFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(buf, 0, buf.Length);
                        fs.Flush();
                        fs.Close();
                    }

                    //状态栏更改
                    toolStripStatusLabel1.Text = "生成Excel成功，共耗时" + timer.ElapsedMilliseconds + "毫秒。";
                    Application.DoEvents();

                    //关闭秒钟
                    timer.Reset();
                    timer.Stop();

                    //成功提示
                    if (MessageBox.Show("导出成功，是否立即打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        ProcessStartInfo psi = new ProcessStartInfo(localFilePath);
                        psi.UseShellExecute = true;
                        Process.Start(psi);
                    
                    }

                    //赋初始值
                    toolStripStatusLabel1.Visible = false;
                    toolStripProgressBar1.Visible = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    //关闭秒钟
                    timer.Reset();
                    timer.Stop();
                    //赋初始值
                    toolStripStatusLabel1.Visible = false;
                    toolStripProgressBar1.Visible = false;
                }
            }
        }

        //执行筛选，对datatable进行筛选并设置全局筛选条件
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            string filtername = this.comboBox1.Text;
            string filterstr = this.textBox1.Text;
            string filtersqlstr = "";
            if (dt != null)
            {
                if (filtername == "未筛选")
                {
                    dt.DefaultView.RowFilter = "";
                    FilterStr = "";
                    this.toolStripStatusLabel1.Text = "总计" + dt.DefaultView.ToTable().Rows.Count.ToString() + " 条";
                }
                else
                {
                    filtersqlstr = filtername + "  like '%" + filterstr + "%'";
                    if (this.comboBox2.Text != "所有盘点状态")
                    {
                        filtersqlstr = filtersqlstr + " and  资产盘点单状态 like '%" + this.comboBox2.Text + "%'";
                    }
                    dt.DefaultView.RowFilter = filtersqlstr;
                    FilterStr = filtersqlstr;

                    this.toolStripStatusLabel1.Text = "筛选结果总计" + dt.DefaultView.ToTable().Rows.Count.ToString() + " 条" ;

                }
                
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}