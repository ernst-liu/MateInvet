using System;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Speech.Synthesis;



namespace MateInvet
{
    using System.Data.OleDb;
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        //  text为输入文本内容
        public static void Speak(string text)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer(); 
            
                speech.Rate = int.Parse("-1");//语速  介于-10于10之间
                speech.Speak(text);
          
        }
        private void button9_Click(object sender, EventArgs e)
        {
                   this.openFileDialog1.ShowDialog();            
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            this.textBox1.Text = this.openFileDialog1.FileName;
            try
            {
                String strPath = this.textBox1.Text;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//关键是红色区域
               string strCon2007="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串) 
                //"HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
 
                OleDbConnection Con = new OleDbConnection(strCon2007);//建立连接
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
                MessageBox.Show(ex.Message);//捕捉异常
                String strPath = this.textBox1.Text;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//关键是红色区域

                OleDbConnection Con = new OleDbConnection(strCon);//建立连接
                string strSql = "select * from [FWO610310$]";//表名的写法也应注意不同，对应的excel表为sheet1，在这里要在其后加美元符号$，并用中括号
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//建立要执行的命令
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//建立数据适配器
                DataSet ds = new DataSet();//新建数据集
                da.Fill(ds, "shyman");//把数据适配器中的数据读到数据集中的一个表中（此处表名为shyman，可以任取表名）
                                      //指定datagridview1的数据源为数据集ds的第一张表（也就是shyman表），也可以写ds.Table["shyman"]

                dataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                String strPath = this.textBox1.Text;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//关键是红色区域
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
                MessageBox.Show(ex.Message);//捕捉异常
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            DataTable rentTable = (DataTable)dataGridView1.DataSource;//获取数据源
            
                int r;
                if (int.TryParse(textBox3.Text, out r))
                {
                    this.label2.Text="";
                    if (textBox3.Text.ToString().Length != 9)
                    return;
                    for (int i = 0; i < rentTable.Rows.Count; i++)
                    {
                        if (rentTable.Rows[i]["物资编号"].ToString() == textBox3.Text)
                        {
                            dataGridView1.Rows[i].Selected = true;
                            dataGridView1.FirstDisplayedScrollingRowIndex = i;
                            this.label1.Text ="物资编码 "+ this.textBox3.Text+" 资产名称 "+rentTable.Rows[i]["资产名称"].ToString()+"使用人为"+ rentTable.Rows[i]["使用人员"].ToString();
                            this.textBox3.Text = "";
                        dataGridView1.Rows[i].Cells["资产盘点单状态"].Value = "已盘点";
                             Speak(this.label1.Text);
                            return;
                        }
                    dataGridView1.Rows[i].Selected = false;
                }
                        this.label1.Text = "未找到物资编号！"+ this.textBox3.Text;
                         
                        Speak(this.label1.Text);
                         this.textBox3.Text = "";

            }
                else
                   this.label2.Text="请输入物资编号！";
          
        }
    }
}