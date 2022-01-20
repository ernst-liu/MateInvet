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
        //  textΪ�����ı�����
        public static void Speak(string text)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer(); 
            
                speech.Rate = int.Parse("-1");//����  ����-10��10֮��
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
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//�ؼ��Ǻ�ɫ����
               string strCon2007="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                //�����ӿ��Բ���.xls��.xlsx�ļ� (֧��Excel2003 �� Excel2007 �������ַ���) 
                //"HDR=yes;"��˵Excel�ļ��ĵ�һ������������������"HDR=No;"������ǰ����෴��"IMEX=1 "������е��������Ͳ�һ�£�ʹ��"IMEX=1"�ɱ����������ͳ�ͻ�� 
 
                OleDbConnection Con = new OleDbConnection(strCon2007);//��������
                string strSql = "select * from [FWO610310$]";//������д��ҲӦע�ⲻͬ����Ӧ��excel��Ϊsheet1��������Ҫ��������Ԫ����$������������
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//����Ҫִ�е�����
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//��������������
                DataSet ds = new DataSet();//�½����ݼ�
                da.Fill(ds, "shyman");//�������������е����ݶ������ݼ��е�һ�����У��˴�����Ϊshyman��������ȡ������
                                      //ָ��datagridview1������ԴΪ���ݼ�ds�ĵ�һ�ű�Ҳ����shyman����Ҳ����дds.Table["shyman"]

                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//��׽�쳣
                String strPath = this.textBox1.Text;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//�ؼ��Ǻ�ɫ����

                OleDbConnection Con = new OleDbConnection(strCon);//��������
                string strSql = "select * from [FWO610310$]";//������д��ҲӦע�ⲻͬ����Ӧ��excel��Ϊsheet1��������Ҫ��������Ԫ����$������������
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//����Ҫִ�е�����
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//��������������
                DataSet ds = new DataSet();//�½����ݼ�
                da.Fill(ds, "shyman");//�������������е����ݶ������ݼ��е�һ�����У��˴�����Ϊshyman��������ȡ������
                                      //ָ��datagridview1������ԴΪ���ݼ�ds�ĵ�һ�ű�Ҳ����shyman����Ҳ����дds.Table["shyman"]

                dataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                String strPath = this.textBox1.Text;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//�ؼ��Ǻ�ɫ����
                OleDbConnection Con = new OleDbConnection(strCon);//��������
                string strSql = "select * from [FWO610310$]";//������д��ҲӦע�ⲻͬ����Ӧ��excel��Ϊsheet1��������Ҫ��������Ԫ����$������������
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//����Ҫִ�е�����
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//��������������
                DataSet ds = new DataSet();//�½����ݼ�
                da.Fill(ds, "shyman");//�������������е����ݶ������ݼ��е�һ�����У��˴�����Ϊshyman��������ȡ������
                                      //ָ��datagridview1������ԴΪ���ݼ�ds�ĵ�һ�ű�Ҳ����shyman����Ҳ����дds.Table["shyman"]

                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);//��׽�쳣
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            DataTable rentTable = (DataTable)dataGridView1.DataSource;//��ȡ����Դ
            
                int r;
                if (int.TryParse(textBox3.Text, out r))
                {
                    this.label2.Text="";
                    if (textBox3.Text.ToString().Length != 9)
                    return;
                    for (int i = 0; i < rentTable.Rows.Count; i++)
                    {
                        if (rentTable.Rows[i]["���ʱ��"].ToString() == textBox3.Text)
                        {
                            dataGridView1.Rows[i].Selected = true;
                            dataGridView1.FirstDisplayedScrollingRowIndex = i;
                            this.label1.Text ="���ʱ��� "+ this.textBox3.Text+" �ʲ����� "+rentTable.Rows[i]["�ʲ�����"].ToString()+"ʹ����Ϊ"+ rentTable.Rows[i]["ʹ����Ա"].ToString();
                            this.textBox3.Text = "";
                        dataGridView1.Rows[i].Cells["�ʲ��̵㵥״̬"].Value = "���̵�";
                             Speak(this.label1.Text);
                            return;
                        }
                    dataGridView1.Rows[i].Selected = false;
                }
                        this.label1.Text = "δ�ҵ����ʱ�ţ�"+ this.textBox3.Text;
                         
                        Speak(this.label1.Text);
                         this.textBox3.Text = "";

            }
                else
                   this.label2.Text="���������ʱ�ţ�";
          
        }
    }
}