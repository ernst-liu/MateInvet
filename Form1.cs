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
        //�̵��������ַ���
        public String OperName;
        //���������ַ���
        public String FilterStr = "";
        //�ʶ��ı�������textΪ�ʶ��ı�����
        public static void Speak(string text)
        {
            SpeechSynthesizer speech = new SpeechSynthesizer();
            speech.Rate = int.Parse("-1");//����  ����-10��10֮��
            speech.Speak(text);
        }
        //�򿪵����ļ��������
        private void button9_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }
        //����excel������������
        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                String strPath = this.openFileDialog1.FileName;
                string strCon = "provider=microsoft.jet.oledb.4.0;data source=" + strPath + ";extended properties=excel 8.0";//�ؼ��Ǻ�ɫ����
                string strCon2007 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                //�����ӿ��Բ���.xls��.xlsx�ļ� (֧��Excel2003 �� Excel2007 �������ַ���) 
                //"HDR=yes;"��˵Excel�ļ��ĵ�һ������������������"HDR=No;"������ǰ����෴��"IMEX=1 "������е��������Ͳ�һ�£�ʹ��"IMEX=1"�ɱ����������ͳ�ͻ�� 
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
                //toolStripStatusLabel1.Text = ex.Message;
                String strPath = this.openFileDialog1.FileName;
                String strCon2007 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                OleDbConnection Con = new OleDbConnection(strCon2007);//��������
                string strSql = "select * from [FWO610310$]";//������д��ҲӦע�ⲻͬ����Ӧ��excel��Ϊsheet1��������Ҫ��������Ԫ����$������������
                OleDbCommand Cmd = new OleDbCommand(strSql, Con);//����Ҫִ�е�����
                OleDbDataAdapter da = new OleDbDataAdapter(Cmd);//��������������
                DataSet ds = new DataSet();//�½����ݼ�
                da.Fill(ds, "shyman");//�������������е����ݶ������ݼ��е�һ�����У��˴�����Ϊshyman��������ȡ��������ָ��datagridview1������ԴΪ���ݼ�ds�ĵ�һ�ű�Ҳ����shyman����Ҳ����дds.Table["shyman"]
                dataGridView1.DataSource = ds.Tables[0];
            }
        }
        //�ʲ����������ʼƥ��
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            this.toolStripStatusLabel1.Text = "";
            DataTable sourceTable = (DataTable)dataGridView1.DataSource;//��ȡ����Դ
            sourceTable.DefaultView.RowFilter = FilterStr; //����ǰ�ڲ�ѯ������ù�������
            DataTable rentTable = sourceTable.DefaultView.ToTable(); //��ȡɸѡ������
            int r;
            int icount = rentTable.Rows.Count;
            int icomplete = 0;
            DataView dv = rentTable.Copy().DefaultView;
            dv.RowFilter = "�ʲ��̵㵥״̬='���̵�'";
            icomplete = dv.ToTable().Rows.Count;
            //rentTable.DefaultView.RowFilter= FilterStr;
            if (rentTable.Rows.Count <= 0) return;
            if (int.TryParse(textBox3.Text, out r))
            {
                if (textBox3.Text.ToString().Length != 9)
                    return;
                for (int i = 0; i < rentTable.Rows.Count; i++)
                {
                    if (rentTable.Rows[i]["���ʱ��"].ToString() == textBox3.Text)
                    {
                        dataGridView1.Rows[i].Selected = true;
                        dataGridView1.FirstDisplayedScrollingRowIndex = i;
                        this.label1.Text = "���ʱ��� " + this.textBox3.Text + " �ʲ����� " + rentTable.Rows[i]["�ʲ�����"].ToString() + "ʹ����Ϊ" + rentTable.Rows[i]["ʹ����Ա"].ToString();
                        this.textBox3.Text = "";
                        dataGridView1.Rows[i].Cells["�ʲ��̵㵥״̬"].Value = "���̵�";
                        dataGridView1.Rows[i].Cells["�̵���"].Value = "����";
                        dataGridView1.Rows[i].Cells["�Ǽ���"].Value = OperName;
                        dataGridView1.Rows[i].Cells["�Ǽ�����"].Value = DateTime.Now.ToString("G");
                        if (this.checkBox1.Checked == true)
                        Speak(this.label1.Text);
                        icomplete = icomplete + 1;
                        this.toolStripStatusLabel1.Text = "�̵���� " + icomplete.ToString() + " / " + icount.ToString();
                        return;
                    }
                    dataGridView1.Rows[i].Selected = false;
                }
                this.label1.Text = "δ�ҵ�" + this.textBox3.Text+"";
                if (this.checkBox1.Checked == true)
                    Speak(this.label1.Text+"��������ӯԭ��!");
                string splusreason;
                InputDialog.Show("��¼����ӯԭ��", out splusreason);
                if (splusreason != "") 
                {
                    splusreason = "��" + splusreason + "��";
                    DataRow drinsert = ((DataTable)dataGridView1.DataSource).NewRow(); 
                    drinsert["���ʱ��"] = this.textBox3.Text;
                    drinsert["�Ǽ���"] = OperName;
                    drinsert["�ʲ��̵㵥״̬"] = "���̵�";
                    drinsert["�Ǽ�����"]  = DateTime.Now.ToString("G");
                    drinsert["�̵���"] = "��ӯ" + splusreason;
                    ((DataTable)dataGridView1.DataSource).Rows.Add(drinsert);
                    //this.dataGridView1.Rows[dataGridView1.RowCount-1].Selected = true;
                }

                this.textBox3.Text = "";
            }
            else
            {
                this.toolStripStatusLabel1.Text = "���������ʱ�ţ�";
                toolStripStatusLabel1.Text = "�̵���� " + icomplete.ToString() + " / " + icount.ToString();
            }
        }
        //��ʼ�̵�����ͣ�̵㰴ť
        private void button6_Click(object sender, EventArgs e)
        {
            if (this.dataGridView1.DataSource != null)
            {
                if (button6.Text == "��ʼ�̵�") //�̵��ʲ��߼�
                {
                    button6.Text = "��ͣ�̵�";
                    this.textBox3.Visible = true;
                    this.label2.Visible = true;
                    InputDialog.Show("��¼���̵�������", out OperName);
                    DataTable rentTable = (DataTable)dataGridView1.DataSource;//��ȡ����Դ
                    for (int i = 0; i < rentTable.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Cells["�̵���"].Value = "�̿���";
                        dataGridView1.Rows[i].Cells["�Ǽ���"].Value = OperName;
                        dataGridView1.Rows[i].Cells["�Ǽ�����"].Value = DateTime.Now.ToString("G");
                    }
                }
                else if (button6.Text == "��ͣ�̵�")
                {
                    button6.Text = "�����̵�";
                    this.textBox3.Visible = false;
                    this.label2.Visible = false;
                }
                else if (button6.Text == "�����̵�")
                {
                    InputDialog.Show("��¼���̵�������", out OperName);
                    this.textBox3.Visible = true;
                    this.label2.Visible = true;
                    button6.Text = "��ͣ�̵�";
                }
            }
            else
                MessageBox.Show("�̵��Ϊ�գ��뵼���̵����ٴε��");
        }
        //�����̵㲢�����̵��
        private void button8_Click(object sender, EventArgs e)
        {
            //ͳ���̵�������Ϣ
            DataTable sourceTable = (DataTable)dataGridView1.DataSource;//��ȡ����Դ
            sourceTable.DefaultView.RowFilter = FilterStr; //����ǰ�ڲ�ѯ������ù�������
            DataTable rentTable = sourceTable.DefaultView.ToTable(); //��ȡɸѡ������
            int icount = rentTable.Rows.Count;            
            //ͳ�����̵���
            DataView dv = rentTable.Copy().DefaultView;
            dv.RowFilter = "�ʲ��̵㵥״̬='���̵�'";
            int icomplete = dv.ToTable().Rows.Count;
            //ͳ����ӯ��
            dv.RowFilter = "�̵���  like '%��ӯ%'";
            int iplus = dv.ToTable().Rows.Count;
            //������ӯԭ��
            string sPlusreason = "����  ��һ������ӯ��ԭ��˵�����£�";
            sPlusreason = sPlusreason+ GroupByDt(dv.ToTable());
            //ͳ���̿�������
            dv.RowFilter = "�̵���='����'";
            int inormal = dv.ToTable().Rows.Count;            
            //ͳ���̿���
            dv.RowFilter = "�̵���  like '%�̿�%'";
            int iminus = dv.ToTable().Rows.Count;

            //�����̿�ԭ��
            string sMinreason = " ���������̿���ԭ��˵�����£�";
            sMinreason = sMinreason + GroupByDt(dv.ToTable());
            
            string ReslutStr = String.Format("���ε���{0}�����ʼ�¼���̵���{1}���ʲ��������̵�����{2}������ӯ{3}�����̿�{4}����", icount,icomplete,inormal,iplus,iminus);
            ReslutStr = ReslutStr + sPlusreason + sMinreason;
            MessageBox.Show(ReslutStr);
            DataRow drinsert = ((DataTable)dataGridView1.DataSource).NewRow();
            drinsert[0] = ReslutStr;
            ((DataTable)dataGridView1.DataSource).Rows.Add(drinsert);
            this.dataGridView1.Rows[dataGridView1.RowCount-1].Selected = true;
            if (dataGridView1.RowCount > 0)
            {
                ExportDataToExcel(this.dataGridView1.DataSource as DataTable, "�̵�����ϸ��" + System.DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            }
            else MessageBox.Show("������Ϊ�գ��޷�������");
        }

        //��Datatable ���ݽ��з���ϲ��ַ����������ַ�����
        public string GroupByDt(DataTable dt)
        {
            string resultstr = "";
            int icount = 0;
            DataView dv = dt.DefaultView;
            dv.Sort = "�̵���";
            dt =dv.ToTable();

            for (int i = 0; i < dt.Rows.Count;)
            {
                  string name = dt.Rows[i]["�̵���"].ToString();
                  icount = 0;
                //�ڲ�Ҳ��ѭ��ͬһ������������ͬ��nameʱ�����ڲ�ѭ��
                for (; i < dt.Rows.Count;)
                {
                    if (name == dt.Rows[i]["�̵���"].ToString())
                    {
                        string nametemp = name;
                        if (icount == 0)
                            resultstr = resultstr + ";"+name ;// + nametemp.Replace("��ӯ��", "").Replace("�̿���",""); 
                        icount++;
                          i++;
                        if (i == dt.Rows.Count) resultstr = resultstr + " " + icount.ToString() + "�� ";
                    }
                    else
                    {
                        resultstr = resultstr + " "+ icount.ToString() + "�� ";
                        break;
                    }
                }
               
            }
            return resultstr;


        }
        //��datatable ����Ϊexcel����ʾ�Ƿ��
        //ʹ����NPOI�������װ
        public void ExportDataToExcel(DataTable TableName, string FileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //�����ļ�����
            saveFileDialog.Title = "����Excel�ļ�";
            //�����ļ�����
            saveFileDialog.Filter = "Excel ������(*.xlsx)|*.xlsx|Excel 97-2003 ������(*.xls)|*.xls";
            //����Ĭ���ļ�������ʾ˳��  
            saveFileDialog.FilterIndex = 1;
            //�Ƿ��Զ����ļ����������չ��
            saveFileDialog.AddExtension = true;
            //�Ƿ�����ϴδ򿪵�Ŀ¼
            saveFileDialog.RestoreDirectory = true;
            //����Ĭ���ļ���
            saveFileDialog.FileName = FileName;
            //����ȷ��ѡ��İ�ť  
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //����ļ�·�� 
                string localFilePath = saveFileDialog.FileName.ToString();
                //���ݳ�ʼ��
                int TotalCount;     //������
                int RowRead = 0;    //�Ѷ�����
                int Percent = 0;    //�ٷֱ�
                TotalCount = TableName.Rows.Count;
                toolStripStatusLabel1.Text = "����" + TotalCount + "������";
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
                //����
                Stopwatch timer = new Stopwatch();
                timer.Start();
                try
                {
                    //��ȡ����  
                    IRow rowHeader = sheet.CreateRow(0);
                    for (int i = 0; i < TableName.Columns.Count; i++)
                    {
                        ICell cell = rowHeader.CreateCell(i);
                        cell.SetCellValue(TableName.Columns[i].ColumnName);
                    }

                    //��ȡ����  
                    for (int i = 0; i < TableName.Rows.Count; i++)
                    {
                        IRow rowData = sheet.CreateRow(i + 1);
                        for (int j = 0; j < TableName.Columns.Count; j++)
                        {
                            ICell cell = rowData.CreateCell(j);
                            cell.SetCellValue(TableName.Rows[i][j].ToString());
                        }
                        //״̬����ʾ
                        RowRead++;
                        Percent = (int)(100 * RowRead / TotalCount);
                        toolStripProgressBar1.Maximum = TotalCount;
                        toolStripProgressBar1.Value = RowRead;
                        toolStripStatusLabel1.Text = "����" + TotalCount + "�����ݣ��Ѷ�ȡ" + Percent.ToString() + "%�����ݡ�";
                        Application.DoEvents();
                    }

                    //״̬������
                    //   toolStripStatusLabel1.Text = "��������Excel...";
                    Application.DoEvents();

                    //תΪ�ֽ�����  
                    MemoryStream stream = new MemoryStream();
                    workbook.Write(stream);
                    var buf = stream.ToArray();

                    //����ΪExcel�ļ�  
                    using (FileStream fs = new FileStream(localFilePath, FileMode.Create, FileAccess.Write))
                    {
                        fs.Write(buf, 0, buf.Length);
                        fs.Flush();
                        fs.Close();
                    }

                    //״̬������
                    toolStripStatusLabel1.Text = "����Excel�ɹ�������ʱ" + timer.ElapsedMilliseconds + "���롣";
                    Application.DoEvents();

                    //�ر�����
                    timer.Reset();
                    timer.Stop();

                    //�ɹ���ʾ
                    if (MessageBox.Show("�����ɹ����Ƿ������򿪣�", "��ʾ", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    {
                        ProcessStartInfo psi = new ProcessStartInfo(localFilePath);
                        psi.UseShellExecute = true;
                        Process.Start(psi);
                    
                    }

                    //����ʼֵ
                    toolStripStatusLabel1.Visible = false;
                    toolStripProgressBar1.Visible = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "��ʾ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    //�ر�����
                    timer.Reset();
                    timer.Stop();
                    //����ʼֵ
                    toolStripStatusLabel1.Visible = false;
                    toolStripProgressBar1.Visible = false;
                }
            }
        }

        //ִ��ɸѡ����datatable����ɸѡ������ȫ��ɸѡ����
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            string filtername = this.comboBox1.Text;
            string filterstr = this.textBox1.Text;
            string filtersqlstr = "";
            if (dt != null)
            {
                if (filtername == "δɸѡ")
                {
                    dt.DefaultView.RowFilter = "";
                    FilterStr = "";
                    this.toolStripStatusLabel1.Text = "�ܼ�" + dt.DefaultView.ToTable().Rows.Count.ToString() + " ��";
                }
                else
                {
                    filtersqlstr = filtername + "  like '%" + filterstr + "%'";
                    if (this.comboBox2.Text != "�����̵�״̬")
                    {
                        filtersqlstr = filtersqlstr + " and  �ʲ��̵㵥״̬ like '%" + this.comboBox2.Text + "%'";
                    }
                    dt.DefaultView.RowFilter = filtersqlstr;
                    FilterStr = filtersqlstr;

                    this.toolStripStatusLabel1.Text = "ɸѡ����ܼ�" + dt.DefaultView.ToTable().Rows.Count.ToString() + " ��" ;

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