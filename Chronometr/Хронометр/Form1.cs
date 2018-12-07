using System;
using System.Collections.Generic;
using System.Threading;
using System.Timers;

using System.Linq;
using System.Text;
using System.Media;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.DirectX;
using Microsoft.DirectX.DirectInput;
using System.Data.OleDb;
using System.Data;
using System.IO;
//using Excel;
using Excel = Microsoft.Office.Interop.Excel;



namespace WindowsFormsApp2
{

    public partial class Form1 : Form
    {

        private static string newText;

        private static DateTime PreviousTime;
        private static DateTime NewTime;
        private static TimeSpan NewTimeSpan;
        private static string NameSor;


        private static bool first = false;
        private static bool second = false;
        private static bool third = false;
        private static bool fourth = false;
        private static int m=0, w=0, c=0,key=0;
        
        //private static string TextTime = "";
        //private static int TryName;
        

        public Form1()
        {

            InitializeComponent();

            InitializePlayer();
        }

        Device device;

        System.Timers.Timer timer = new System.Timers.Timer(1);

        SoundPlayer startsound, stopsound;

        


        private void InitializePlayer()
        {
            startsoundfilepath.Text = Properties.Settings.Default.StartSound;
            stopsoundfilepath.Text = Properties.Settings.Default.StopSound;

            startsound = new SoundPlayer();
            if (startsoundfilepath.Text != "")
            {
                startsound.SoundLocation = startsoundfilepath.Text;
                startsound.Load();
            }

            stopsound = new SoundPlayer();

            if (stopsoundfilepath.Text != "")
            {
                stopsound.SoundLocation = stopsoundfilepath.Text;
                stopsound.Load();
            }

            OpenExcel(Properties.Settings.Default.PathtoExcel);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            foreach (DeviceInstance instance in Manager.GetDevices(DeviceClass.GameControl, EnumDevicesFlags.AttachedOnly))
            {
                device = new Device(instance.ProductGuid);
                device.SetCooperativeLevel(null, CooperativeLevelFlags.Background | CooperativeLevelFlags.NonExclusive);

                if (instance.ProductName.Length > 0)
                {
                    label2.Text = "Устройство подключено";
                    StartButton.Enabled= true;
                    StopButton.Enabled = true;
                    //ResetButton.Enabled = true;
                }
                device.Acquire();
            }
            if (Properties.Settings.Default.Pol == 1)
            {
                //ViewLeaders(1);
                radioButtonWoman.Checked = true;
            }
            if (Properties.Settings.Default.Pol == 2)
            {
                //ViewLeaders(2);
                radioButtonMan.Checked = true;
            }
            if (Properties.Settings.Default.Pol == 3)
            {
                //ViewLeaders(3);
                radioButtonComands.Checked = true;
            }
            StopButton1.Enabled = false;
            button7.Enabled = false;
            button9.Enabled = false;
            button11.Enabled = false;

        }
        
        private void StartButton_Click(object sender, EventArgs e)
        {
            button5.Enabled = false;
            ToTimerButton.Enabled = false;
            ExportButton.Enabled = false;
            
            if (!timer.Enabled)
            {
                StartButton.Text = "Pause";

                
                //pictureBoxStart1.Visible = true;
                startsound.Play();
                Thread.Sleep(1000);
                //pictureBoxStart1.Visible = false;
                //pictureBoxStart2.Visible = true;
                startsound.Play();
                Thread.Sleep(1000);
                //pictureBoxStart2.Visible = false;
                //pictureBoxStart3.Visible = true;
                startsound.Play();
                Thread.Sleep(1000);
                StopButton1.Enabled = true;
                button7.Enabled = true;
                button9.Enabled = true;
                button11.Enabled = true;
                //pictureBoxStart3.Visible = false;
                PreviousTime = DateTime.Now.Subtract(NewTimeSpan);
                timer.Elapsed += timer_tick;
                timer.Enabled = true;
                
            }
            else
            {
                StartButton.Text = "Start";
                

                timer.Enabled = false;

            }
            



        }

        private void timer_tick(Object source, ElapsedEventArgs e)
        {
            
                NewTime = new DateTime(1, 1, 1, 0, 0, 0);

                NewTimeSpan = e.SignalTime - PreviousTime;
                
                NewTime = NewTime.Add(NewTimeSpan);
                newText = String.Format("{0:HH':'mm':'ss.ffffff}", NewTime);

            if (TextTimer.InvokeRequired) TextTimer.Invoke(new Action<string>((s) => TextTimer.Text = s), newText);
            else TextTimer.Text = newText;

            UpdateJoystickState(newText);

            if (first == false)
            {

                if (TextTimer1.InvokeRequired) TextTimer1.Invoke(new Action<string>((s) => TextTimer1.Text = s), newText);
                else TextTimer1.Text = newText;
            }
            if (second == false)
            {
                if (TextTimer2.InvokeRequired) TextTimer2.Invoke(new Action<string>((s) => TextTimer2.Text = s), newText);
                else TextTimer.Text = newText;
            }
            if (third == false)
            {
                if (TextTimer3.InvokeRequired) TextTimer3.Invoke(new Action<string>((s) => TextTimer3.Text = s), newText);
                else TextTimer.Text = newText;
            }
            if (fourth == false)
            {
                if (TextTimer4.InvokeRequired) TextTimer4.Invoke(new Action<string>((s) => TextTimer4.Text = s), newText);
                else TextTimer4.Text = newText;
            }

            
       
        }


        private void SetTime(string p,string name,string time)
        {
                        //string format = "hh:mm:ss.zzzzzz";
            if (p == "м" || p == "М")
            {
                if (dataGridView2.ColumnCount > 3)
                {
                    for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
                    {
                        DataGridViewTextBoxCell txtxCell = (DataGridViewTextBoxCell)dataGridView2.Rows[i].Cells[3];
                        if (txtxCell.Value != null)
                        {
                            if (name == txtxCell.Value.ToString())
                            {
                                dataGridView2.Rows[i].Cells[0].Value = false;
                                dataGridView2.Rows[i].Cells[10].Value = time;
                                
                            }
                        }
                    }
                    key = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        if (dataGridView2.Rows[i].Cells[10].Value != null)
                            if (dataGridView2.Rows[i].Cells[10].Value.ToString() != "")
                            {
                                dataGridView2.Rows[i].Cells[11].Value = Convert.ToString(key + 1);
                                key++;
                            }
                    }
                    if (radioButtonWoman.Checked == true)
                    {
                        ViewLeaders(1);
                    }
                    if (radioButtonMan.Checked == true)
                    {
                        ViewLeaders(2);
                    }
                    if (radioButtonComands.Checked == true)
                    {
                        ViewLeaders(3);
                    }
                }
            }
            if (p == "ж" || p == "ж")
            {
                if (dataGridView1.ColumnCount > 3)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        DataGridViewTextBoxCell txtxCell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[3];
                        if (txtxCell.Value != null)
                        {
                            if (name == txtxCell.Value.ToString())
                            {
                                dataGridView1.Rows[i].Cells[0].Value = false;
                                dataGridView1.Rows[i].Cells[10].Value = time;
                                
                            }
                        }
                    }
                    key = 0;
                    for (int i=0;i<dataGridView1.Rows.Count;i++)
                    {
                        if (dataGridView1.Rows[i].Cells[10].Value != null)
                            if(dataGridView1.Rows[i].Cells[10].Value.ToString() != "")
                            {
                                dataGridView1.Rows[i].Cells[11].Value = Convert.ToString(key + 1);
                                key++;
                            }
                    }
                    if (radioButtonWoman.Checked == true)
                    {
                        ViewLeaders(1);
                    }
                    if (radioButtonMan.Checked == true)
                    {
                        ViewLeaders(2);
                    }
                    if (radioButtonComands.Checked == true)
                    {
                        ViewLeaders(3);
                    }
                }
                //SetPlace(1);
            }
            if (p == "К" || p == "к")
            {
                if (dataGridView3.ColumnCount > 3)
                {
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        DataGridViewTextBoxCell txtxCell = (DataGridViewTextBoxCell)dataGridView3.Rows[i].Cells[3];
                        if (txtxCell.Value != null)
                        {
                            if (name == txtxCell.Value.ToString())
                            {
                                dataGridView3.Rows[i].Cells[0].Value = false;
                                dataGridView3.Rows[i].Cells[7].Value = time;
                                
                            }

                                

                            
                        }
                    }
                    key = 0;
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        if (dataGridView3.Rows[i].Cells[7].Value != null)
                            if (dataGridView3.Rows[i].Cells[7].Value.ToString() != "")
                            {
                                dataGridView3.Rows[i].Cells[8].Value = Convert.ToString(key + 1);
                                key++;
                            }
                    }
                    if (radioButtonWoman.Checked == true)
                    {
                        ViewLeaders(1);
                    }
                    if (radioButtonMan.Checked == true)
                    {
                        ViewLeaders(2);
                    }
                    if (radioButtonComands.Checked == true)
                    {
                        ViewLeaders(3);
                    }
                }
            }

        }

        private void UpdateJoystickState(string Time)
        {
            JoystickState j = device.CurrentJoystickState;

            //string info = "";
            
            

            byte[] buttons = j.GetButtons();
            for (int i = 0; i < buttons.Length; i++)
            {
                if (buttons[i] != 0)
                {
                    

                    if (i == 5)
                    {
                        
                        if (fourth == false)
                        {
                            button11.Enabled = false;
                            SetTime(PolBox4.Text, Name4.Text, TextTimer4.Text);
                            stopsound.Play();
                            fourth = true;
                            
                            
                        }
                    }
                    if (i == 7)
                    {
                        
                        if (third == false)
                        {
                            button9.Enabled = false;
                            SetTime(PolBox3.Text,Name3.Text,TextTimer3.Text);
                            stopsound.Play();
                            third = true;
                            
                            
                        }
                    }
                    if (i == 6)
                    {
                        
                        if (first == false)
                        {
                            StopButton1.Enabled = false;
                            SetTime(PolBox1.Text,Name1.Text,TextTimer1.Text);
                            stopsound.Play();
                            first = true;
                            
                        }
                    }
                    if (i == 4)
                    {
                        
                        if (second == false)
                        {
                            button7.Enabled = false;
                            SetTime(PolBox2.Text,Name2.Text,TextTimer2.Text);
                            stopsound.Play();
                            second = true;
                            
                        }
                    }
                }
            }
        }


        private void StopButton_Click(object sender, EventArgs e)
        {
            timer.Enabled = false;
            StartButton.Text = "Start";
            TextTimer.Text = "00:00:00.000000";
            NewTimeSpan = DateTime.Now - DateTime.Now;
            first = false;
            second = false;
            third = false;
            fourth = false;
            ToTimerButton.Enabled = true;
            ExportButton.Enabled = true;
            button5.Enabled = true;
            StopButton1.Enabled = false;
            button7.Enabled = false;
            button9.Enabled = false;
            button11.Enabled = false;
        }

        private void PauseButton_Click(object sender, EventArgs e)
        {
            if (timer.Enabled)
            {
                timer.Enabled = false;
                //Pause = true;
            }
            else
            {
                //PreviousTime = DateTime.Now;
                timer.Enabled = true;
                //Pause = false;
            }
            
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.CheckFileExists = true;

            dlg.Filter = "Wav files (*.wav)|*.wav";
            dlg.DefaultExt = ".wav";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                startsoundfilepath.Text = dlg.FileName;

                startsound.SoundLocation = startsoundfilepath.Text;
                Properties.Settings.Default.StartSound = startsoundfilepath.Text;
                Properties.Settings.Default.Save();
                startsound.Load();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.CheckFileExists = true;

            dlg.Filter = "Wav files (*.wav)|*.wav";
            dlg.DefaultExt = ".wav";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                stopsoundfilepath.Text = dlg.FileName;

                stopsound.SoundLocation = stopsoundfilepath.Text;
                Properties.Settings.Default.StopSound = stopsoundfilepath.Text;
                Properties.Settings.Default.Save();
                stopsound.Load();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            stopsound.Play();
        }

        

        private void button5_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.PathtoExcel = "";
            OpenFileDialog OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.Filter = "XLS files (*.xls, *.xlt)|*.xls;*.xlt|XLSX files (*.xlsx, *.xlsm, *.xltx, *.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|ODS files (*.ods, *.ots)|*.ods;*.ots|CSV files (*.csv, *.tsv)|*.csv;*.tsv|HTML files (*.html, *.htm)|*.html;*.htm";
            OpenFileDialog.FilterIndex = 1;

            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.PathtoExcel = OpenFileDialog.FileName;
                Properties.Settings.Default.Save();
            }
            OpenExcel(Properties.Settings.Default.PathtoExcel);
            
        }

        private void OpenExcel (string FilePath)
        {
            int kol = 0;
            if (FilePath != "")
            {
                try
                {

                    //string range = "SELECT * FROM " + "[1$" + txtFrom.Text + ":" + txtTill.Text + "]";
                    string PathConn = "Provider=Microsoft.ACE.OleDb.12.0;Data Source=" + Properties.Settings.Default.PathtoExcel + ";Extended Properties='Excel 12.0 XML;HDR=Yes;IMEX=1';";
                    OleDbConnection conn = new OleDbConnection(PathConn);
                    conn.Open();
                    DataSet ds = new DataSet();
                    System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                    string select = String.Format("SELECT * FROM [{0}]", sheet1);
                    OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(select, conn);


                    myDataAdapter.Fill(ds);
                    titleBox.Text = ds.Tables[0].Rows[0][0].ToString();
                    NameSor = titleBox.Text;
                    for (int i = 5; i < ds.Tables[0].Rows.Count; i++)
                    {
                        if (String.IsNullOrEmpty(ds.Tables[0].Rows[i][0].ToString()))
                        {
                            kol = i;
                            break;
                        }
                        if (ds.Tables[0].Rows[i][5].ToString().Contains("м") || ds.Tables[0].Rows[i][5].ToString().Contains("М"))
                            m++;
                        if (ds.Tables[0].Rows[i][5].ToString().Contains("ж") || ds.Tables[0].Rows[i][5].ToString().Contains("К"))
                            w++;
                        if (ds.Tables[0].Rows[i][5].ToString().Contains("к") || ds.Tables[0].Rows[i][5].ToString().Contains("К"))
                            c++;
                    }


                    if (m>0)
                    {
                        System.Data.DataTable dt = new System.Data.DataTable();

                        dt.Columns.Add("f", typeof(bool));
                        dt.Columns.Add("№ п/п", typeof(int));
                        dt.Columns.Add("Номер участника", typeof(string));
                        dt.Columns.Add("Участник", typeof(string));
                        dt.Columns.Add("Год", typeof(string));
                        dt.Columns.Add("Пол", typeof(string));
                        dt.Columns.Add("Разряд", typeof(string));
                        dt.Columns.Add("Зачёт", typeof(string));
                        dt.Columns.Add("Делегация", typeof(string));
                        dt.Columns.Add("Территория", typeof(string));
                        dt.Columns.Add("Время", typeof(string));
                        dt.Columns.Add("Место", typeof(string));

                        int k=1;
                        for (int i = 5; i < kol; i++)
                        {
                            DataRow pasteRow = dt.NewRow();
                            DataRow tempRow = ds.Tables[0].Rows[i];
                            if (tempRow[5].ToString().Contains("м") || tempRow[5].ToString().Contains("М")) {
                                pasteRow[1] = k; //Npp
                                pasteRow[2] = tempRow[2]; //
                                pasteRow[3] = tempRow[1]; //
                                pasteRow[4] = tempRow[4]; //
                                pasteRow[5] = tempRow[5]; //ПОЛ
                                pasteRow[6] = tempRow[3]; //
                                pasteRow[7] = tempRow[6]; //
                                pasteRow[8] = tempRow[7]; //
                                pasteRow[9] = tempRow[8]; //


                                dt.Rows.Add(pasteRow);
                                k++;
                            }
                            
                            
                        }



                        this.dataGridView2.DataSource = dt;
                        dataGridView2.Columns[0].Width = 30;
                        dataGridView2.Columns[0].HeaderText = "";
                        dataGridView2.Sort(dataGridView2.Columns[10], System.ComponentModel.ListSortDirection.Ascending);
                        conn.Close();
                        conn.Dispose();

                        ToTimerButton.Enabled = true;
                    }

                    if (w>0)
                    {
                        System.Data.DataTable dt = new System.Data.DataTable();

                        dt.Columns.Add("f", typeof(bool));
                        dt.Columns.Add("№ п/п", typeof(int));
                        dt.Columns.Add("Номер участника", typeof(string));
                        dt.Columns.Add("Участник", typeof(string));
                        dt.Columns.Add("Год", typeof(string));
                        dt.Columns.Add("Пол", typeof(string));
                        dt.Columns.Add("Разряд", typeof(string));
                        dt.Columns.Add("Зачёт", typeof(string));
                        dt.Columns.Add("Делегация", typeof(string));
                        dt.Columns.Add("Территория", typeof(string));
                        dt.Columns.Add("Время", typeof(string));
                        dt.Columns.Add("Место", typeof(string));

                        int k = 1;
                        for (int i = 5; i < kol; i++)
                        {
                            DataRow pasteRow = dt.NewRow();
                            DataRow tempRow = ds.Tables[0].Rows[i];
                            if (tempRow[5].ToString().Contains("Ж") || tempRow[5].ToString().Contains("ж"))
                            {

                                pasteRow[1] = k;
                                pasteRow[2] = tempRow[2];
                                pasteRow[3] = tempRow[1];
                                pasteRow[4] = tempRow[4];
                                pasteRow[5] = tempRow[5];
                                pasteRow[6] = tempRow[3];
                                pasteRow[7] = tempRow[6];
                                pasteRow[8] = tempRow[7];
                                pasteRow[9] = tempRow[8];


                                dt.Rows.Add(pasteRow);
                                k++;
                            }
                        }



                        this.dataGridView1.DataSource = dt;
                        dataGridView1.Columns[0].Width = 30;
                        dataGridView1.Columns[0].HeaderText = "";
                        dataGridView1.Sort(dataGridView1.Columns[10], System.ComponentModel.ListSortDirection.Ascending);
                        conn.Close();
                        conn.Dispose();

                        ToTimerButton.Enabled = true;
                    }

                    if (c > 0)
                    {
                        System.Data.DataTable dt = new System.Data.DataTable();

                        dt.Columns.Add("f", typeof(bool));
                        dt.Columns.Add("№ п/п", typeof(int));
                        dt.Columns.Add("Номер участника", typeof(string));
                        dt.Columns.Add("Участник", typeof(string));
                        //dt.Columns.Add("Год", typeof(string));
                        
                        //dt.Columns.Add("Разряд", typeof(string));
                        dt.Columns.Add("Зачёт", typeof(string));
                        dt.Columns.Add("Делегация", typeof(string));
                        dt.Columns.Add("Территория", typeof(string));
                        dt.Columns.Add("Время", typeof(string));
                        dt.Columns.Add("Место", typeof(string));

                        int k = 1;
                        for (int i = 5; i < kol; i++)
                        {
                            DataRow pasteRow = dt.NewRow();
                            DataRow tempRow = ds.Tables[0].Rows[i];
                            if (tempRow[5].ToString().Contains("К") || tempRow[5].ToString().Contains("к"))
                            {
                                pasteRow[1] = k; //Npp
                                pasteRow[2] = tempRow[2]; //
                                pasteRow[3] = tempRow[1]; //
                                pasteRow[4] = tempRow[6]; //
                                //pasteRow[5] = tempRow[5]; //ПОЛ
                                pasteRow[5] = tempRow[7]; //
                                pasteRow[6] = tempRow[8]; //
                                //pasteRow[7] = tempRow[8]; //
                                //pasteRow[8] = tempRow[8]; //


                                dt.Rows.Add(pasteRow);
                                k++;
                            }


                        }



                        this.dataGridView3.DataSource = dt;
                        dataGridView3.Columns[0].Width = 30;
                        dataGridView3.Columns[0].HeaderText = "";
                        dataGridView3.Sort(dataGridView3.Columns[7], System.ComponentModel.ListSortDirection.Ascending);
                        conn.Close();
                        conn.Dispose();

                        ToTimerButton.Enabled = true;
                    }

                }
                catch
                {
                    MessageBox.Show("ERROR!");
                }
            }
        }

        private void ToTimerButton_Click(object sender, EventArgs e)
        {
            Name1.Text = "";
            Name2.Text = "";
            Name3.Text = "";
            Name4.Text = "";
            Npptext1.Text = "";
            Npptext2.Text = "";
            Npptext3.Text = "";
            Npptext4.Text = "";
            Ntext1.Text = "";
            Ntext2.Text = "";
            Ntext3.Text = "";
            Ntext4.Text = "";
            List<Part> partsName = new List<Part>();
            //List<Part> partsNpp = new List<Part>();
            //List<Part> partsN = new List<Part>();
            if (tabControl2.SelectedTab.Text == "Женщины")
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].ValueType == typeof(bool))
                    {
                        try
                        {

                            if (Convert.ToBoolean(dataGridView1.Rows[i].Cells[0].Value) == true & dataGridView1.Rows[i].Cells[10].Value.ToString() == "")
                            {
                                partsName.Add(new Part()
                                {
                                    PartName = dataGridView1.Rows[i].Cells[3].Value.ToString(),
                                    PartNpp = dataGridView1.Rows[i].Cells[1].Value.ToString(),
                                    PartN = dataGridView1.Rows[i].Cells[2].Value.ToString(),
                                    Pol = dataGridView1.Rows[i].Cells[5].Value.ToString()
                                });
                                //partsNpp.Add(new Part() { PartNpp = dataGridView1.Rows[i].Cells[1].Value.ToString() });
                                //partsN.Add(new Part() { PartN = dataGridView1.Rows[i].Cells[2].Value.ToString() });

                            }
                        }
                        catch
                        {

                        }
                    }
                }
            if (tabControl2.SelectedTab.Text == "Мужчины")
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[0].ValueType == typeof(bool))
                    {
                        try
                        {

                            if (Convert.ToBoolean(dataGridView2.Rows[i].Cells[0].Value) == true & dataGridView2.Rows[i].Cells[10].Value.ToString() == "")
                            {
                                partsName.Add(new Part()
                                {
                                    PartName = dataGridView2.Rows[i].Cells[3].Value.ToString(),
                                    PartNpp = dataGridView2.Rows[i].Cells[1].Value.ToString(),
                                    PartN = dataGridView2.Rows[i].Cells[2].Value.ToString(),
                                    Pol = dataGridView2.Rows[i].Cells[5].Value.ToString()
                                });
                                //partsNpp.Add(new Part() { PartNpp = dataGridView1.Rows[i].Cells[1].Value.ToString() });
                                //partsN.Add(new Part() { PartN = dataGridView1.Rows[i].Cells[2].Value.ToString() });

                            }
                        }
                        catch
                        {

                        }


                    }

                }
                
            }
            if (tabControl2.SelectedTab.Text == "Команды")
            {
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (dataGridView3.Rows[i].Cells[0].ValueType == typeof(bool))
                    {
                        try
                        {

                            if (Convert.ToBoolean(dataGridView3.Rows[i].Cells[0].Value) == true & dataGridView3.Rows[i].Cells[7].Value.ToString() == "")
                            {
                                partsName.Add(new Part()
                                {
                                    PartName = dataGridView3.Rows[i].Cells[3].Value.ToString(),
                                    PartNpp = dataGridView3.Rows[i].Cells[1].Value.ToString(),
                                    PartN = dataGridView3.Rows[i].Cells[2].Value.ToString(),
                                    Pol = "к"
                                });
                                //partsNpp.Add(new Part() { PartNpp = dataGridView1.Rows[i].Cells[1].Value.ToString() });
                                //partsN.Add(new Part() { PartN = dataGridView1.Rows[i].Cells[2].Value.ToString() });

                            }
                        }
                        catch
                        {

                        }
                    }
                }
            }
            //if (partsName.Count > 0) {
            //    if (partsName.Count == 1) {
            //        Npptext1.Text = partsName[0].PartNpp;
            //        Ntext1.Text = partsName[0].PartN;
            //        Name1.Text = partsName[0].PartName;
            //        PolBox1.Text = partsName[0].Pol;
            //    }
            //    if (partsName.Count == 2)
            //    {
            //        Npptext1.Text = partsName[0].PartNpp;
            //        Ntext1.Text = partsName[0].PartN;
            //        Name1.Text = partsName[0].PartName;
            //        PolBox1.Text = partsName[0].Pol;
            //        Npptext2.Text = partsName[1].PartNpp;
            //        Ntext2.Text = partsName[1].PartN;
            //        Name2.Text = partsName[1].PartName;
            //        PolBox2.Text = partsName[1].Pol;
            //    }
            //    if (partsName.Count == 3)
            //    {
            //        Npptext1.Text = partsName[0].PartNpp;
            //        Ntext1.Text = partsName[0].PartN;
            //        Name1.Text = partsName[0].PartName;
            //        PolBox1.Text = partsName[0].Pol;
            //        Npptext2.Text = partsName[1].PartNpp;
            //        Ntext2.Text = partsName[1].PartN;
            //        Name2.Text = partsName[1].PartName;
            //        PolBox2.Text = partsName[1].Pol;
            //        Npptext3.Text = partsName[2].PartNpp;
            //        Ntext3.Text = partsName[2].PartN;
            //        Name3.Text = partsName[2].PartName;
            //        PolBox3.Text = partsName[2].Pol;
            //    }
            //    if (partsName.Count == 4)
            //    {
            try
            {
                Npptext1.Text = partsName[0].PartNpp;
                Ntext1.Text = partsName[0].PartN;
                Name1.Text = partsName[0].PartName;
                PolBox1.Text = partsName[0].Pol;
            }
            catch { }
            try
            {
                Npptext2.Text = partsName[1].PartNpp;
                Ntext2.Text = partsName[1].PartN;
                Name2.Text = partsName[1].PartName;
                PolBox2.Text = partsName[1].Pol;
            }
            catch { }
            try
            {
                Npptext3.Text = partsName[2].PartNpp;
                Ntext3.Text = partsName[2].PartN;
                Name3.Text = partsName[2].PartName;
                PolBox3.Text = partsName[2].Pol;
            }
            catch { }
            try
            {
                Npptext4.Text = partsName[3].PartNpp;
                Ntext4.Text = partsName[3].PartN;
                Name4.Text = partsName[3].PartName;
                PolBox4.Text = partsName[3].Pol;
            }
            catch { }
                    
                    
            //    }

            //}

        }


        private void ExportButton_Click(object sender, EventArgs e)
        {
            
            ExportToExcel();
            progressBar1.Value = 1;
        }

        private void ExportToExcel()
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            progressBar1.Value = 2;
            if (w > 0)
            {
                ExcelApp.Workbooks.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;

                workSheet.Name = "Женщины";

                workSheet.PageSetup.Zoom = 70;

                workSheet.Columns[1].ColumnWidth = 3.57;
                workSheet.Columns[2].ColumnWidth = 5.71;

                workSheet.Columns[3].ColumnWidth = 24.29;
                workSheet.Columns[4].ColumnWidth = 4.86;
                workSheet.Columns[5].ColumnWidth = 6.29;
                workSheet.Columns[6].ColumnWidth = 20.14;
                workSheet.Columns[7].ColumnWidth = 20.14;
                workSheet.Columns[8].ColumnWidth = 11.14;
                workSheet.Columns[9].ColumnWidth = 4.14;
                workSheet.Columns[10].ColumnWidth = 10;
                workSheet.Columns[11].ColumnWidth = 6.57;
                progressBar1.Value = 5;
                workSheet.Cells[1, 1] = "Министерство физической культуры и спорта Хабаровского края";

                Excel.Range oRange;
                oRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[4, 1], workSheet.Cells[4, 11]];
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[5, 1], workSheet.Cells[5, 3]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[6, 1], workSheet.Cells[6, 11]];
                oRange.Merge(Type.Missing);
                progressBar1.Value = 10;
                                (workSheet.Cells[1, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[1, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;



                workSheet.Cells[2, 1] = "КГАУ \"Краевой центр развития спорта\"";
                (workSheet.Cells[2, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[2, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[2, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[3, 1] = "РОО \"Федерация спортивного туризма Хабаровского края\"";
                (workSheet.Cells[3, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[3, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[3, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[4, 1] = NameSor;
                ((Excel.Range)workSheet.Rows[4]).RowHeight = 65.25;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Size = 28;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[4, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                (workSheet.Cells[4, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[4, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                string format = "d MMMM yyyy";
                workSheet.Cells[5, 1] = DateTime.Now.Date.ToString(format) + " года";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Size = 10;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Italic = true;

                workSheet.Cells[6, 1] = "Протокол соревнований в дисциплине: \"дистанция - пешеходная\" 3 класса, код ВРВС 0840091811Я ЖЕНЩИНЫ";
                ((Excel.Range)workSheet.Rows[6]).RowHeight = 72;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Size = 16;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[6, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[6, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[6, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                ((Excel.Range)workSheet.Rows[8]).RowHeight = 42.75;
                ((Excel.Range)workSheet.Rows[9]).RowHeight = 135;
                progressBar1.Value = 15;
                workSheet.Cells[8, 1] = "№ п/п";
                (workSheet.Cells[8, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 1] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 2] = "Номер участника";
                (workSheet.Cells[8, 2] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 2] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 3] = "Участник";
                (workSheet.Cells[8, 3] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 3] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                workSheet.Cells[8, 4] = "Год";
                (workSheet.Cells[8, 4] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 4] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 4] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 4] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 5] = "Разряд";
                (workSheet.Cells[8, 5] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 5] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 5] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 6] = "Делегация";
                (workSheet.Cells[8, 6] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 6] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 7] = "Территория";
                (workSheet.Cells[8, 7] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 7] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 7] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 8] = "Результат";
                (workSheet.Cells[8, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[8, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 8] = "Результат";
                (workSheet.Cells[9, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 8] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 9] = "Место";
                (workSheet.Cells[9, 9] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 9] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 9] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 9] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 10] = "% от результата победителя";
                (workSheet.Cells[9, 10] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 10] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 10] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 10] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 10] as Excel.Range).WrapText = true;

                workSheet.Cells[9, 11] = "Выполненный норматив";
                (workSheet.Cells[9, 11] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 11] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 11] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 11] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 11] as Excel.Range).WrapText = true;

                progressBar1.Value = 20;


                key = 0;
                for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[10].Value != null)
                        if (dataGridView1.Rows[i].Cells[10].Value.ToString() != "")
                        {
                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView1.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView1.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView1.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView1.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView1.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView1.Rows[i].Cells[9].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 8] = dataGridView1.Rows[i].Cells[10].Value;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 8], workSheet.Cells[key + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView1.Rows[i].Cells[11].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }

                }
                for(int i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[10] != null)
                        if (dataGridView1.Rows[i].Cells[10].Value.ToString() == "")
                        {
                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView1.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView1.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView1.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView1.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView1.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView1.Rows[i].Cells[9].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 8] = dataGridView1.Rows[i].Cells[10].Value;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 8], workSheet.Cells[key + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView1.Rows[i].Cells[11].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }

                }

                progressBar1.Value = 30;
                oRange = workSheet.Range[workSheet.Cells[8, 1], workSheet.Cells[9, 1]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 2], workSheet.Cells[9, 2]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 3], workSheet.Cells[9, 3]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 4], workSheet.Cells[9, 4]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 5], workSheet.Cells[9, 5]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 6], workSheet.Cells[9, 6]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 7], workSheet.Cells[9, 7]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 8], workSheet.Cells[8, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[9, 8], workSheet.Cells[9, 8]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 9], workSheet.Cells[9, 9]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 10], workSheet.Cells[9, 10]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 11], workSheet.Cells[9, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[6, 11]];
                //oRange.Group
                progressBar1.Value = 35;

            }

            if (m > 0)
            {
                progressBar1.Value = 36;
                ExcelApp.Worksheets.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;

                workSheet.Name = "Мужчины";

                workSheet.PageSetup.Zoom = 70;

                workSheet.Columns[1].ColumnWidth = 3.57;
                workSheet.Columns[2].ColumnWidth = 5.71;

                workSheet.Columns[3].ColumnWidth = 24.29;
                workSheet.Columns[4].ColumnWidth = 4.86;
                workSheet.Columns[5].ColumnWidth = 6.29;
                workSheet.Columns[6].ColumnWidth = 20.14;
                workSheet.Columns[7].ColumnWidth = 20.14;
                workSheet.Columns[8].ColumnWidth = 11.14;
                workSheet.Columns[9].ColumnWidth = 4.14;
                workSheet.Columns[10].ColumnWidth = 10;
                workSheet.Columns[11].ColumnWidth = 6.57;

                workSheet.Cells[1, 1] = "Министерство физической культуры и спорта Хабаровского края";

                Excel.Range oRange;
                oRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[4, 1], workSheet.Cells[4, 11]];
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[5, 1], workSheet.Cells[5, 3]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[6, 1], workSheet.Cells[6, 11]];
                oRange.Merge(Type.Missing);

                
                (workSheet.Cells[1, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[1, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                progressBar1.Value = 37;

                workSheet.Cells[2, 1] = "КГАУ \"Краевой центр развития спорта\"";
                (workSheet.Cells[2, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[2, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[2, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[3, 1] = "РОО \"Федерация спортивного туризма Хабаровского края\"";
                (workSheet.Cells[3, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[3, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[3, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[4, 1] = NameSor;
                ((Excel.Range)workSheet.Rows[4]).RowHeight = 65.25;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Size = 28;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[4, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                (workSheet.Cells[4, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[4, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                string format = "d MMMM yyyy";
                workSheet.Cells[5, 1] = DateTime.Now.Date.ToString(format) + " года";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Size = 10;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Italic = true;

                workSheet.Cells[6, 1] = "Протокол соревнований в дисциплине: \"дистанция - пешеходная\" 3 класса, код ВРВС 0840091811Я Мужчины";
                ((Excel.Range)workSheet.Rows[6]).RowHeight = 72;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Size = 16;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[6, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[6, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[6, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                ((Excel.Range)workSheet.Rows[8]).RowHeight = 42.75;
                ((Excel.Range)workSheet.Rows[9]).RowHeight = 135;
                progressBar1.Value = 42;
                workSheet.Cells[8, 1] = "№ п/п";
                (workSheet.Cells[8, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 1] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 2] = "Номер участника";
                (workSheet.Cells[8, 2] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 2] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 3] = "Участник";
                (workSheet.Cells[8, 3] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 3] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                workSheet.Cells[8, 4] = "Год";
                (workSheet.Cells[8, 4] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 4] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 4] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 4] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 5] = "Разряд";
                (workSheet.Cells[8, 5] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 5] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 5] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 6] = "Делегация";
                (workSheet.Cells[8, 6] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 6] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 7] = "Территория";
                (workSheet.Cells[8, 7] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 7] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 7] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 8] = "Результат";
                (workSheet.Cells[8, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[8, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 8] = "Результат";
                (workSheet.Cells[9, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 8] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 9] = "Место";
                (workSheet.Cells[9, 9] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 9] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 9] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 9] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 10] = "% от результата победителя";
                (workSheet.Cells[9, 10] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 10] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 10] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 10] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 10] as Excel.Range).WrapText = true;

                workSheet.Cells[9, 11] = "Выполненный норматив";
                (workSheet.Cells[9, 11] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 11] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 11] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 11] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 11] as Excel.Range).WrapText = true;

                progressBar1.Value = 47;
                key = 0;
                for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
                {
                    if (dataGridView2.Rows[i].Cells[10] != null)
                        if (dataGridView2.Rows[i].Cells[10].Value.ToString() != "")
                        {

                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView2.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView2.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView2.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView2.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView2.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView2.Rows[i].Cells[9].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 8] = dataGridView2.Rows[i].Cells[10].Value;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 8], workSheet.Cells[key + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView2.Rows[i].Cells[11].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }


                }
                for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
                {
                    if (dataGridView2.Rows[i].Cells[10] != null)
                        if (dataGridView2.Rows[i].Cells[10].Value.ToString() == "")
                        {
                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView2.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView2.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView2.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView2.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView2.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView2.Rows[i].Cells[9].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 8] = dataGridView2.Rows[i].Cells[10].Value;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 8], workSheet.Cells[key + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView2.Rows[i].Cells[11].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }


                }

                progressBar1.Value = 55;

                oRange = workSheet.Range[workSheet.Cells[8, 1], workSheet.Cells[9, 1]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 2], workSheet.Cells[9, 2]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 3], workSheet.Cells[9, 3]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 4], workSheet.Cells[9, 4]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 5], workSheet.Cells[9, 5]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 6], workSheet.Cells[9, 6]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 7], workSheet.Cells[9, 7]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 8], workSheet.Cells[8, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[9, 8], workSheet.Cells[9, 8]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 9], workSheet.Cells[9, 9]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 10], workSheet.Cells[9, 10]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 11], workSheet.Cells[9, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                progressBar1.Value = 60;
            }


            if (c>0)
            {
                progressBar1.Value = 61;
                ExcelApp.Worksheets.Add();
                Excel._Worksheet workSheet = (Excel.Worksheet)ExcelApp.ActiveSheet;

                workSheet.Name = "Команды";

                workSheet.PageSetup.Zoom = 70;

                workSheet.Columns[1].ColumnWidth = 3.57;
                workSheet.Columns[2].ColumnWidth = 5.71;

                workSheet.Columns[3].ColumnWidth = 24.29;
                workSheet.Columns[4].ColumnWidth = 4.86;
                workSheet.Columns[5].ColumnWidth = 6.29;
                workSheet.Columns[6].ColumnWidth = 20.14;
                workSheet.Columns[7].ColumnWidth = 20.14;
                workSheet.Columns[8].ColumnWidth = 11.14;
                workSheet.Columns[9].ColumnWidth = 4.14;
                workSheet.Columns[10].ColumnWidth = 10;
                workSheet.Columns[11].ColumnWidth = 6.57;

                workSheet.Cells[1, 1] = "Министерство физической культуры и спорта Хабаровского края";

                Excel.Range oRange;
                oRange = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[3, 1], workSheet.Cells[3, 11]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[4, 1], workSheet.Cells[4, 11]];
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                //oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[5, 1], workSheet.Cells[5, 3]];
                oRange.Merge(Type.Missing);
                oRange = workSheet.Range[workSheet.Cells[6, 1], workSheet.Cells[6, 11]];
                oRange.Merge(Type.Missing);


                (workSheet.Cells[1, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[1, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[1, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[1, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                progressBar1.Value = 75;

                workSheet.Cells[2, 1] = "КГАУ \"Краевой центр развития спорта\"";
                (workSheet.Cells[2, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[2, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[2, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[2, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[3, 1] = "РОО \"Федерация спортивного туризма Хабаровского края\"";
                (workSheet.Cells[3, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Size = 12;
                (workSheet.Cells[3, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[3, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[3, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                workSheet.Cells[4, 1] = NameSor;
                ((Excel.Range)workSheet.Rows[4]).RowHeight = 65.25;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Size = 28;
                (workSheet.Cells[4, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[4, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                (workSheet.Cells[4, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[4, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                string format = "d MMMM yyyy";
                workSheet.Cells[5, 1] = DateTime.Now.Date.ToString(format) + " года";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Bold = false;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Size = 10;
                (workSheet.Cells[5, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[5, 1] as Excel.Range).Font.Italic = true;

                workSheet.Cells[6, 1] = "Протокол соревнований в дисциплине: \"дистанция - пешеходная\" 3 класса, код ВРВС 0840091811Я Команды";
                ((Excel.Range)workSheet.Rows[6]).RowHeight = 72;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Size = 16;
                (workSheet.Cells[6, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[6, 1] as Excel.Range).WrapText = true;
                (workSheet.Cells[6, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[6, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                ((Excel.Range)workSheet.Rows[8]).RowHeight = 42.75;
                ((Excel.Range)workSheet.Rows[9]).RowHeight = 135;
                progressBar1.Value = 80;
                workSheet.Cells[8, 1] = "№ п/п";
                (workSheet.Cells[8, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 1] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 1] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 2] = "Номер участника";
                (workSheet.Cells[8, 2] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 2] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 2] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 3] = "Участник";
                (workSheet.Cells[8, 3] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 3] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 3] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                workSheet.Cells[8, 4] = "Год";
                (workSheet.Cells[8, 4] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 4] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 4] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 4] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 4] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 5] = "Разряд";
                (workSheet.Cells[8, 5] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 5] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 5] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 5] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[8, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[8, 6] = "Делегация";
                (workSheet.Cells[8, 6] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 6] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 6] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 7] = "Территория";
                (workSheet.Cells[8, 7] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 7] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 7] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[8, 7] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                workSheet.Cells[8, 8] = "Результат";
                (workSheet.Cells[8, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[8, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[8, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[8, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 8] = "Результат";
                (workSheet.Cells[9, 8] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 8] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 8] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 8] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 8] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 9] = "Место";
                (workSheet.Cells[9, 9] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 9] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 9] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 9] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 9] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[9, 10] = "% от результата победителя";
                (workSheet.Cells[9, 10] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 10] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 10] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 10] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 10] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 10] as Excel.Range).WrapText = true;

                workSheet.Cells[9, 11] = "Выполненный норматив";
                (workSheet.Cells[9, 11] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Size = 11;
                (workSheet.Cells[9, 11] as Excel.Range).Font.Name = "Arial";
                (workSheet.Cells[9, 11] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                (workSheet.Cells[9, 11] as Excel.Range).Orientation = Excel.XlOrientation.xlUpward;
                (workSheet.Cells[9, 11] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 11] as Excel.Range).WrapText = true;

                key = 0;
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    if (dataGridView3.Rows[i].Cells[7] != null)
                        if (dataGridView3.Rows[i].Cells[7].Value.ToString() != "")
                        {
                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView3.Rows[i].Cells[1].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView3.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView3.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView3.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView3.Rows[i].Cells[5].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView3.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[i + 10, 8] = dataGridView3.Rows[i].Cells[7].Value;
                            (workSheet.Cells[i + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[i + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[i + 10, 8], workSheet.Cells[i + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView3.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }
                }
                for (int i = 0; i < dataGridView3.Rows.Count - 1; i++)
                {
                    if (dataGridView3.Rows[i].Cells[7] != null)
                        if (dataGridView3.Rows[i].Cells[7].Value.ToString() != "")
                        {
                            workSheet.Cells[key + 10, 1] = key + 1;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 1] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 1], workSheet.Cells[key + 10, 1]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[key + 10, 2] = dataGridView3.Rows[i].Cells[1].Value;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 2] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 2], workSheet.Cells[key + 10, 2]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 3] = dataGridView3.Rows[i].Cells[2].Value;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 3] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 3], workSheet.Cells[key + 10, 3]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 4] = dataGridView3.Rows[i].Cells[3].Value;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 4] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 4], workSheet.Cells[key + 10, 4]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 5] = dataGridView3.Rows[i].Cells[4].Value;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 5] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 5], workSheet.Cells[key + 10, 5]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 6] = dataGridView3.Rows[i].Cells[5].Value;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 6] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 6], workSheet.Cells[key + 10, 6]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 7] = dataGridView3.Rows[i].Cells[6].Value;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 7] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 7], workSheet.Cells[key + 10, 7]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

                            workSheet.Cells[i + 10, 8] = dataGridView3.Rows[i].Cells[7].Value;
                            (workSheet.Cells[i + 10, 8] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[i + 10, 8] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[i + 10, 8], workSheet.Cells[i + 10, 8]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                            workSheet.Cells[key + 10, 9] = dataGridView3.Rows[i].Cells[8].Value;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Size = 10;
                            (workSheet.Cells[key + 10, 9] as Excel.Range).Font.Name = "Arial";
                            oRange = workSheet.Range[workSheet.Cells[key + 10, 9], workSheet.Cells[key + 10, 9]];
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                            oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                            key++;
                        }
                }
                progressBar1.Value = 90;
                oRange = workSheet.Range[workSheet.Cells[8, 1], workSheet.Cells[9, 1]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 2], workSheet.Cells[9, 2]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 3], workSheet.Cells[9, 3]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 4], workSheet.Cells[9, 4]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 5], workSheet.Cells[9, 5]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 6], workSheet.Cells[9, 6]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 7], workSheet.Cells[9, 7]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[8, 8], workSheet.Cells[8, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Merge(Type.Missing);

                oRange = workSheet.Range[workSheet.Cells[9, 8], workSheet.Cells[9, 8]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 9], workSheet.Cells[9, 9]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 10], workSheet.Cells[9, 10]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;

                oRange = workSheet.Range[workSheet.Cells[9, 11], workSheet.Cells[9, 11]];
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                oRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            }


            ExcelApp.Visible = true;

            progressBar1.Value = 100;
            //ExcelApp.Workbooks.Close();

        }


        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void StartButton1_Click(object sender, EventArgs e)
        {
            second = true;
            third = true;
            fourth = true;
            PreviousTime = DateTime.Now.Subtract(NewTimeSpan);
            timer.Elapsed += timer_tick;
            timer.Enabled = true;
        }

        private void StopButton1_Click(object sender, EventArgs e)
        {
            if (first == false)
            {
                SetTime(PolBox1.Text, Name1.Text, TextTimer1.Text);
                stopsound.Play();
                first = true;
                StopButton1.Enabled = false;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (second == false)
            {
                SetTime(PolBox2.Text, Name2.Text, TextTimer2.Text);
                stopsound.Play();
                second = true;
                button7.Enabled = false;
            }
        }

        private void radioButtonWoman_CheckedChanged(object sender, EventArgs e)
        {
            ViewLeaders(1);
        }

        private void ViewLeaders(int p)
        {
            string text = "";
            //LeaderBox.Text = "";
            if (LeaderBox.InvokeRequired) LeaderBox.Invoke(new Action<string>((s) => LeaderBox.Text = s), "");
            else LeaderBox.Text = "";
            if (p == 1)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[11].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "1")
                            text += "1: " + dataGridView1.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "2")
                            text += "2: " + dataGridView1.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "3")
                            text += "3: " + dataGridView1.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "4")
                            text += "4: " + dataGridView1.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView1.Rows[i].Cells[11].Value.ToString() == "5")
                            text += "5: " + dataGridView1.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                    }
                }
                Properties.Settings.Default.Pol = 1;
                Properties.Settings.Default.Save();
            }
            if (p == 2)
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[11].Value != null)
                    {
                        if (dataGridView2.Rows[i].Cells[11].Value.ToString() == "1")
                            text += "1: " + dataGridView2.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView2.Rows[i].Cells[11].Value.ToString() == "2")
                            text += "2: " + dataGridView2.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView2.Rows[i].Cells[11].Value.ToString() == "3")
                            text += "3: " + dataGridView2.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView2.Rows[i].Cells[11].Value.ToString() == "4")
                            text += "4: " + dataGridView2.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView2.Rows[i].Cells[11].Value.ToString() == "5")
                            text += "5: " + dataGridView2.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                    }
                }
                Properties.Settings.Default.Pol = 2;
                Properties.Settings.Default.Save();
            }
            if (p == 3)
            {
                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (dataGridView3.Rows[i].Cells[8].Value != null)
                    {
                        if (dataGridView3.Rows[i].Cells[8].Value.ToString() == "1")
                            text += "1: " + dataGridView3.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView3.Rows[i].Cells[8].Value.ToString() == "2")
                            text += "2: " + dataGridView3.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView3.Rows[i].Cells[8].Value.ToString() == "3")
                            text += "3: " + dataGridView3.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView3.Rows[i].Cells[8].Value.ToString() == "4")
                            text += "4: " + dataGridView3.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                        if (dataGridView3.Rows[i].Cells[8].Value.ToString() == "5")
                            text += "5: " + dataGridView3.Rows[i].Cells[3].Value.ToString() + Environment.NewLine;
                    }
                }
                Properties.Settings.Default.Pol = 3;
                Properties.Settings.Default.Save();
            }
            if (LeaderBox.InvokeRequired) LeaderBox.Invoke(new Action<string>((s) => LeaderBox.Text = s), text);
            else LeaderBox.Text = text;

        }

        private void radioButtonMan_CheckedChanged(object sender, EventArgs e)
        {
            ViewLeaders(2);
        }

        private void radioButtonComands_CheckedChanged(object sender, EventArgs e)
        {
            ViewLeaders(3);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (third == false)
            {
                SetTime(PolBox3.Text, Name3.Text, TextTimer3.Text);
                stopsound.Play();
                third = true;
                button9.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (fourth == false)
            {
                SetTime(PolBox4.Text, Name4.Text, TextTimer4.Text);
                stopsound.Play();
                fourth = true;
                button11.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            startsound.Play();
        }
    }

    public class Part : IEquatable<Part>
    { 
        public string PartName { get; set; }

        public string PartNpp { get; set; }

        public string PartN { get; set; }
        
        public string Pol { get; set; }

        public int PartId { get; set; }

        public override string ToString()
        {
            return "ID: " + PartId + "   Name: " + PartName;
        }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            Part objAsPart = obj as Part;
            if (objAsPart == null) return false;
            else return Equals(objAsPart);
        }
        public override int GetHashCode()
        {
            return PartId;
        }

        public bool Equals(Part other)
        {
            if (other == null) return false;
            return (this.PartId.Equals(other.PartId));
        }
    }
}
