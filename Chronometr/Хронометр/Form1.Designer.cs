using Microsoft.DirectX;
using Microsoft.DirectX.DirectInput;

namespace WindowsFormsApp2
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.TextTimer = new System.Windows.Forms.Label();
            this.StartButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.StopButton = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.titleBox = new System.Windows.Forms.TextBox();
            this.LeaderBox = new System.Windows.Forms.TextBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.PolBox4 = new System.Windows.Forms.TextBox();
            this.button11 = new System.Windows.Forms.Button();
            this.Name4 = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.Npptext4 = new System.Windows.Forms.TextBox();
            this.TextTimer4 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.Ntext4 = new System.Windows.Forms.TextBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.PolBox3 = new System.Windows.Forms.TextBox();
            this.button9 = new System.Windows.Forms.Button();
            this.Name3 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.Npptext3 = new System.Windows.Forms.TextBox();
            this.TextTimer3 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.Ntext3 = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.PolBox2 = new System.Windows.Forms.TextBox();
            this.button7 = new System.Windows.Forms.Button();
            this.Name2 = new System.Windows.Forms.TextBox();
            this.TextTimer2 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.Npptext2 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.Ntext2 = new System.Windows.Forms.TextBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.PolBox1 = new System.Windows.Forms.TextBox();
            this.StopButton1 = new System.Windows.Forms.Button();
            this.Name1 = new System.Windows.Forms.TextBox();
            this.TextTimer1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.Npptext1 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.Ntext1 = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.radioButtonComands = new System.Windows.Forms.RadioButton();
            this.radioButtonWoman = new System.Windows.Forms.RadioButton();
            this.radioButtonMan = new System.Windows.Forms.RadioButton();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.Woman = new System.Windows.Forms.TabPage();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Man = new System.Windows.Forms.TabPage();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.Commands = new System.Windows.Forms.TabPage();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.button5 = new System.Windows.Forms.Button();
            this.ToTimerButton = new System.Windows.Forms.Button();
            this.ExportButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.stopsoundfilepath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.startsoundfilepath = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.Woman.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.Man.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            this.Commands.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TextTimer
            // 
            this.TextTimer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TextTimer.AutoSize = true;
            this.TextTimer.Font = new System.Drawing.Font("Microsoft Sans Serif", 60F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TextTimer.Location = new System.Drawing.Point(249, 188);
            this.TextTimer.Name = "TextTimer";
            this.TextTimer.Size = new System.Drawing.Size(648, 91);
            this.TextTimer.TabIndex = 0;
            this.TextTimer.Text = "00:00:00.000000";
            // 
            // StartButton
            // 
            this.StartButton.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.StartButton.Enabled = false;
            this.StartButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StartButton.Location = new System.Drawing.Point(253, 280);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(177, 85);
            this.StartButton.TabIndex = 2;
            this.StartButton.Text = "Старт";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // label2
            // 
            this.label2.AutoEllipsis = true;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(121, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Контроллер не найден";
            // 
            // StopButton
            // 
            this.StopButton.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.StopButton.Enabled = false;
            this.StopButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.StopButton.Location = new System.Drawing.Point(727, 280);
            this.StopButton.Name = "StopButton";
            this.StopButton.Size = new System.Drawing.Size(170, 85);
            this.StopButton.TabIndex = 4;
            this.StopButton.Text = "Стоп";
            this.StopButton.UseVisualStyleBackColor = true;
            this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1165, 747);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.titleBox);
            this.tabPage1.Controls.Add(this.LeaderBox);
            this.tabPage1.Controls.Add(this.groupBox5);
            this.tabPage1.Controls.Add(this.groupBox8);
            this.tabPage1.Controls.Add(this.groupBox4);
            this.tabPage1.Controls.Add(this.pictureBox2);
            this.tabPage1.Controls.Add(this.pictureBox1);
            this.tabPage1.Controls.Add(this.groupBox3);
            this.tabPage1.Controls.Add(this.StartButton);
            this.tabPage1.Controls.Add(this.StopButton);
            this.tabPage1.Controls.Add(this.TextTimer);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1157, 721);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Основной";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // titleBox
            // 
            this.titleBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.titleBox.BackColor = System.Drawing.SystemColors.Window;
            this.titleBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.titleBox.Font = new System.Drawing.Font("Arial", 20F, System.Drawing.FontStyle.Bold);
            this.titleBox.Location = new System.Drawing.Point(141, 6);
            this.titleBox.Multiline = true;
            this.titleBox.Name = "titleBox";
            this.titleBox.ReadOnly = true;
            this.titleBox.Size = new System.Drawing.Size(866, 88);
            this.titleBox.TabIndex = 27;
            this.titleBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // LeaderBox
            // 
            this.LeaderBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LeaderBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.LeaderBox.Location = new System.Drawing.Point(436, 313);
            this.LeaderBox.Multiline = true;
            this.LeaderBox.Name = "LeaderBox";
            this.LeaderBox.Size = new System.Drawing.Size(285, 408);
            this.LeaderBox.TabIndex = 24;
            // 
            // groupBox5
            // 
            this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox5.Controls.Add(this.PolBox4);
            this.groupBox5.Controls.Add(this.button11);
            this.groupBox5.Controls.Add(this.Name4);
            this.groupBox5.Controls.Add(this.label18);
            this.groupBox5.Controls.Add(this.Npptext4);
            this.groupBox5.Controls.Add(this.TextTimer4);
            this.groupBox5.Controls.Add(this.label20);
            this.groupBox5.Controls.Add(this.Ntext4);
            this.groupBox5.Location = new System.Drawing.Point(945, 424);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(212, 297);
            this.groupBox5.TabIndex = 21;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Четвертый участник";
            // 
            // PolBox4
            // 
            this.PolBox4.Location = new System.Drawing.Point(0, 276);
            this.PolBox4.Name = "PolBox4";
            this.PolBox4.Size = new System.Drawing.Size(10, 20);
            this.PolBox4.TabIndex = 29;
            this.PolBox4.Visible = false;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(68, 266);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(75, 23);
            this.button11.TabIndex = 22;
            this.button11.Text = "Стоп";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // Name4
            // 
            this.Name4.BackColor = System.Drawing.SystemColors.Window;
            this.Name4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Name4.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Name4.Location = new System.Drawing.Point(6, 71);
            this.Name4.Multiline = true;
            this.Name4.Name = "Name4";
            this.Name4.ReadOnly = true;
            this.Name4.Size = new System.Drawing.Size(200, 103);
            this.Name4.TabIndex = 14;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(44, 19);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(38, 13);
            this.label18.TabIndex = 3;
            this.label18.Text = "№ п/п";
            // 
            // Npptext4
            // 
            this.Npptext4.BackColor = System.Drawing.SystemColors.Window;
            this.Npptext4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Npptext4.Location = new System.Drawing.Point(88, 19);
            this.Npptext4.Name = "Npptext4";
            this.Npptext4.ReadOnly = true;
            this.Npptext4.Size = new System.Drawing.Size(82, 13);
            this.Npptext4.TabIndex = 20;
            // 
            // TextTimer4
            // 
            this.TextTimer4.AutoSize = true;
            this.TextTimer4.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TextTimer4.Location = new System.Drawing.Point(-2, 232);
            this.TextTimer4.Name = "TextTimer4";
            this.TextTimer4.Size = new System.Drawing.Size(218, 31);
            this.TextTimer4.TabIndex = 1;
            this.TextTimer4.Text = "00:00:00.000000";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(10, 42);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(72, 13);
            this.label20.TabIndex = 4;
            this.label20.Text = "№ участника";
            // 
            // Ntext4
            // 
            this.Ntext4.BackColor = System.Drawing.SystemColors.Window;
            this.Ntext4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Ntext4.Location = new System.Drawing.Point(88, 42);
            this.Ntext4.Name = "Ntext4";
            this.Ntext4.ReadOnly = true;
            this.Ntext4.Size = new System.Drawing.Size(82, 13);
            this.Ntext4.TabIndex = 19;
            // 
            // groupBox8
            // 
            this.groupBox8.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox8.Controls.Add(this.PolBox3);
            this.groupBox8.Controls.Add(this.button9);
            this.groupBox8.Controls.Add(this.Name3);
            this.groupBox8.Controls.Add(this.label14);
            this.groupBox8.Controls.Add(this.Npptext3);
            this.groupBox8.Controls.Add(this.TextTimer3);
            this.groupBox8.Controls.Add(this.label16);
            this.groupBox8.Controls.Add(this.Ntext3);
            this.groupBox8.Location = new System.Drawing.Point(727, 424);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(212, 297);
            this.groupBox8.TabIndex = 15;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Третий участник";
            // 
            // PolBox3
            // 
            this.PolBox3.Location = new System.Drawing.Point(0, 277);
            this.PolBox3.Name = "PolBox3";
            this.PolBox3.Size = new System.Drawing.Size(10, 20);
            this.PolBox3.TabIndex = 28;
            this.PolBox3.Visible = false;
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(74, 266);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(75, 23);
            this.button9.TabIndex = 22;
            this.button9.Text = "Стоп";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // Name3
            // 
            this.Name3.BackColor = System.Drawing.SystemColors.Window;
            this.Name3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Name3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Name3.Location = new System.Drawing.Point(6, 71);
            this.Name3.Multiline = true;
            this.Name3.Name = "Name3";
            this.Name3.ReadOnly = true;
            this.Name3.Size = new System.Drawing.Size(200, 103);
            this.Name3.TabIndex = 14;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(44, 19);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(38, 13);
            this.label14.TabIndex = 3;
            this.label14.Text = "№ п/п";
            // 
            // Npptext3
            // 
            this.Npptext3.BackColor = System.Drawing.SystemColors.Window;
            this.Npptext3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Npptext3.Location = new System.Drawing.Point(88, 19);
            this.Npptext3.Name = "Npptext3";
            this.Npptext3.ReadOnly = true;
            this.Npptext3.Size = new System.Drawing.Size(82, 13);
            this.Npptext3.TabIndex = 20;
            // 
            // TextTimer3
            // 
            this.TextTimer3.AutoSize = true;
            this.TextTimer3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TextTimer3.Location = new System.Drawing.Point(-2, 232);
            this.TextTimer3.Name = "TextTimer3";
            this.TextTimer3.Size = new System.Drawing.Size(218, 31);
            this.TextTimer3.TabIndex = 1;
            this.TextTimer3.Text = "00:00:00.000000";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(10, 42);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(72, 13);
            this.label16.TabIndex = 4;
            this.label16.Text = "№ участника";
            // 
            // Ntext3
            // 
            this.Ntext3.BackColor = System.Drawing.SystemColors.Window;
            this.Ntext3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Ntext3.Location = new System.Drawing.Point(88, 42);
            this.Ntext3.Name = "Ntext3";
            this.Ntext3.ReadOnly = true;
            this.Ntext3.Size = new System.Drawing.Size(82, 13);
            this.Ntext3.TabIndex = 19;
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.PolBox2);
            this.groupBox4.Controls.Add(this.button7);
            this.groupBox4.Controls.Add(this.Name2);
            this.groupBox4.Controls.Add(this.TextTimer2);
            this.groupBox4.Controls.Add(this.label9);
            this.groupBox4.Controls.Add(this.Npptext2);
            this.groupBox4.Controls.Add(this.label11);
            this.groupBox4.Controls.Add(this.Ntext2);
            this.groupBox4.Location = new System.Drawing.Point(218, 424);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(212, 297);
            this.groupBox4.TabIndex = 14;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Второй участник";
            // 
            // PolBox2
            // 
            this.PolBox2.Location = new System.Drawing.Point(0, 277);
            this.PolBox2.Name = "PolBox2";
            this.PolBox2.Size = new System.Drawing.Size(10, 20);
            this.PolBox2.TabIndex = 24;
            this.PolBox2.Visible = false;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(65, 266);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 23);
            this.button7.TabIndex = 22;
            this.button7.Text = "Стоп";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // Name2
            // 
            this.Name2.BackColor = System.Drawing.SystemColors.Window;
            this.Name2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Name2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Name2.Location = new System.Drawing.Point(6, 71);
            this.Name2.Multiline = true;
            this.Name2.Name = "Name2";
            this.Name2.ReadOnly = true;
            this.Name2.Size = new System.Drawing.Size(200, 103);
            this.Name2.TabIndex = 14;
            // 
            // TextTimer2
            // 
            this.TextTimer2.AutoSize = true;
            this.TextTimer2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TextTimer2.Location = new System.Drawing.Point(-2, 232);
            this.TextTimer2.Name = "TextTimer2";
            this.TextTimer2.Size = new System.Drawing.Size(218, 31);
            this.TextTimer2.TabIndex = 1;
            this.TextTimer2.Text = "00:00:00.000000";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(44, 19);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(38, 13);
            this.label9.TabIndex = 3;
            this.label9.Text = "№ п/п";
            // 
            // Npptext2
            // 
            this.Npptext2.BackColor = System.Drawing.SystemColors.Window;
            this.Npptext2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Npptext2.Location = new System.Drawing.Point(88, 19);
            this.Npptext2.Name = "Npptext2";
            this.Npptext2.ReadOnly = true;
            this.Npptext2.Size = new System.Drawing.Size(82, 13);
            this.Npptext2.TabIndex = 20;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(10, 42);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 13);
            this.label11.TabIndex = 4;
            this.label11.Text = "№ участника";
            // 
            // Ntext2
            // 
            this.Ntext2.BackColor = System.Drawing.SystemColors.Window;
            this.Ntext2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Ntext2.Location = new System.Drawing.Point(88, 42);
            this.Ntext2.Name = "Ntext2";
            this.Ntext2.ReadOnly = true;
            this.Ntext2.Size = new System.Drawing.Size(82, 13);
            this.Ntext2.TabIndex = 19;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(1013, 6);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(94, 94);
            this.pictureBox2.TabIndex = 13;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(41, 6);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(94, 94);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox3.Controls.Add(this.PolBox1);
            this.groupBox3.Controls.Add(this.StopButton1);
            this.groupBox3.Controls.Add(this.Name1);
            this.groupBox3.Controls.Add(this.TextTimer1);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.Npptext1);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.Ntext1);
            this.groupBox3.Location = new System.Drawing.Point(0, 424);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(212, 297);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Первый участник";
            // 
            // PolBox1
            // 
            this.PolBox1.Location = new System.Drawing.Point(0, 276);
            this.PolBox1.Name = "PolBox1";
            this.PolBox1.Size = new System.Drawing.Size(10, 20);
            this.PolBox1.TabIndex = 23;
            this.PolBox1.Visible = false;
            // 
            // StopButton1
            // 
            this.StopButton1.Location = new System.Drawing.Point(64, 266);
            this.StopButton1.Name = "StopButton1";
            this.StopButton1.Size = new System.Drawing.Size(75, 23);
            this.StopButton1.TabIndex = 22;
            this.StopButton1.Text = "Стоп";
            this.StopButton1.UseVisualStyleBackColor = true;
            this.StopButton1.Click += new System.EventHandler(this.StopButton1_Click);
            // 
            // Name1
            // 
            this.Name1.BackColor = System.Drawing.SystemColors.Window;
            this.Name1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Name1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Name1.Location = new System.Drawing.Point(6, 71);
            this.Name1.Multiline = true;
            this.Name1.Name = "Name1";
            this.Name1.ReadOnly = true;
            this.Name1.Size = new System.Drawing.Size(200, 103);
            this.Name1.TabIndex = 14;
            // 
            // TextTimer1
            // 
            this.TextTimer1.AutoSize = true;
            this.TextTimer1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.TextTimer1.Location = new System.Drawing.Point(-2, 232);
            this.TextTimer1.Name = "TextTimer1";
            this.TextTimer1.Size = new System.Drawing.Size(218, 31);
            this.TextTimer1.TabIndex = 1;
            this.TextTimer1.Text = "00:00:00.000000";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(44, 19);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 13);
            this.label5.TabIndex = 3;
            this.label5.Text = "№ п/п";
            // 
            // Npptext1
            // 
            this.Npptext1.BackColor = System.Drawing.SystemColors.Window;
            this.Npptext1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Npptext1.Location = new System.Drawing.Point(88, 19);
            this.Npptext1.Name = "Npptext1";
            this.Npptext1.ReadOnly = true;
            this.Npptext1.Size = new System.Drawing.Size(82, 13);
            this.Npptext1.TabIndex = 20;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(12, 45);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "№ участника";
            // 
            // Ntext1
            // 
            this.Ntext1.BackColor = System.Drawing.SystemColors.Window;
            this.Ntext1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Ntext1.Location = new System.Drawing.Point(88, 45);
            this.Ntext1.Name = "Ntext1";
            this.Ntext1.ReadOnly = true;
            this.Ntext1.Size = new System.Drawing.Size(82, 13);
            this.Ntext1.TabIndex = 19;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox9);
            this.tabPage2.Controls.Add(this.groupBox7);
            this.tabPage2.Controls.Add(this.groupBox6);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1157, 721);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Настройки";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox9
            // 
            this.groupBox9.Location = new System.Drawing.Point(0, 286);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(214, 147);
            this.groupBox9.TabIndex = 18;
            this.groupBox9.TabStop = false;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.radioButtonComands);
            this.groupBox7.Controls.Add(this.radioButtonWoman);
            this.groupBox7.Controls.Add(this.radioButtonMan);
            this.groupBox7.Location = new System.Drawing.Point(0, 187);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(214, 93);
            this.groupBox7.TabIndex = 17;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Таблица лидеров";
            // 
            // radioButtonComands
            // 
            this.radioButtonComands.AutoSize = true;
            this.radioButtonComands.Location = new System.Drawing.Point(9, 67);
            this.radioButtonComands.Name = "radioButtonComands";
            this.radioButtonComands.Size = new System.Drawing.Size(72, 17);
            this.radioButtonComands.TabIndex = 19;
            this.radioButtonComands.TabStop = true;
            this.radioButtonComands.Text = "Команды";
            this.radioButtonComands.UseVisualStyleBackColor = true;
            this.radioButtonComands.CheckedChanged += new System.EventHandler(this.radioButtonComands_CheckedChanged);
            // 
            // radioButtonWoman
            // 
            this.radioButtonWoman.AutoSize = true;
            this.radioButtonWoman.Location = new System.Drawing.Point(9, 19);
            this.radioButtonWoman.Name = "radioButtonWoman";
            this.radioButtonWoman.Size = new System.Drawing.Size(77, 17);
            this.radioButtonWoman.TabIndex = 17;
            this.radioButtonWoman.TabStop = true;
            this.radioButtonWoman.Text = "Женщины";
            this.radioButtonWoman.UseVisualStyleBackColor = true;
            this.radioButtonWoman.CheckedChanged += new System.EventHandler(this.radioButtonWoman_CheckedChanged);
            // 
            // radioButtonMan
            // 
            this.radioButtonMan.AutoSize = true;
            this.radioButtonMan.Location = new System.Drawing.Point(9, 43);
            this.radioButtonMan.Name = "radioButtonMan";
            this.radioButtonMan.Size = new System.Drawing.Size(72, 17);
            this.radioButtonMan.TabIndex = 18;
            this.radioButtonMan.TabStop = true;
            this.radioButtonMan.Text = "Мужчины";
            this.radioButtonMan.UseVisualStyleBackColor = true;
            this.radioButtonMan.CheckedChanged += new System.EventHandler(this.radioButtonMan_CheckedChanged);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.progressBar1);
            this.groupBox6.Controls.Add(this.tabControl2);
            this.groupBox6.Controls.Add(this.button5);
            this.groupBox6.Controls.Add(this.ToTimerButton);
            this.groupBox6.Controls.Add(this.ExportButton);
            this.groupBox6.Location = new System.Drawing.Point(220, 0);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(937, 721);
            this.groupBox6.TabIndex = 16;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Таблица участников";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(602, 16);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(250, 23);
            this.progressBar1.TabIndex = 17;
            // 
            // tabControl2
            // 
            this.tabControl2.Controls.Add(this.Woman);
            this.tabControl2.Controls.Add(this.Man);
            this.tabControl2.Controls.Add(this.Commands);
            this.tabControl2.Location = new System.Drawing.Point(6, 45);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(931, 676);
            this.tabControl2.TabIndex = 16;
            // 
            // Woman
            // 
            this.Woman.Controls.Add(this.dataGridView1);
            this.Woman.Location = new System.Drawing.Point(4, 22);
            this.Woman.Name = "Woman";
            this.Woman.Padding = new System.Windows.Forms.Padding(3);
            this.Woman.Size = new System.Drawing.Size(923, 650);
            this.Woman.TabIndex = 0;
            this.Woman.Text = "Женщины";
            this.Woman.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(923, 650);
            this.dataGridView1.TabIndex = 10;
            // 
            // Man
            // 
            this.Man.Controls.Add(this.dataGridView2);
            this.Man.Location = new System.Drawing.Point(4, 22);
            this.Man.Name = "Man";
            this.Man.Padding = new System.Windows.Forms.Padding(3);
            this.Man.Size = new System.Drawing.Size(923, 650);
            this.Man.TabIndex = 1;
            this.Man.Text = "Мужчины";
            this.Man.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(0, 0);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(924, 651);
            this.dataGridView2.TabIndex = 0;
            // 
            // Commands
            // 
            this.Commands.Controls.Add(this.dataGridView3);
            this.Commands.Location = new System.Drawing.Point(4, 22);
            this.Commands.Name = "Commands";
            this.Commands.Padding = new System.Windows.Forms.Padding(3);
            this.Commands.Size = new System.Drawing.Size(923, 650);
            this.Commands.TabIndex = 2;
            this.Commands.Text = "Команды";
            this.Commands.UseVisualStyleBackColor = true;
            // 
            // dataGridView3
            // 
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(0, 0);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.Size = new System.Drawing.Size(924, 651);
            this.dataGridView3.TabIndex = 0;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(6, 16);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 15;
            this.button5.Text = "Импорт";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // ToTimerButton
            // 
            this.ToTimerButton.Enabled = false;
            this.ToTimerButton.Location = new System.Drawing.Point(87, 16);
            this.ToTimerButton.Name = "ToTimerButton";
            this.ToTimerButton.Size = new System.Drawing.Size(75, 23);
            this.ToTimerButton.TabIndex = 13;
            this.ToTimerButton.Text = "Перенести";
            this.ToTimerButton.UseVisualStyleBackColor = true;
            this.ToTimerButton.Click += new System.EventHandler(this.ToTimerButton_Click);
            // 
            // ExportButton
            // 
            this.ExportButton.Location = new System.Drawing.Point(859, 16);
            this.ExportButton.Name = "ExportButton";
            this.ExportButton.Size = new System.Drawing.Size(75, 23);
            this.ExportButton.TabIndex = 11;
            this.ExportButton.Text = "Экспорт";
            this.ExportButton.UseVisualStyleBackColor = true;
            this.ExportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.stopsoundfilepath);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.startsoundfilepath);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Location = new System.Drawing.Point(0, 47);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(214, 133);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sounds";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(157, 76);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(39, 20);
            this.button4.TabIndex = 11;
            this.button4.Text = "Play";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(112, 76);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(39, 20);
            this.button3.TabIndex = 10;
            this.button3.Text = "Load";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // stopsoundfilepath
            // 
            this.stopsoundfilepath.Location = new System.Drawing.Point(6, 76);
            this.stopsoundfilepath.Name = "stopsoundfilepath";
            this.stopsoundfilepath.ReadOnly = true;
            this.stopsoundfilepath.Size = new System.Drawing.Size(100, 20);
            this.stopsoundfilepath.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 60);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Stop";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "PreStart";
            // 
            // startsoundfilepath
            // 
            this.startsoundfilepath.Location = new System.Drawing.Point(6, 33);
            this.startsoundfilepath.Name = "startsoundfilepath";
            this.startsoundfilepath.ReadOnly = true;
            this.startsoundfilepath.Size = new System.Drawing.Size(100, 20);
            this.startsoundfilepath.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(112, 33);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(39, 20);
            this.button1.TabIndex = 4;
            this.button1.Text = "Load";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(157, 33);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(39, 20);
            this.button2.TabIndex = 6;
            this.button2.Text = "Play";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(214, 41);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Status";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1189, 771);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Хронометр 2";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.tabControl2.ResumeLayout(false);
            this.Woman.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.Man.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.Commands.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label TextTimer;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button StopButton;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox startsoundfilepath;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox stopsoundfilepath;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button ToTimerButton;
        private System.Windows.Forms.Button ExportButton;
        private System.Windows.Forms.Label TextTimer1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Name1;
        private System.Windows.Forms.TextBox Ntext1;
        private System.Windows.Forms.TextBox Npptext1;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.TextBox Name4;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox Npptext4;
        private System.Windows.Forms.Label TextTimer4;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox Ntext4;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.TextBox Name3;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox Npptext3;
        private System.Windows.Forms.Label TextTimer3;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox Ntext3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox Name2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox Npptext2;
        private System.Windows.Forms.Label TextTimer2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox Ntext2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button StopButton1;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.RadioButton radioButtonComands;
        private System.Windows.Forms.RadioButton radioButtonMan;
        private System.Windows.Forms.RadioButton radioButtonWoman;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.TabControl tabControl2;
        private System.Windows.Forms.TabPage Woman;
        private System.Windows.Forms.TabPage Man;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.TabPage Commands;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.TextBox LeaderBox;
        private System.Windows.Forms.TextBox titleBox;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.TextBox PolBox4;
        private System.Windows.Forms.TextBox PolBox3;
        private System.Windows.Forms.TextBox PolBox2;
        private System.Windows.Forms.TextBox PolBox1;
    }
}

