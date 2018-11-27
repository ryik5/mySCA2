namespace mySCA2
{
    partial class FormPersonViewerSCA
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormPersonViewerSCA));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.comboBoxFio = new System.Windows.Forms.ComboBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.numUpDownHour = new System.Windows.Forms.NumericUpDown();
            this.numUpDownMinute = new System.Windows.Forms.NumericUpDown();
            this.groupBoxRemoveDays = new System.Windows.Forms.GroupBox();
            this.checkBoxReEnter = new System.Windows.Forms.CheckBox();
            this.checkBoxStartWorkInTime = new System.Windows.Forms.CheckBox();
            this.checkBoxCelebrate = new System.Windows.Forms.CheckBox();
            this.checkBoxWeekend = new System.Windows.Forms.CheckBox();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.StatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.ProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.StatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.textBoxGroup = new System.Windows.Forms.TextBox();
            this.textBoxNav = new System.Windows.Forms.TextBox();
            this.textBoxGroupDescription = new System.Windows.Forms.TextBox();
            this.textBoxFIO = new System.Windows.Forms.TextBox();
            this.labelGroupDescription = new System.Windows.Forms.Label();
            this.groupBoxPeriod = new System.Windows.Forms.GroupBox();
            this.labelInfoStart = new System.Windows.Forms.Label();
            this.labelInfoEnd = new System.Windows.Forms.Label();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.FunctionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.GetFioItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExitItem = new System.Windows.Forms.ToolStripMenuItem();
            this.GroupsMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PersonOrGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CreateGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddPersonToGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.MembersGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeletePersonFromGroupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ViewMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExportIntoExcelItem = new System.Windows.Forms.ToolStripMenuItem();
            this.VisualItem = new System.Windows.Forms.ToolStripMenuItem();
            this.VisualWorkedTimeItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SelectColorMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ColorRegistrationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.BlueItem = new System.Windows.Forms.ToolStripMenuItem();
            this.RedItem = new System.Windows.Forms.ToolStripMenuItem();
            this.YellowItem = new System.Windows.Forms.ToolStripMenuItem();
            this.GreenItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ReportsItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AnualDatesMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.EnterEditAnualItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddAnualDateItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteAnualDateItem = new System.Windows.Forms.ToolStripMenuItem();
            this.QuickSettingsItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SettingsEquipmentItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SettingsProgrammItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SettingsOtherItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ClearReportItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ClearDataItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ClearAllItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TestCryptionItem = new System.Windows.Forms.ToolStripMenuItem();
            this.QuickLoadDataItem = new System.Windows.Forms.ToolStripMenuItem();
            this.QuickFilterItem = new System.Windows.Forms.ToolStripMenuItem();
            this.UpdateControllingItem = new System.Windows.Forms.ToolStripMenuItem();
            this.HelpMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.HelpAboutItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SetupItem = new System.Windows.Forms.ToolStripMenuItem();
            this.HelpSystemItem = new System.Windows.Forms.ToolStripMenuItem();
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.labelHour = new System.Windows.Forms.Label();
            this.labelMinute = new System.Windows.Forms.Label();
            this.groupBoxTime = new System.Windows.Forms.GroupBox();
            this.monthCalendar = new System.Windows.Forms.MonthCalendar();
            this.labelGroup = new System.Windows.Forms.Label();
            this.panelView = new System.Windows.Forms.Panel();
            this.groupBoxProperties = new System.Windows.Forms.GroupBox();
            this.buttonPropertiesCancel = new System.Windows.Forms.Button();
            this.buttonPropertiesSave = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownHour)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownMinute)).BeginInit();
            this.groupBoxRemoveDays.SuspendLayout();
            this.statusStrip.SuspendLayout();
            this.groupBoxPeriod.SuspendLayout();
            this.menuStrip.SuspendLayout();
            this.groupBoxTime.SuspendLayout();
            this.panelView.SuspendLayout();
            this.groupBoxProperties.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.SteelBlue;
            this.dataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            resources.ApplyResources(this.dataGridView1, "dataGridView1");
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedHeaders;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.LightBlue;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.DarkSlateGray;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.StandardTab = true;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // comboBoxFio
            // 
            this.comboBoxFio.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBoxFio.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxFio.FormattingEnabled = true;
            resources.ApplyResources(this.comboBoxFio, "comboBoxFio");
            this.comboBoxFio.Name = "comboBoxFio";
            this.comboBoxFio.Sorted = true;
            this.toolTip1.SetToolTip(this.comboBoxFio, resources.GetString("comboBoxFio.ToolTip"));
            this.comboBoxFio.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // numUpDownHour
            // 
            resources.ApplyResources(this.numUpDownHour, "numUpDownHour");
            this.numUpDownHour.Maximum = new decimal(new int[] {
            23,
            0,
            0,
            0});
            this.numUpDownHour.Name = "numUpDownHour";
            this.toolTip1.SetToolTip(this.numUpDownHour, resources.GetString("numUpDownHour.ToolTip"));
            // 
            // numUpDownMinute
            // 
            resources.ApplyResources(this.numUpDownMinute, "numUpDownMinute");
            this.numUpDownMinute.Maximum = new decimal(new int[] {
            59,
            0,
            0,
            0});
            this.numUpDownMinute.Name = "numUpDownMinute";
            this.toolTip1.SetToolTip(this.numUpDownMinute, resources.GetString("numUpDownMinute.ToolTip"));
            // 
            // groupBoxRemoveDays
            // 
            resources.ApplyResources(this.groupBoxRemoveDays, "groupBoxRemoveDays");
            this.groupBoxRemoveDays.Controls.Add(this.checkBoxReEnter);
            this.groupBoxRemoveDays.Controls.Add(this.checkBoxStartWorkInTime);
            this.groupBoxRemoveDays.Controls.Add(this.checkBoxCelebrate);
            this.groupBoxRemoveDays.Controls.Add(this.checkBoxWeekend);
            this.groupBoxRemoveDays.Name = "groupBoxRemoveDays";
            this.groupBoxRemoveDays.TabStop = false;
            this.toolTip1.SetToolTip(this.groupBoxRemoveDays, resources.GetString("groupBoxRemoveDays.ToolTip"));
            // 
            // checkBoxReEnter
            // 
            resources.ApplyResources(this.checkBoxReEnter, "checkBoxReEnter");
            this.checkBoxReEnter.Name = "checkBoxReEnter";
            this.toolTip1.SetToolTip(this.checkBoxReEnter, resources.GetString("checkBoxReEnter.ToolTip"));
            this.checkBoxReEnter.UseVisualStyleBackColor = true;
            this.checkBoxReEnter.CheckStateChanged += new System.EventHandler(this.checkBoxReEnter_CheckStateChanged);
            // 
            // checkBoxStartWorkInTime
            // 
            resources.ApplyResources(this.checkBoxStartWorkInTime, "checkBoxStartWorkInTime");
            this.checkBoxStartWorkInTime.Name = "checkBoxStartWorkInTime";
            this.toolTip1.SetToolTip(this.checkBoxStartWorkInTime, resources.GetString("checkBoxStartWorkInTime.ToolTip"));
            this.checkBoxStartWorkInTime.UseVisualStyleBackColor = true;
            this.checkBoxStartWorkInTime.CheckStateChanged += new System.EventHandler(this.checkBoxStartWorkInTime_CheckStateChanged);
            // 
            // checkBoxCelebrate
            // 
            resources.ApplyResources(this.checkBoxCelebrate, "checkBoxCelebrate");
            this.checkBoxCelebrate.Name = "checkBoxCelebrate";
            this.toolTip1.SetToolTip(this.checkBoxCelebrate, resources.GetString("checkBoxCelebrate.ToolTip"));
            this.checkBoxCelebrate.UseVisualStyleBackColor = true;
            this.checkBoxCelebrate.CheckStateChanged += new System.EventHandler(this.checkBoxCelebrate_CheckStateChanged);
            // 
            // checkBoxWeekend
            // 
            resources.ApplyResources(this.checkBoxWeekend, "checkBoxWeekend");
            this.checkBoxWeekend.Name = "checkBoxWeekend";
            this.toolTip1.SetToolTip(this.checkBoxWeekend, resources.GetString("checkBoxWeekend.ToolTip"));
            this.checkBoxWeekend.UseVisualStyleBackColor = true;
            this.checkBoxWeekend.CheckStateChanged += new System.EventHandler(this.checkBoxWeekend_CheckStateChanged);
            // 
            // statusStrip
            // 
            this.statusStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel1,
            this.ProgressBar1,
            this.StatusLabel2});
            resources.ApplyResources(this.statusStrip, "statusStrip");
            this.statusStrip.Name = "statusStrip";
            this.toolTip1.SetToolTip(this.statusStrip, resources.GetString("statusStrip.ToolTip"));
            // 
            // StatusLabel1
            // 
            this.StatusLabel1.Name = "StatusLabel1";
            resources.ApplyResources(this.StatusLabel1, "StatusLabel1");
            // 
            // ProgressBar1
            // 
            this.ProgressBar1.Name = "ProgressBar1";
            resources.ApplyResources(this.ProgressBar1, "ProgressBar1");
            // 
            // StatusLabel2
            // 
            this.StatusLabel2.Name = "StatusLabel2";
            resources.ApplyResources(this.StatusLabel2, "StatusLabel2");
            // 
            // dateTimePickerStart
            // 
            resources.ApplyResources(this.dateTimePickerStart, "dateTimePickerStart");
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.toolTip1.SetToolTip(this.dateTimePickerStart, resources.GetString("dateTimePickerStart.ToolTip"));
            this.dateTimePickerStart.CloseUp += new System.EventHandler(this.dateTimePickerStart_CloseUp);
            // 
            // dateTimePickerEnd
            // 
            resources.ApplyResources(this.dateTimePickerEnd, "dateTimePickerEnd");
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.toolTip1.SetToolTip(this.dateTimePickerEnd, resources.GetString("dateTimePickerEnd.ToolTip"));
            this.dateTimePickerEnd.CloseUp += new System.EventHandler(this.dateTimePickerEnd_CloseUp);
            // 
            // textBoxGroup
            // 
            resources.ApplyResources(this.textBoxGroup, "textBoxGroup");
            this.textBoxGroup.Name = "textBoxGroup";
            this.toolTip1.SetToolTip(this.textBoxGroup, resources.GetString("textBoxGroup.ToolTip"));
            this.textBoxGroup.TextChanged += new System.EventHandler(this.textBoxGroup_TextChanged);
            // 
            // textBoxNav
            // 
            resources.ApplyResources(this.textBoxNav, "textBoxNav");
            this.textBoxNav.Name = "textBoxNav";
            this.textBoxNav.ReadOnly = true;
            this.toolTip1.SetToolTip(this.textBoxNav, resources.GetString("textBoxNav.ToolTip"));
            // 
            // textBoxGroupDescription
            // 
            resources.ApplyResources(this.textBoxGroupDescription, "textBoxGroupDescription");
            this.textBoxGroupDescription.Name = "textBoxGroupDescription";
            this.toolTip1.SetToolTip(this.textBoxGroupDescription, resources.GetString("textBoxGroupDescription.ToolTip"));
            // 
            // textBoxFIO
            // 
            resources.ApplyResources(this.textBoxFIO, "textBoxFIO");
            this.textBoxFIO.Name = "textBoxFIO";
            this.textBoxFIO.ReadOnly = true;
            this.toolTip1.SetToolTip(this.textBoxFIO, resources.GetString("textBoxFIO.ToolTip"));
            // 
            // labelGroupDescription
            // 
            resources.ApplyResources(this.labelGroupDescription, "labelGroupDescription");
            this.labelGroupDescription.Name = "labelGroupDescription";
            this.toolTip1.SetToolTip(this.labelGroupDescription, resources.GetString("labelGroupDescription.ToolTip"));
            // 
            // groupBoxPeriod
            // 
            resources.ApplyResources(this.groupBoxPeriod, "groupBoxPeriod");
            this.groupBoxPeriod.Controls.Add(this.dateTimePickerEnd);
            this.groupBoxPeriod.Controls.Add(this.dateTimePickerStart);
            this.groupBoxPeriod.Controls.Add(this.labelInfoStart);
            this.groupBoxPeriod.Controls.Add(this.labelInfoEnd);
            this.groupBoxPeriod.Name = "groupBoxPeriod";
            this.groupBoxPeriod.TabStop = false;
            this.toolTip1.SetToolTip(this.groupBoxPeriod, resources.GetString("groupBoxPeriod.ToolTip"));
            // 
            // labelInfoStart
            // 
            resources.ApplyResources(this.labelInfoStart, "labelInfoStart");
            this.labelInfoStart.Name = "labelInfoStart";
            // 
            // labelInfoEnd
            // 
            resources.ApplyResources(this.labelInfoEnd, "labelInfoEnd");
            this.labelInfoEnd.Name = "labelInfoEnd";
            // 
            // menuStrip
            // 
            this.menuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.FunctionMenuItem,
            this.GroupsMenuItem,
            this.ViewMenuItem,
            this.AnualDatesMenuItem,
            this.QuickSettingsItem,
            this.QuickLoadDataItem,
            this.QuickFilterItem,
            this.UpdateControllingItem,
            this.HelpMenuItem});
            resources.ApplyResources(this.menuStrip, "menuStrip");
            this.menuStrip.Name = "menuStrip";
            // 
            // FunctionMenuItem
            // 
            this.FunctionMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.GetFioItem,
            this.ExitItem});
            this.FunctionMenuItem.Name = "FunctionMenuItem";
            resources.ApplyResources(this.FunctionMenuItem, "FunctionMenuItem");
            // 
            // GetFioItem
            // 
            this.GetFioItem.Name = "GetFioItem";
            resources.ApplyResources(this.GetFioItem, "GetFioItem");
            this.GetFioItem.Click += new System.EventHandler(this.buttonGetFio_Click);
            // 
            // ExitItem
            // 
            this.ExitItem.Name = "ExitItem";
            resources.ApplyResources(this.ExitItem, "ExitItem");
            this.ExitItem.Click += new System.EventHandler(this.ApplicationExit);
            // 
            // GroupsMenuItem
            // 
            this.GroupsMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PersonOrGroupItem,
            this.CreateGroupItem,
            this.AddPersonToGroupItem,
            this.ListGroupItem,
            this.MembersGroupItem,
            this.DeleteGroupItem,
            this.DeletePersonFromGroupItem});
            this.GroupsMenuItem.Name = "GroupsMenuItem";
            resources.ApplyResources(this.GroupsMenuItem, "GroupsMenuItem");
            // 
            // PersonOrGroupItem
            // 
            this.PersonOrGroupItem.Name = "PersonOrGroupItem";
            resources.ApplyResources(this.PersonOrGroupItem, "PersonOrGroupItem");
            this.PersonOrGroupItem.Click += new System.EventHandler(this.PersonOrGroupItem_Click);
            // 
            // CreateGroupItem
            // 
            this.CreateGroupItem.Name = "CreateGroupItem";
            resources.ApplyResources(this.CreateGroupItem, "CreateGroupItem");
            this.CreateGroupItem.Click += new System.EventHandler(this.CreateGroupItem_Click);
            // 
            // AddPersonToGroupItem
            // 
            this.AddPersonToGroupItem.Name = "AddPersonToGroupItem";
            resources.ApplyResources(this.AddPersonToGroupItem, "AddPersonToGroupItem");
            this.AddPersonToGroupItem.Click += new System.EventHandler(this.AddPersonToGroupItem_Click);
            // 
            // ListGroupItem
            // 
            this.ListGroupItem.Name = "ListGroupItem";
            resources.ApplyResources(this.ListGroupItem, "ListGroupItem");
            this.ListGroupItem.Click += new System.EventHandler(this.ListGroupItem_Click);
            // 
            // MembersGroupItem
            // 
            this.MembersGroupItem.Name = "MembersGroupItem";
            resources.ApplyResources(this.MembersGroupItem, "MembersGroupItem");
            this.MembersGroupItem.Click += new System.EventHandler(this.MembersGroupItem_Click);
            // 
            // DeleteGroupItem
            // 
            this.DeleteGroupItem.Name = "DeleteGroupItem";
            resources.ApplyResources(this.DeleteGroupItem, "DeleteGroupItem");
            this.DeleteGroupItem.Click += new System.EventHandler(this.DeleteGroupItem_Click);
            // 
            // DeletePersonFromGroupItem
            // 
            this.DeletePersonFromGroupItem.Name = "DeletePersonFromGroupItem";
            resources.ApplyResources(this.DeletePersonFromGroupItem, "DeletePersonFromGroupItem");
            this.DeletePersonFromGroupItem.Click += new System.EventHandler(this.DeletePersonFromGroupItem_Click);
            // 
            // ViewMenuItem
            // 
            this.ViewMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ExportIntoExcelItem,
            this.VisualItem,
            this.VisualWorkedTimeItem,
            this.SelectColorMenuItem,
            this.ReportsItem});
            this.ViewMenuItem.Name = "ViewMenuItem";
            resources.ApplyResources(this.ViewMenuItem, "ViewMenuItem");
            this.ViewMenuItem.Click += new System.EventHandler(this.ViewMenuItem_Click);
            // 
            // ExportIntoExcelItem
            // 
            this.ExportIntoExcelItem.Name = "ExportIntoExcelItem";
            resources.ApplyResources(this.ExportIntoExcelItem, "ExportIntoExcelItem");
            this.ExportIntoExcelItem.Click += new System.EventHandler(this.buttonExport_Click);
            // 
            // VisualItem
            // 
            this.VisualItem.Name = "VisualItem";
            resources.ApplyResources(this.VisualItem, "VisualItem");
            this.VisualItem.Click += new System.EventHandler(this.VisualItem_Click);
            // 
            // VisualWorkedTimeItem
            // 
            this.VisualWorkedTimeItem.Name = "VisualWorkedTimeItem";
            resources.ApplyResources(this.VisualWorkedTimeItem, "VisualWorkedTimeItem");
            this.VisualWorkedTimeItem.Click += new System.EventHandler(this.VisualWorkedTimeItem_Click);
            // 
            // SelectColorMenuItem
            // 
            this.SelectColorMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ColorRegistrationMenuItem});
            this.SelectColorMenuItem.Name = "SelectColorMenuItem";
            resources.ApplyResources(this.SelectColorMenuItem, "SelectColorMenuItem");
            // 
            // ColorRegistrationMenuItem
            // 
            this.ColorRegistrationMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.BlueItem,
            this.RedItem,
            this.YellowItem,
            this.GreenItem});
            this.ColorRegistrationMenuItem.Name = "ColorRegistrationMenuItem";
            resources.ApplyResources(this.ColorRegistrationMenuItem, "ColorRegistrationMenuItem");
            // 
            // BlueItem
            // 
            this.BlueItem.Name = "BlueItem";
            resources.ApplyResources(this.BlueItem, "BlueItem");
            this.BlueItem.Click += new System.EventHandler(this.BlueItem_Click);
            // 
            // RedItem
            // 
            this.RedItem.Name = "RedItem";
            resources.ApplyResources(this.RedItem, "RedItem");
            this.RedItem.Click += new System.EventHandler(this.RedItem_Click);
            // 
            // YellowItem
            // 
            this.YellowItem.Name = "YellowItem";
            resources.ApplyResources(this.YellowItem, "YellowItem");
            this.YellowItem.Click += new System.EventHandler(this.YellowItem_Click);
            // 
            // GreenItem
            // 
            this.GreenItem.Name = "GreenItem";
            resources.ApplyResources(this.GreenItem, "GreenItem");
            this.GreenItem.Click += new System.EventHandler(this.GreenItem_Click);
            // 
            // ReportsItem
            // 
            this.ReportsItem.Name = "ReportsItem";
            resources.ApplyResources(this.ReportsItem, "ReportsItem");
            this.ReportsItem.Click += new System.EventHandler(this.ReportsItem_Click);
            // 
            // AnualDatesMenuItem
            // 
            this.AnualDatesMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.EnterEditAnualItem,
            this.AddAnualDateItem,
            this.DeleteAnualDateItem});
            this.AnualDatesMenuItem.Name = "AnualDatesMenuItem";
            resources.ApplyResources(this.AnualDatesMenuItem, "AnualDatesMenuItem");
            // 
            // EnterEditAnualItem
            // 
            this.EnterEditAnualItem.Name = "EnterEditAnualItem";
            resources.ApplyResources(this.EnterEditAnualItem, "EnterEditAnualItem");
            this.EnterEditAnualItem.Click += new System.EventHandler(this.EnterEditAnualItem_Click);
            // 
            // AddAnualDateItem
            // 
            this.AddAnualDateItem.Name = "AddAnualDateItem";
            resources.ApplyResources(this.AddAnualDateItem, "AddAnualDateItem");
            this.AddAnualDateItem.Click += new System.EventHandler(this.AddAnualDateItem_Click);
            // 
            // DeleteAnualDateItem
            // 
            this.DeleteAnualDateItem.Name = "DeleteAnualDateItem";
            resources.ApplyResources(this.DeleteAnualDateItem, "DeleteAnualDateItem");
            this.DeleteAnualDateItem.Click += new System.EventHandler(this.DeleteAnualDateItem_Click);
            // 
            // QuickSettingsItem
            // 
            this.QuickSettingsItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.SettingsEquipmentItem,
            this.SettingsProgrammItem,
            this.SettingsOtherItem,
            this.ClearReportItem,
            this.ClearDataItem,
            this.ClearAllItem,
            this.TestCryptionItem});
            this.QuickSettingsItem.Name = "QuickSettingsItem";
            resources.ApplyResources(this.QuickSettingsItem, "QuickSettingsItem");
            // 
            // SettingsEquipmentItem
            // 
            this.SettingsEquipmentItem.Name = "SettingsEquipmentItem";
            resources.ApplyResources(this.SettingsEquipmentItem, "SettingsEquipmentItem");
            // 
            // SettingsProgrammItem
            // 
            this.SettingsProgrammItem.Name = "SettingsProgrammItem";
            resources.ApplyResources(this.SettingsProgrammItem, "SettingsProgrammItem");
            this.SettingsProgrammItem.Click += new System.EventHandler(this.SettingsProgrammItem_Click);
            // 
            // SettingsOtherItem
            // 
            this.SettingsOtherItem.Name = "SettingsOtherItem";
            resources.ApplyResources(this.SettingsOtherItem, "SettingsOtherItem");
            // 
            // ClearReportItem
            // 
            this.ClearReportItem.Name = "ClearReportItem";
            resources.ApplyResources(this.ClearReportItem, "ClearReportItem");
            this.ClearReportItem.Click += new System.EventHandler(this.ClearReportItem_Click);
            // 
            // ClearDataItem
            // 
            this.ClearDataItem.Name = "ClearDataItem";
            resources.ApplyResources(this.ClearDataItem, "ClearDataItem");
            this.ClearDataItem.Click += new System.EventHandler(this.ClearDataItem_Click);
            // 
            // ClearAllItem
            // 
            this.ClearAllItem.BackColor = System.Drawing.Color.SandyBrown;
            this.ClearAllItem.Name = "ClearAllItem";
            resources.ApplyResources(this.ClearAllItem, "ClearAllItem");
            this.ClearAllItem.Click += new System.EventHandler(this.ClearAllItem_Click);
            // 
            // TestCryptionItem
            // 
            this.TestCryptionItem.Name = "TestCryptionItem";
            resources.ApplyResources(this.TestCryptionItem, "TestCryptionItem");
            this.TestCryptionItem.Click += new System.EventHandler(this.TestCryptionItem_Click);
            // 
            // QuickLoadDataItem
            // 
            this.QuickLoadDataItem.Name = "QuickLoadDataItem";
            resources.ApplyResources(this.QuickLoadDataItem, "QuickLoadDataItem");
            this.QuickLoadDataItem.Click += new System.EventHandler(this.GetDataItem_Click);
            // 
            // QuickFilterItem
            // 
            this.QuickFilterItem.Name = "QuickFilterItem";
            resources.ApplyResources(this.QuickFilterItem, "QuickFilterItem");
            this.QuickFilterItem.Click += new System.EventHandler(this.FilterItem_Click);
            // 
            // UpdateControllingItem
            // 
            this.UpdateControllingItem.Name = "UpdateControllingItem";
            resources.ApplyResources(this.UpdateControllingItem, "UpdateControllingItem");
            this.UpdateControllingItem.Click += new System.EventHandler(this.UpdateControllingItem_Click);
            // 
            // HelpMenuItem
            // 
            this.HelpMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.HelpAboutItem,
            this.SetupItem,
            this.HelpSystemItem});
            this.HelpMenuItem.Name = "HelpMenuItem";
            resources.ApplyResources(this.HelpMenuItem, "HelpMenuItem");
            // 
            // HelpAboutItem
            // 
            this.HelpAboutItem.Name = "HelpAboutItem";
            resources.ApplyResources(this.HelpAboutItem, "HelpAboutItem");
            this.HelpAboutItem.Click += new System.EventHandler(this.AboutSoft);
            // 
            // SetupItem
            // 
            this.SetupItem.Name = "SetupItem";
            resources.ApplyResources(this.SetupItem, "SetupItem");
            this.SetupItem.Click += new System.EventHandler(this.SetupItem_Click);
            // 
            // HelpSystemItem
            // 
            this.HelpSystemItem.Name = "HelpSystemItem";
            resources.ApplyResources(this.HelpSystemItem, "HelpSystemItem");
            this.HelpSystemItem.Click += new System.EventHandler(this.infoItem_Click);
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            resources.ApplyResources(this.notifyIcon1, "notifyIcon1");
            // 
            // timer1
            // 
            this.timer1.Interval = 500;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // labelHour
            // 
            resources.ApplyResources(this.labelHour, "labelHour");
            this.labelHour.Name = "labelHour";
            // 
            // labelMinute
            // 
            resources.ApplyResources(this.labelMinute, "labelMinute");
            this.labelMinute.Name = "labelMinute";
            // 
            // groupBoxTime
            // 
            resources.ApplyResources(this.groupBoxTime, "groupBoxTime");
            this.groupBoxTime.Controls.Add(this.numUpDownHour);
            this.groupBoxTime.Controls.Add(this.labelHour);
            this.groupBoxTime.Controls.Add(this.numUpDownMinute);
            this.groupBoxTime.Controls.Add(this.labelMinute);
            this.groupBoxTime.Name = "groupBoxTime";
            this.groupBoxTime.TabStop = false;
            // 
            // monthCalendar
            // 
            resources.ApplyResources(this.monthCalendar, "monthCalendar");
            this.monthCalendar.MaxSelectionCount = 60;
            this.monthCalendar.Name = "monthCalendar";
            // 
            // labelGroup
            // 
            resources.ApplyResources(this.labelGroup, "labelGroup");
            this.labelGroup.Name = "labelGroup";
            // 
            // panelView
            // 
            resources.ApplyResources(this.panelView, "panelView");
            this.panelView.Controls.Add(this.dataGridView1);
            this.panelView.Name = "panelView";
            this.panelView.SizeChanged += new System.EventHandler(this.panelView_SizeChanged);
            // 
            // groupBoxProperties
            // 
            resources.ApplyResources(this.groupBoxProperties, "groupBoxProperties");
            this.groupBoxProperties.Controls.Add(this.buttonPropertiesCancel);
            this.groupBoxProperties.Controls.Add(this.buttonPropertiesSave);
            this.groupBoxProperties.Name = "groupBoxProperties";
            this.groupBoxProperties.TabStop = false;
            // 
            // buttonPropertiesCancel
            // 
            resources.ApplyResources(this.buttonPropertiesCancel, "buttonPropertiesCancel");
            this.buttonPropertiesCancel.Name = "buttonPropertiesCancel";
            this.buttonPropertiesCancel.UseVisualStyleBackColor = true;
            this.buttonPropertiesCancel.Click += new System.EventHandler(this.buttonPropertiesCancel_Click);
            // 
            // buttonPropertiesSave
            // 
            resources.ApplyResources(this.buttonPropertiesSave, "buttonPropertiesSave");
            this.buttonPropertiesSave.Name = "buttonPropertiesSave";
            this.buttonPropertiesSave.UseVisualStyleBackColor = true;
            this.buttonPropertiesSave.Click += new System.EventHandler(this.buttonPropertiesSave_Click);
            // 
            // FormPersonViewerSCA
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.groupBoxProperties);
            this.Controls.Add(this.panelView);
            this.Controls.Add(this.textBoxFIO);
            this.Controls.Add(this.labelGroupDescription);
            this.Controls.Add(this.textBoxGroupDescription);
            this.Controls.Add(this.textBoxNav);
            this.Controls.Add(this.labelGroup);
            this.Controls.Add(this.textBoxGroup);
            this.Controls.Add(this.monthCalendar);
            this.Controls.Add(this.groupBoxRemoveDays);
            this.Controls.Add(this.groupBoxPeriod);
            this.Controls.Add(this.groupBoxTime);
            this.Controls.Add(this.menuStrip);
            this.Controls.Add(this.comboBoxFio);
            this.MainMenuStrip = this.menuStrip;
            this.Name = "FormPersonViewerSCA";
            this.toolTip1.SetToolTip(this, resources.GetString("$this.ToolTip"));
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownHour)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownMinute)).EndInit();
            this.groupBoxRemoveDays.ResumeLayout(false);
            this.groupBoxRemoveDays.PerformLayout();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.groupBoxPeriod.ResumeLayout(false);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.groupBoxTime.ResumeLayout(false);
            this.groupBoxTime.PerformLayout();
            this.panelView.ResumeLayout(false);
            this.groupBoxProperties.ResumeLayout(false);
            this.groupBoxProperties.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox comboBoxFio;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ToolStripProgressBar ProgressBar1;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel2;
        private System.Windows.Forms.NumericUpDown numUpDownHour;
        private System.Windows.Forms.NumericUpDown numUpDownMinute;
        private System.Windows.Forms.Label labelHour;
        private System.Windows.Forms.Label labelMinute;
        private System.Windows.Forms.DateTimePicker dateTimePickerStart;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label labelInfoEnd;
        private System.Windows.Forms.Label labelInfoStart;
        private System.Windows.Forms.GroupBox groupBoxTime;
        private System.Windows.Forms.GroupBox groupBoxPeriod;
        private System.Windows.Forms.GroupBox groupBoxRemoveDays;
        private System.Windows.Forms.CheckBox checkBoxCelebrate;
        private System.Windows.Forms.CheckBox checkBoxWeekend;
        private System.Windows.Forms.CheckBox checkBoxStartWorkInTime;
        private System.Windows.Forms.MonthCalendar monthCalendar;
        private System.Windows.Forms.CheckBox checkBoxReEnter;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem FunctionMenuItem;
        private System.Windows.Forms.ToolStripMenuItem GetFioItem;
        private System.Windows.Forms.ToolStripMenuItem GroupsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ListGroupItem;
        private System.Windows.Forms.ToolStripMenuItem ViewMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExportIntoExcelItem;
        private System.Windows.Forms.ToolStripMenuItem CreateGroupItem;
        private System.Windows.Forms.ToolStripMenuItem DeleteGroupItem;
        private System.Windows.Forms.TextBox textBoxGroup;
        private System.Windows.Forms.Label labelGroup;
        private System.Windows.Forms.TextBox textBoxNav;
        private System.Windows.Forms.ToolStripMenuItem MembersGroupItem;
        private System.Windows.Forms.Label labelGroupDescription;
        private System.Windows.Forms.TextBox textBoxGroupDescription;
        private System.Windows.Forms.ToolStripMenuItem ExitItem;
        private System.Windows.Forms.ToolStripMenuItem AddPersonToGroupItem;
        private System.Windows.Forms.ToolStripMenuItem HelpMenuItem;
        private System.Windows.Forms.ToolStripMenuItem HelpAboutItem;
        private System.Windows.Forms.ToolStripMenuItem HelpSystemItem;
        private System.Windows.Forms.ToolStripMenuItem AnualDatesMenuItem;
        private System.Windows.Forms.ToolStripMenuItem EnterEditAnualItem;
        private System.Windows.Forms.ToolStripMenuItem AddAnualDateItem;
        private System.Windows.Forms.ToolStripMenuItem DeleteAnualDateItem;
        private System.Windows.Forms.ToolStripMenuItem QuickLoadDataItem;
        private System.Windows.Forms.ToolStripMenuItem QuickFilterItem;
        private System.Windows.Forms.ToolStripMenuItem QuickSettingsItem;
        private System.Windows.Forms.ToolStripMenuItem SettingsEquipmentItem;
        private System.Windows.Forms.ToolStripMenuItem SettingsProgrammItem;
        private System.Windows.Forms.ToolStripMenuItem SettingsOtherItem;
        private System.Windows.Forms.ToolStripMenuItem DeletePersonFromGroupItem;
        private System.Windows.Forms.TextBox textBoxFIO;
        private System.Windows.Forms.ToolStripMenuItem ClearReportItem;
        private System.Windows.Forms.ToolStripMenuItem ClearDataItem;
        private System.Windows.Forms.ToolStripMenuItem ClearAllItem;
        private System.Windows.Forms.ToolStripMenuItem VisualItem;
        private System.Windows.Forms.ToolStripMenuItem ReportsItem;
        private System.Windows.Forms.Panel panelView;
        private System.Windows.Forms.ToolStripMenuItem VisualWorkedTimeItem;
        private System.Windows.Forms.ToolStripMenuItem SelectColorMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ColorRegistrationMenuItem;
        private System.Windows.Forms.ToolStripMenuItem BlueItem;
        private System.Windows.Forms.ToolStripMenuItem RedItem;
        private System.Windows.Forms.ToolStripMenuItem GreenItem;
        private System.Windows.Forms.ToolStripMenuItem YellowItem;
        private System.Windows.Forms.GroupBox groupBoxProperties;
        private System.Windows.Forms.Button buttonPropertiesSave;
        private System.Windows.Forms.ToolStripMenuItem TestCryptionItem;
        private System.Windows.Forms.Button buttonPropertiesCancel;
        private System.Windows.Forms.ToolStripMenuItem PersonOrGroupItem;
        private System.Windows.Forms.ToolStripMenuItem SetupItem;
        private System.Windows.Forms.ToolStripMenuItem UpdateControllingItem;
    }
}

