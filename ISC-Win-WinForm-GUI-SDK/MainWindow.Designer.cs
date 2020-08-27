using System.Windows.Forms;

namespace ISC_Win_WinForm_GUI
{
    partial class MainWindow
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatus_DeviceStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.tabControl_MainFunctions = new System.Windows.Forms.TabControl();
            this.tabPage_Scan = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.checkBox_zoom = new System.Windows.Forms.CheckBox();
            this.checkBox_tooltip = new System.Windows.Forms.CheckBox();
            this.Check_Overlay = new System.Windows.Forms.CheckBox();
            this.Button_Scan = new System.Windows.Forms.Button();
            this.RadioButton_Reference = new System.Windows.Forms.RadioButton();
            this.RadioButton_Intensity = new System.Windows.Forms.RadioButton();
            this.RadioButton_Absorbance = new System.Windows.Forms.RadioButton();
            this.RadioButton_Reflectance = new System.Windows.Forms.RadioButton();
            this.Label_EstimatedScanTime = new System.Windows.Forms.Label();
            this.Label_CurrentConfig = new System.Windows.Forms.Label();
            this.Label_ScanStatus = new System.Windows.Forms.Label();
            this.MyChart = new LiveCharts.WinForms.CartesianChart();
            this.tabScanPage = new System.Windows.Forms.TabControl();
            this.tabPage_ScanSetting = new System.Windows.Forms.TabPage();
            this.Button_ClearAllErrors = new System.Windows.Forms.Button();
            this.GroupBox_ContScan = new System.Windows.Forms.GroupBox();
            this.checkBox_StopOnError = new System.Windows.Forms.CheckBox();
            this.Label_ContScan = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Text_ContDelay = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Text_ContScan = new System.Windows.Forms.TextBox();
            this.GroupBox_GainControl = new System.Windows.Forms.GroupBox();
            this.CheckBox_AutoGain = new System.Windows.Forms.CheckBox();
            this.ComboBox_PGAGain = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.GroupBox_SaveScan = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.TextBox_FileNamePrefix3 = new System.Windows.Forms.TextBox();
            this.TextBox_FileNamePrefix2 = new System.Windows.Forms.TextBox();
            this.CheckBox_SaveRJDX = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveAJDX = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveIJDX = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveOneCSV = new System.Windows.Forms.CheckBox();
            this.TextBox_FileNamePrefix1 = new System.Windows.Forms.TextBox();
            this.CheckBox_FileNamePrefix = new System.Windows.Forms.CheckBox();
            this.TextBox_SaveDirPath = new System.Windows.Forms.TextBox();
            this.Button_SaveDirChange = new System.Windows.Forms.Button();
            this.CheckBox_SaveRCSV = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveACSV = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveICSV = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveDAT = new System.Windows.Forms.CheckBox();
            this.CheckBox_SaveCombCSV = new System.Windows.Forms.CheckBox();
            this.GroupBox_ScanAvg = new System.Windows.Forms.GroupBox();
            this.textBox_ScanAvg = new System.Windows.Forms.TextBox();
            this.label34 = new System.Windows.Forms.Label();
            this.GroupBox_LampControl = new System.Windows.Forms.GroupBox();
            this.CheckBox_LampOn = new System.Windows.Forms.CheckBox();
            this.TextBox_LampStableTime = new System.Windows.Forms.TextBox();
            this.RadioButton_LampStableTime = new System.Windows.Forms.RadioButton();
            this.RadioButton_LampOff = new System.Windows.Forms.RadioButton();
            this.RadioButton_LampOn = new System.Windows.Forms.RadioButton();
            this.GroupBox_RefSelect = new System.Windows.Forms.GroupBox();
            this.label_ref = new System.Windows.Forms.Label();
            this.RadioButton_RefFac = new System.Windows.Forms.RadioButton();
            this.RadioButton_RefPre = new System.Windows.Forms.RadioButton();
            this.RadioButton_RefNew = new System.Windows.Forms.RadioButton();
            this.tabPage_ScanConfig = new System.Windows.Forms.TabPage();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.Button_MoveCfgT2L = new System.Windows.Forms.Button();
            this.Button_MoveCfgL2T = new System.Windows.Forms.Button();
            this.Button_CopyCfgT2L = new System.Windows.Forms.Button();
            this.Button_CopyCfgL2T = new System.Windows.Forms.Button();
            this.ListBox_LocalCfgs = new System.Windows.Forms.ListBox();
            this.label_ActiveConfig = new System.Windows.Forms.Label();
            this.ListBox_TargetCfgs = new System.Windows.Forms.ListBox();
            this.Button_SetActive = new System.Windows.Forms.Button();
            this.label94 = new System.Windows.Forms.Label();
            this.Button_CfgCancel = new System.Windows.Forms.Button();
            this.Button_CfgSave = new System.Windows.Forms.Button();
            this.Button_CfgDelete = new System.Windows.Forms.Button();
            this.Button_CfgEdit = new System.Windows.Forms.Button();
            this.Button_CfgNew = new System.Windows.Forms.Button();
            this.GroupBox_CfgDetails = new System.Windows.Forms.GroupBox();
            this.label_totalPatterns = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label_maxPattern5 = new System.Windows.Forms.Label();
            this.label_maxPattern4 = new System.Windows.Forms.Label();
            this.label_maxPattern3 = new System.Windows.Forms.Label();
            this.label_maxPattern2 = new System.Windows.Forms.Label();
            this.label_maxPattern1 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label_pattern5 = new System.Windows.Forms.Label();
            this.label_pattern4 = new System.Windows.Forms.Label();
            this.label_pattern3 = new System.Windows.Forms.Label();
            this.label_pattern2 = new System.Windows.Forms.Label();
            this.label_pattern1 = new System.Windows.Forms.Label();
            this.comboBox_cfgNumSec = new System.Windows.Forms.ComboBox();
            this.Label_CfgNumPatterns = new System.Windows.Forms.Label();
            this.label93 = new System.Windows.Forms.Label();
            this.label92 = new System.Windows.Forms.Label();
            this.label91 = new System.Windows.Forms.Label();
            this.label90 = new System.Windows.Forms.Label();
            this.label89 = new System.Windows.Forms.Label();
            this.label88 = new System.Windows.Forms.Label();
            this.TextBox_CfgDigRes1 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgDigRes2 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgDigRes3 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgDigRes4 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgDigRes5 = new System.Windows.Forms.TextBox();
            this.ComboBox_CfgExposure1 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgExposure2 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgExposure3 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgExposure4 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgExposure5 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgWidth1 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgWidth2 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgWidth3 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgWidth4 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgWidth5 = new System.Windows.Forms.ComboBox();
            this.TextBox_CfgRangeEnd1 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeEnd2 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeEnd3 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeEnd4 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeEnd5 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeStart1 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeStart2 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeStart3 = new System.Windows.Forms.TextBox();
            this.TextBox_CfgRangeStart4 = new System.Windows.Forms.TextBox();
            this.ComboBox_CfgScanType1 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgScanType2 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgScanType3 = new System.Windows.Forms.ComboBox();
            this.ComboBox_CfgScanType4 = new System.Windows.Forms.ComboBox();
            this.TextBox_CfgRangeStart5 = new System.Windows.Forms.TextBox();
            this.ComboBox_CfgScanType5 = new System.Windows.Forms.ComboBox();
            this.label87 = new System.Windows.Forms.Label();
            this.label86 = new System.Windows.Forms.Label();
            this.label85 = new System.Windows.Forms.Label();
            this.label84 = new System.Windows.Forms.Label();
            this.label83 = new System.Windows.Forms.Label();
            this.label82 = new System.Windows.Forms.Label();
            this.TextBox_CfgAvg = new System.Windows.Forms.TextBox();
            this.label81 = new System.Windows.Forms.Label();
            this.TextBox_CfgName = new System.Windows.Forms.TextBox();
            this.label80 = new System.Windows.Forms.Label();
            this.tabPage_SaveScans = new System.Windows.Forms.TabPage();
            this.panel_Saved_Scan = new System.Windows.Forms.Panel();
            this.button_clear = new System.Windows.Forms.Button();
            this.textBox_filter = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.dataGridView_savescan = new System.Windows.Forms.DataGridView();
            this.label33 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.Label_SavedAvg = new System.Windows.Forms.Label();
            this.label77 = new System.Windows.Forms.Label();
            this.label76 = new System.Windows.Forms.Label();
            this.label75 = new System.Windows.Forms.Label();
            this.label74 = new System.Windows.Forms.Label();
            this.label73 = new System.Windows.Forms.Label();
            this.Label_SavedExposure1 = new System.Windows.Forms.Label();
            this.Label_SavedExposure2 = new System.Windows.Forms.Label();
            this.Label_SavedExposure3 = new System.Windows.Forms.Label();
            this.Label_SavedExposure4 = new System.Windows.Forms.Label();
            this.Label_SavedExposure5 = new System.Windows.Forms.Label();
            this.label67 = new System.Windows.Forms.Label();
            this.Label_SavedDigRes1 = new System.Windows.Forms.Label();
            this.Label_SavedDigRes2 = new System.Windows.Forms.Label();
            this.Label_SavedDigRes3 = new System.Windows.Forms.Label();
            this.Label_SavedDigRes4 = new System.Windows.Forms.Label();
            this.Label_SavedDigRes5 = new System.Windows.Forms.Label();
            this.label61 = new System.Windows.Forms.Label();
            this.Label_SavedWidth1 = new System.Windows.Forms.Label();
            this.Label_SavedWidth2 = new System.Windows.Forms.Label();
            this.Label_SavedWidth3 = new System.Windows.Forms.Label();
            this.Label_SavedWidth4 = new System.Windows.Forms.Label();
            this.Label_SavedWidth5 = new System.Windows.Forms.Label();
            this.label55 = new System.Windows.Forms.Label();
            this.Label_SavedRangeEnd1 = new System.Windows.Forms.Label();
            this.Label_SavedRangeEnd2 = new System.Windows.Forms.Label();
            this.Label_SavedRangeEnd3 = new System.Windows.Forms.Label();
            this.Label_SavedRangeEnd4 = new System.Windows.Forms.Label();
            this.Label_SavedRangeEnd5 = new System.Windows.Forms.Label();
            this.label49 = new System.Windows.Forms.Label();
            this.Label_SavedRangeStart1 = new System.Windows.Forms.Label();
            this.Label_SavedRangeStart2 = new System.Windows.Forms.Label();
            this.Label_SavedRangeStart3 = new System.Windows.Forms.Label();
            this.Label_SavedRangeStart4 = new System.Windows.Forms.Label();
            this.Label_SavedRangeStart5 = new System.Windows.Forms.Label();
            this.label43 = new System.Windows.Forms.Label();
            this.Label_SavedScanType1 = new System.Windows.Forms.Label();
            this.Label_SavedScanType2 = new System.Windows.Forms.Label();
            this.Label_SavedScanType3 = new System.Windows.Forms.Label();
            this.Label_SavedScanType4 = new System.Windows.Forms.Label();
            this.Label_SavedScanType5 = new System.Windows.Forms.Label();
            this.label37 = new System.Windows.Forms.Label();
            this.label36 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.Button_DisplayDirChange = new System.Windows.Forms.Button();
            this.TextBox_SavedFileDirPath = new System.Windows.Forms.TextBox();
            this.label79 = new System.Windows.Forms.Label();
            this.tabPage_Utility = new System.Windows.Forms.TabPage();
            this.GroupBox_BleName = new System.Windows.Forms.GroupBox();
            this.Button_Get_BLE_Display_Name = new System.Windows.Forms.Button();
            this.Button_Set_BLE_Display_Name = new System.Windows.Forms.Button();
            this.Button_Clear_BLE_Display_Name = new System.Windows.Forms.Button();
            this.TextBox_BLE_Display_Name = new System.Windows.Forms.TextBox();
            this.groupBox_Device = new System.Windows.Forms.GroupBox();
            this.button_restore_fac_ref_warning = new System.Windows.Forms.Button();
            this.Button_LockButton = new System.Windows.Forms.Button();
            this.Label_ButtonStatus = new System.Windows.Forms.Label();
            this.Button_UnlockButton = new System.Windows.Forms.Button();
            this.button_DeviceRestoreFacRef = new System.Windows.Forms.Button();
            this.label_RestoreFacRef = new System.Windows.Forms.Label();
            this.button_DeviceUpdateRef = new System.Windows.Forms.Button();
            this.label119 = new System.Windows.Forms.Label();
            this.button_DeviceResetSys = new System.Windows.Forms.Button();
            this.label118 = new System.Windows.Forms.Label();
            this.groupBox_ActivationKey = new System.Windows.Forms.GroupBox();
            this.label_ActivateStatus = new System.Windows.Forms.Label();
            this.Status = new System.Windows.Forms.Label();
            this.button_KeySet = new System.Windows.Forms.Button();
            this.TextBox_Key = new System.Windows.Forms.TextBox();
            this.label117 = new System.Windows.Forms.Label();
            this.groupBox_DevInfo = new System.Windows.Forms.GroupBox();
            this.Label_BleNameValue = new System.Windows.Forms.Label();
            this.Label_Blename = new System.Windows.Forms.Label();
            this.label_DevInfoLampUsageValue = new System.Windows.Forms.Label();
            this.label_DevInfoLampUsage = new System.Windows.Forms.Label();
            this.label_DevInfoUUID = new System.Windows.Forms.Label();
            this.label114 = new System.Windows.Forms.Label();
            this.label_DevInfoManfacSerNum = new System.Windows.Forms.Label();
            this.label_MFC_Seri_Num = new System.Windows.Forms.Label();
            this.label_DevInfoDevSerNum = new System.Windows.Forms.Label();
            this.label104 = new System.Windows.Forms.Label();
            this.label_DevInfoModelName = new System.Windows.Forms.Label();
            this.label106 = new System.Windows.Forms.Label();
            this.label_DevInfoDetectorBoardVer = new System.Windows.Forms.Label();
            this.label108 = new System.Windows.Forms.Label();
            this.label_DevInfoMainBoardVer = new System.Windows.Forms.Label();
            this.label110 = new System.Windows.Forms.Label();
            this.label_DevInfoDLPCVer = new System.Windows.Forms.Label();
            this.label102 = new System.Windows.Forms.Label();
            this.label_DevInfoTivaSWVer = new System.Windows.Forms.Label();
            this.label98 = new System.Windows.Forms.Label();
            this.label_DevInfoGUIVer = new System.Windows.Forms.Label();
            this.label95 = new System.Windows.Forms.Label();
            this.GroupBox_CalibCoeffs = new System.Windows.Forms.GroupBox();
            this.Button_CalWriteCoeffs = new System.Windows.Forms.Button();
            this.Button_CalReadCoeffs = new System.Windows.Forms.Button();
            this.Button_CalRestoreDefaultCoeffs = new System.Windows.Forms.Button();
            this.Button_CalWriteGenCoeffs = new System.Windows.Forms.Button();
            this.CheckBox_CalWriteEnable = new System.Windows.Forms.CheckBox();
            this.TextBox_ShiftVectCoeff2 = new System.Windows.Forms.TextBox();
            this.label28 = new System.Windows.Forms.Label();
            this.TextBox_ShiftVectCoeff1 = new System.Windows.Forms.TextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.TextBox_ShiftVectCoeff0 = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.TextBox_P2WCoeff2 = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.Label_ScanCfgVer = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.TextBox_P2WCoeff1 = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.Label_RefCalVer = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.TextBox_P2WCoeff0 = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.Label_CalCoeffVer = new System.Windows.Forms.Label();
            this.label29 = new System.Windows.Forms.Label();
            this.GroupBox_Sensors = new System.Windows.Forms.GroupBox();
            this.Button_SensorRead = new System.Windows.Forms.Button();
            this.Label_SensorLampCM2Value = new System.Windows.Forms.Label();
            this.Label_SensorLampCM1Value = new System.Windows.Forms.Label();
            this.Label_SensorLampVM2Value = new System.Windows.Forms.Label();
            this.Label_SensorLampVM1Value = new System.Windows.Forms.Label();
            this.Label_SensorTivaTemp = new System.Windows.Forms.Label();
            this.Label_SensorSysTemp = new System.Windows.Forms.Label();
            this.Label_SensorHumidity = new System.Windows.Forms.Label();
            this.Label_SensorBattCapacity = new System.Windows.Forms.Label();
            this.Label_SensorBattStatus = new System.Windows.Forms.Label();
            this.Label_SensorLampCM2 = new System.Windows.Forms.Label();
            this.Label_SensorLampCM1 = new System.Windows.Forms.Label();
            this.Label_SensorLampVM2 = new System.Windows.Forms.Label();
            this.Label_SensorLampVM1 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label_sys_humidity = new System.Windows.Forms.Label();
            this.lb_BattCapTitle = new System.Windows.Forms.Label();
            this.lb_BattChargerStatusTitle = new System.Windows.Forms.Label();
            this.GroupBox_DateTime = new System.Windows.Forms.GroupBox();
            this.Button_DateTimeGet = new System.Windows.Forms.Button();
            this.Button_DateTimeSync = new System.Windows.Forms.Button();
            this.TextBox_DateTime = new System.Windows.Forms.TextBox();
            this.GroupBox_LampUsage = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Button_LampUsageGet = new System.Windows.Forms.Button();
            this.Button_LampUsageSet = new System.Windows.Forms.Button();
            this.TextBox_LampUsage = new System.Windows.Forms.TextBox();
            this.GroupBox_DLPC150FWUpdate = new System.Windows.Forms.GroupBox();
            this.Button_DLPC150FWUpdate = new System.Windows.Forms.Button();
            this.ProgressBar_DLPC150FWUpdateStatus = new System.Windows.Forms.ProgressBar();
            this.Button_DLPC150FWBrowse = new System.Windows.Forms.Button();
            this.TextBox_DLPC150FWPath = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.GroupBox_TivaFWUpdate = new System.Windows.Forms.GroupBox();
            this.Button_TivaFWUpdate = new System.Windows.Forms.Button();
            this.ProgressBar_TivaFWUpdateStatus = new System.Windows.Forms.ProgressBar();
            this.Button_TivaFWBrowse = new System.Windows.Forms.Button();
            this.TextBox_TivaFWPath = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.GroupBox_SerialNumber = new System.Windows.Forms.GroupBox();
            this.Button_SerialNumberGet = new System.Windows.Forms.Button();
            this.Button_SerialNumberSet = new System.Windows.Forms.Button();
            this.TextBox_SerialNumber = new System.Windows.Forms.TextBox();
            this.GroupBox_ModelName = new System.Windows.Forms.GroupBox();
            this.Button_ModelNameGet = new System.Windows.Forms.Button();
            this.Button_ModelNameSet = new System.Windows.Forms.Button();
            this.TextBox_ModelName = new System.Windows.Forms.TextBox();
            this.tabPage_about = new System.Windows.Forms.TabPage();
            this.groupBox_About = new System.Windows.Forms.GroupBox();
            this.lb_GUI_Version = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.button_About = new System.Windows.Forms.Button();
            this.button_AboutLicense = new System.Windows.Forms.Button();
            this.label_about_us = new System.Windows.Forms.Label();
            this.label_license_agree = new System.Windows.Forms.Label();
            this.label_ErrorStatus = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.statusStrip1.SuspendLayout();
            this.tabControl_MainFunctions.SuspendLayout();
            this.tabPage_Scan.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabScanPage.SuspendLayout();
            this.tabPage_ScanSetting.SuspendLayout();
            this.GroupBox_ContScan.SuspendLayout();
            this.GroupBox_GainControl.SuspendLayout();
            this.GroupBox_SaveScan.SuspendLayout();
            this.GroupBox_ScanAvg.SuspendLayout();
            this.GroupBox_LampControl.SuspendLayout();
            this.GroupBox_RefSelect.SuspendLayout();
            this.tabPage_ScanConfig.SuspendLayout();
            this.GroupBox_CfgDetails.SuspendLayout();
            this.tabPage_SaveScans.SuspendLayout();
            this.panel_Saved_Scan.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_savescan)).BeginInit();
            this.tabPage_Utility.SuspendLayout();
            this.GroupBox_BleName.SuspendLayout();
            this.groupBox_Device.SuspendLayout();
            this.groupBox_ActivationKey.SuspendLayout();
            this.groupBox_DevInfo.SuspendLayout();
            this.GroupBox_CalibCoeffs.SuspendLayout();
            this.GroupBox_Sensors.SuspendLayout();
            this.GroupBox_DateTime.SuspendLayout();
            this.GroupBox_LampUsage.SuspendLayout();
            this.GroupBox_DLPC150FWUpdate.SuspendLayout();
            this.GroupBox_TivaFWUpdate.SuspendLayout();
            this.GroupBox_SerialNumber.SuspendLayout();
            this.GroupBox_ModelName.SuspendLayout();
            this.tabPage_about.SuspendLayout();
            this.groupBox_About.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatus_DeviceStatus});
            this.statusStrip1.Location = new System.Drawing.Point(0, 660);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1264, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatus_DeviceStatus
            // 
            this.toolStripStatus_DeviceStatus.Image = ((System.Drawing.Image)(resources.GetObject("toolStripStatus_DeviceStatus.Image")));
            this.toolStripStatus_DeviceStatus.Name = "toolStripStatus_DeviceStatus";
            this.toolStripStatus_DeviceStatus.Size = new System.Drawing.Size(130, 17);
            this.toolStripStatus_DeviceStatus.Text = "Device Disconnect!";
            // 
            // tabControl_MainFunctions
            // 
            this.tabControl_MainFunctions.Controls.Add(this.tabPage_Scan);
            this.tabControl_MainFunctions.Controls.Add(this.tabPage_Utility);
            this.tabControl_MainFunctions.Controls.Add(this.tabPage_about);
            this.tabControl_MainFunctions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl_MainFunctions.Location = new System.Drawing.Point(0, 0);
            this.tabControl_MainFunctions.Name = "tabControl_MainFunctions";
            this.tabControl_MainFunctions.SelectedIndex = 0;
            this.tabControl_MainFunctions.Size = new System.Drawing.Size(1264, 660);
            this.tabControl_MainFunctions.TabIndex = 2;
            this.tabControl_MainFunctions.SelectedIndexChanged += new System.EventHandler(this.tabControl_MainFunctions_SelectedIndexChanged);
            // 
            // tabPage_Scan
            // 
            this.tabPage_Scan.Controls.Add(this.splitContainer1);
            this.tabPage_Scan.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage_Scan.Location = new System.Drawing.Point(4, 23);
            this.tabPage_Scan.Name = "tabPage_Scan";
            this.tabPage_Scan.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Scan.Size = new System.Drawing.Size(1256, 633);
            this.tabPage_Scan.TabIndex = 0;
            this.tabPage_Scan.Text = "Scan";
            this.tabPage_Scan.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.checkBox_zoom);
            this.splitContainer1.Panel1.Controls.Add(this.checkBox_tooltip);
            this.splitContainer1.Panel1.Controls.Add(this.Check_Overlay);
            this.splitContainer1.Panel1.Controls.Add(this.Button_Scan);
            this.splitContainer1.Panel1.Controls.Add(this.RadioButton_Reference);
            this.splitContainer1.Panel1.Controls.Add(this.RadioButton_Intensity);
            this.splitContainer1.Panel1.Controls.Add(this.RadioButton_Absorbance);
            this.splitContainer1.Panel1.Controls.Add(this.RadioButton_Reflectance);
            this.splitContainer1.Panel1.Controls.Add(this.Label_EstimatedScanTime);
            this.splitContainer1.Panel1.Controls.Add(this.Label_CurrentConfig);
            this.splitContainer1.Panel1.Controls.Add(this.Label_ScanStatus);
            this.splitContainer1.Panel1.Controls.Add(this.MyChart);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tabScanPage);
            this.splitContainer1.Size = new System.Drawing.Size(1250, 627);
            this.splitContainer1.SplitterDistance = 845;
            this.splitContainer1.TabIndex = 0;
            // 
            // checkBox_zoom
            // 
            this.checkBox_zoom.AutoSize = true;
            this.checkBox_zoom.Location = new System.Drawing.Point(546, 597);
            this.checkBox_zoom.Name = "checkBox_zoom";
            this.checkBox_zoom.Size = new System.Drawing.Size(103, 18);
            this.checkBox_zoom.TabIndex = 13;
            this.checkBox_zoom.TabStop = false;
            this.checkBox_zoom.Text = "Zoom and Pan";
            this.checkBox_zoom.UseVisualStyleBackColor = true;
            this.checkBox_zoom.CheckedChanged += new System.EventHandler(this.checkBox_zoom_CheckedChanged);
            // 
            // checkBox_tooltip
            // 
            this.checkBox_tooltip.AutoSize = true;
            this.checkBox_tooltip.Location = new System.Drawing.Point(472, 598);
            this.checkBox_tooltip.Name = "checkBox_tooltip";
            this.checkBox_tooltip.Size = new System.Drawing.Size(66, 18);
            this.checkBox_tooltip.TabIndex = 12;
            this.checkBox_tooltip.TabStop = false;
            this.checkBox_tooltip.Text = "ToolTip";
            this.checkBox_tooltip.UseVisualStyleBackColor = true;
            this.checkBox_tooltip.CheckedChanged += new System.EventHandler(this.checkBox_tooltip_CheckedChanged);
            // 
            // Check_Overlay
            // 
            this.Check_Overlay.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Check_Overlay.AutoSize = true;
            this.Check_Overlay.Location = new System.Drawing.Point(472, 574);
            this.Check_Overlay.Name = "Check_Overlay";
            this.Check_Overlay.Size = new System.Drawing.Size(66, 18);
            this.Check_Overlay.TabIndex = 9;
            this.Check_Overlay.TabStop = false;
            this.Check_Overlay.Text = "Overlay";
            this.Check_Overlay.UseVisualStyleBackColor = true;
            this.Check_Overlay.CheckedChanged += new System.EventHandler(this.Check_Overlay_CheckedChanged);
            // 
            // Button_Scan
            // 
            this.Button_Scan.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Button_Scan.Location = new System.Drawing.Point(674, 569);
            this.Button_Scan.Name = "Button_Scan";
            this.Button_Scan.Size = new System.Drawing.Size(132, 23);
            this.Button_Scan.TabIndex = 8;
            this.Button_Scan.Text = "Scan";
            this.Button_Scan.UseVisualStyleBackColor = true;
            this.Button_Scan.Click += new System.EventHandler(this.Button_Scan_Click);
            // 
            // RadioButton_Reference
            // 
            this.RadioButton_Reference.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioButton_Reference.AutoSize = true;
            this.RadioButton_Reference.Location = new System.Drawing.Point(359, 574);
            this.RadioButton_Reference.Name = "RadioButton_Reference";
            this.RadioButton_Reference.Size = new System.Drawing.Size(80, 18);
            this.RadioButton_Reference.TabIndex = 7;
            this.RadioButton_Reference.Text = "Reference";
            this.RadioButton_Reference.UseVisualStyleBackColor = true;
            this.RadioButton_Reference.CheckedChanged += new System.EventHandler(this.RadioButton_Reference_CheckedChanged);
            // 
            // RadioButton_Intensity
            // 
            this.RadioButton_Intensity.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioButton_Intensity.AutoSize = true;
            this.RadioButton_Intensity.Location = new System.Drawing.Point(256, 573);
            this.RadioButton_Intensity.Name = "RadioButton_Intensity";
            this.RadioButton_Intensity.Size = new System.Drawing.Size(73, 18);
            this.RadioButton_Intensity.TabIndex = 6;
            this.RadioButton_Intensity.Text = "Intensity";
            this.RadioButton_Intensity.UseVisualStyleBackColor = true;
            this.RadioButton_Intensity.CheckedChanged += new System.EventHandler(this.RadioButton_Intensity_CheckedChanged);
            // 
            // RadioButton_Absorbance
            // 
            this.RadioButton_Absorbance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioButton_Absorbance.AutoSize = true;
            this.RadioButton_Absorbance.Location = new System.Drawing.Point(138, 573);
            this.RadioButton_Absorbance.Name = "RadioButton_Absorbance";
            this.RadioButton_Absorbance.Size = new System.Drawing.Size(89, 18);
            this.RadioButton_Absorbance.TabIndex = 5;
            this.RadioButton_Absorbance.Text = "Absorbance";
            this.RadioButton_Absorbance.UseVisualStyleBackColor = true;
            this.RadioButton_Absorbance.CheckedChanged += new System.EventHandler(this.RadioButton_Absorbance_CheckedChanged);
            // 
            // RadioButton_Reflectance
            // 
            this.RadioButton_Reflectance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioButton_Reflectance.AutoSize = true;
            this.RadioButton_Reflectance.Location = new System.Drawing.Point(23, 573);
            this.RadioButton_Reflectance.Name = "RadioButton_Reflectance";
            this.RadioButton_Reflectance.Size = new System.Drawing.Size(87, 18);
            this.RadioButton_Reflectance.TabIndex = 4;
            this.RadioButton_Reflectance.Text = "Reflectance";
            this.RadioButton_Reflectance.UseVisualStyleBackColor = true;
            this.RadioButton_Reflectance.CheckedChanged += new System.EventHandler(this.RadioButton_Reflectance_CheckedChanged);
            // 
            // Label_EstimatedScanTime
            // 
            this.Label_EstimatedScanTime.Location = new System.Drawing.Point(574, 14);
            this.Label_EstimatedScanTime.Name = "Label_EstimatedScanTime";
            this.Label_EstimatedScanTime.Size = new System.Drawing.Size(252, 27);
            this.Label_EstimatedScanTime.TabIndex = 3;
            // 
            // Label_CurrentConfig
            // 
            this.Label_CurrentConfig.Location = new System.Drawing.Point(315, 14);
            this.Label_CurrentConfig.Name = "Label_CurrentConfig";
            this.Label_CurrentConfig.Size = new System.Drawing.Size(235, 27);
            this.Label_CurrentConfig.TabIndex = 2;
            this.toolTip1.SetToolTip(this.Label_CurrentConfig, "Current scan configuration for scan");
            this.Label_CurrentConfig.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Label_CurrentConfig_MouseClick);
            // 
            // Label_ScanStatus
            // 
            this.Label_ScanStatus.Location = new System.Drawing.Point(20, 14);
            this.Label_ScanStatus.Name = "Label_ScanStatus";
            this.Label_ScanStatus.Size = new System.Drawing.Size(265, 27);
            this.Label_ScanStatus.TabIndex = 1;
            // 
            // MyChart
            // 
            this.MyChart.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MyChart.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MyChart.Location = new System.Drawing.Point(23, 68);
            this.MyChart.Name = "MyChart";
            this.MyChart.Size = new System.Drawing.Size(792, 485);
            this.MyChart.TabIndex = 0;
            this.MyChart.Text = "cartesianChart1";
            // 
            // tabScanPage
            // 
            this.tabScanPage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.tabScanPage.Controls.Add(this.tabPage_ScanSetting);
            this.tabScanPage.Controls.Add(this.tabPage_ScanConfig);
            this.tabScanPage.Controls.Add(this.tabPage_SaveScans);
            this.tabScanPage.Location = new System.Drawing.Point(0, 0);
            this.tabScanPage.Name = "tabScanPage";
            this.tabScanPage.SelectedIndex = 0;
            this.tabScanPage.Size = new System.Drawing.Size(401, 627);
            this.tabScanPage.TabIndex = 0;
            this.tabScanPage.SelectedIndexChanged += new System.EventHandler(this.tabScanPage_SelectedIndexChanged);
            // 
            // tabPage_ScanSetting
            // 
            this.tabPage_ScanSetting.Controls.Add(this.Button_ClearAllErrors);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_ContScan);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_GainControl);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_SaveScan);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_ScanAvg);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_LampControl);
            this.tabPage_ScanSetting.Controls.Add(this.GroupBox_RefSelect);
            this.tabPage_ScanSetting.Location = new System.Drawing.Point(4, 23);
            this.tabPage_ScanSetting.Name = "tabPage_ScanSetting";
            this.tabPage_ScanSetting.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_ScanSetting.Size = new System.Drawing.Size(393, 600);
            this.tabPage_ScanSetting.TabIndex = 0;
            this.tabPage_ScanSetting.Text = "Scan Setting";
            this.tabPage_ScanSetting.UseVisualStyleBackColor = true;
            // 
            // Button_ClearAllErrors
            // 
            this.Button_ClearAllErrors.Location = new System.Drawing.Point(271, 571);
            this.Button_ClearAllErrors.Name = "Button_ClearAllErrors";
            this.Button_ClearAllErrors.Size = new System.Drawing.Size(116, 23);
            this.Button_ClearAllErrors.TabIndex = 7;
            this.Button_ClearAllErrors.Text = "Clear All Errors";
            this.Button_ClearAllErrors.UseVisualStyleBackColor = true;
            this.Button_ClearAllErrors.Click += new System.EventHandler(this.Button_ClearAllErrors_Click);
            // 
            // GroupBox_ContScan
            // 
            this.GroupBox_ContScan.Controls.Add(this.checkBox_StopOnError);
            this.GroupBox_ContScan.Controls.Add(this.Label_ContScan);
            this.GroupBox_ContScan.Controls.Add(this.label3);
            this.GroupBox_ContScan.Controls.Add(this.Text_ContDelay);
            this.GroupBox_ContScan.Controls.Add(this.label2);
            this.GroupBox_ContScan.Controls.Add(this.Text_ContScan);
            this.GroupBox_ContScan.Location = new System.Drawing.Point(3, 284);
            this.GroupBox_ContScan.Name = "GroupBox_ContScan";
            this.GroupBox_ContScan.Size = new System.Drawing.Size(381, 83);
            this.GroupBox_ContScan.TabIndex = 5;
            this.GroupBox_ContScan.TabStop = false;
            this.GroupBox_ContScan.Text = "Continuous Scan Select";
            // 
            // checkBox_StopOnError
            // 
            this.checkBox_StopOnError.AutoSize = true;
            this.checkBox_StopOnError.Checked = true;
            this.checkBox_StopOnError.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox_StopOnError.Location = new System.Drawing.Point(20, 55);
            this.checkBox_StopOnError.Name = "checkBox_StopOnError";
            this.checkBox_StopOnError.Size = new System.Drawing.Size(193, 18);
            this.checkBox_StopOnError.TabIndex = 15;
            this.checkBox_StopOnError.TabStop = false;
            this.checkBox_StopOnError.Text = "Stop continuous scans on error";
            this.checkBox_StopOnError.UseVisualStyleBackColor = true;
            // 
            // Label_ContScan
            // 
            this.Label_ContScan.AutoSize = true;
            this.Label_ContScan.Location = new System.Drawing.Point(154, 26);
            this.Label_ContScan.Name = "Label_ContScan";
            this.Label_ContScan.Size = new System.Drawing.Size(32, 14);
            this.Label_ContScan.TabIndex = 14;
            this.Label_ContScan.Text = "(0/0)";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(203, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 14);
            this.label3.TabIndex = 12;
            this.label3.Text = "Scan Delay (s):";
            // 
            // Text_ContDelay
            // 
            this.Text_ContDelay.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.Text_ContDelay.Location = new System.Drawing.Point(295, 23);
            this.Text_ContDelay.Name = "Text_ContDelay";
            this.Text_ContDelay.Size = new System.Drawing.Size(71, 22);
            this.Text_ContDelay.TabIndex = 13;
            this.Text_ContDelay.Text = "0";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 14);
            this.label2.TabIndex = 10;
            this.label2.Text = "Cont. Scan:";
            // 
            // Text_ContScan
            // 
            this.Text_ContScan.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.Text_ContScan.Location = new System.Drawing.Point(88, 23);
            this.Text_ContScan.Name = "Text_ContScan";
            this.Text_ContScan.Size = new System.Drawing.Size(57, 22);
            this.Text_ContScan.TabIndex = 11;
            this.Text_ContScan.Text = "1";
            this.Text_ContScan.TextChanged += new System.EventHandler(this.Text_ContScan_TextChanged);
            this.Text_ContScan.Validated += new System.EventHandler(this.Text_ContScan_Validated);
            // 
            // GroupBox_GainControl
            // 
            this.GroupBox_GainControl.Controls.Add(this.CheckBox_AutoGain);
            this.GroupBox_GainControl.Controls.Add(this.ComboBox_PGAGain);
            this.GroupBox_GainControl.Controls.Add(this.label1);
            this.GroupBox_GainControl.Location = new System.Drawing.Point(3, 161);
            this.GroupBox_GainControl.Name = "GroupBox_GainControl";
            this.GroupBox_GainControl.Size = new System.Drawing.Size(381, 55);
            this.GroupBox_GainControl.TabIndex = 3;
            this.GroupBox_GainControl.TabStop = false;
            this.GroupBox_GainControl.Text = "GainControl";
            // 
            // CheckBox_AutoGain
            // 
            this.CheckBox_AutoGain.AutoSize = true;
            this.CheckBox_AutoGain.Checked = true;
            this.CheckBox_AutoGain.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CheckBox_AutoGain.Location = new System.Drawing.Point(270, 23);
            this.CheckBox_AutoGain.Name = "CheckBox_AutoGain";
            this.CheckBox_AutoGain.Size = new System.Drawing.Size(51, 18);
            this.CheckBox_AutoGain.TabIndex = 2;
            this.CheckBox_AutoGain.TabStop = false;
            this.CheckBox_AutoGain.Text = "Auto";
            this.CheckBox_AutoGain.UseVisualStyleBackColor = true;
            this.CheckBox_AutoGain.CheckedChanged += new System.EventHandler(this.CheckBox_AutoGain_CheckedChanged);
            // 
            // ComboBox_PGAGain
            // 
            this.ComboBox_PGAGain.FormattingEnabled = true;
            this.ComboBox_PGAGain.Items.AddRange(new object[] {
            "1",
            "2",
            "4",
            "8",
            "16",
            "32",
            "64"});
            this.ComboBox_PGAGain.Location = new System.Drawing.Point(158, 21);
            this.ComboBox_PGAGain.Name = "ComboBox_PGAGain";
            this.ComboBox_PGAGain.Size = new System.Drawing.Size(100, 22);
            this.ComboBox_PGAGain.TabIndex = 1;
            this.ComboBox_PGAGain.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "PGA Gain";
            // 
            // GroupBox_SaveScan
            // 
            this.GroupBox_SaveScan.Controls.Add(this.label11);
            this.GroupBox_SaveScan.Controls.Add(this.label10);
            this.GroupBox_SaveScan.Controls.Add(this.TextBox_FileNamePrefix3);
            this.GroupBox_SaveScan.Controls.Add(this.TextBox_FileNamePrefix2);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveRJDX);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveAJDX);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveIJDX);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveOneCSV);
            this.GroupBox_SaveScan.Controls.Add(this.TextBox_FileNamePrefix1);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_FileNamePrefix);
            this.GroupBox_SaveScan.Controls.Add(this.TextBox_SaveDirPath);
            this.GroupBox_SaveScan.Controls.Add(this.Button_SaveDirChange);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveRCSV);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveACSV);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveICSV);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveDAT);
            this.GroupBox_SaveScan.Controls.Add(this.CheckBox_SaveCombCSV);
            this.GroupBox_SaveScan.Location = new System.Drawing.Point(3, 373);
            this.GroupBox_SaveScan.Name = "GroupBox_SaveScan";
            this.GroupBox_SaveScan.Size = new System.Drawing.Size(381, 147);
            this.GroupBox_SaveScan.TabIndex = 6;
            this.GroupBox_SaveScan.TabStop = false;
            this.GroupBox_SaveScan.Text = "Save Scan As";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(273, 122);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(13, 14);
            this.label11.TabIndex = 19;
            this.label11.Text = "_";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(196, 122);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(13, 14);
            this.label10.TabIndex = 18;
            this.label10.Text = "_";
            // 
            // TextBox_FileNamePrefix3
            // 
            this.TextBox_FileNamePrefix3.Location = new System.Drawing.Point(288, 117);
            this.TextBox_FileNamePrefix3.Name = "TextBox_FileNamePrefix3";
            this.TextBox_FileNamePrefix3.Size = new System.Drawing.Size(60, 22);
            this.TextBox_FileNamePrefix3.TabIndex = 14;
            this.TextBox_FileNamePrefix3.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_FileNamePrefix3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // TextBox_FileNamePrefix2
            // 
            this.TextBox_FileNamePrefix2.Location = new System.Drawing.Point(211, 117);
            this.TextBox_FileNamePrefix2.Name = "TextBox_FileNamePrefix2";
            this.TextBox_FileNamePrefix2.Size = new System.Drawing.Size(60, 22);
            this.TextBox_FileNamePrefix2.TabIndex = 13;
            this.TextBox_FileNamePrefix2.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_FileNamePrefix2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // CheckBox_SaveRJDX
            // 
            this.CheckBox_SaveRJDX.AutoSize = true;
            this.CheckBox_SaveRJDX.Location = new System.Drawing.Point(250, 65);
            this.CheckBox_SaveRJDX.Name = "CheckBox_SaveRJDX";
            this.CheckBox_SaveRJDX.Size = new System.Drawing.Size(107, 18);
            this.CheckBox_SaveRJDX.TabIndex = 8;
            this.CheckBox_SaveRJDX.TabStop = false;
            this.CheckBox_SaveRJDX.Text = "-reflectance.jdx";
            this.CheckBox_SaveRJDX.UseVisualStyleBackColor = true;
            this.CheckBox_SaveRJDX.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveAJDX
            // 
            this.CheckBox_SaveAJDX.AutoSize = true;
            this.CheckBox_SaveAJDX.Location = new System.Drawing.Point(250, 43);
            this.CheckBox_SaveAJDX.Name = "CheckBox_SaveAJDX";
            this.CheckBox_SaveAJDX.Size = new System.Drawing.Size(112, 18);
            this.CheckBox_SaveAJDX.TabIndex = 7;
            this.CheckBox_SaveAJDX.TabStop = false;
            this.CheckBox_SaveAJDX.Text = "-absorbance.jdx";
            this.CheckBox_SaveAJDX.UseVisualStyleBackColor = true;
            this.CheckBox_SaveAJDX.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveIJDX
            // 
            this.CheckBox_SaveIJDX.AutoSize = true;
            this.CheckBox_SaveIJDX.Location = new System.Drawing.Point(250, 21);
            this.CheckBox_SaveIJDX.Name = "CheckBox_SaveIJDX";
            this.CheckBox_SaveIJDX.Size = new System.Drawing.Size(95, 18);
            this.CheckBox_SaveIJDX.TabIndex = 6;
            this.CheckBox_SaveIJDX.TabStop = false;
            this.CheckBox_SaveIJDX.Text = "-intensity.jdx";
            this.CheckBox_SaveIJDX.UseVisualStyleBackColor = true;
            this.CheckBox_SaveIJDX.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveOneCSV
            // 
            this.CheckBox_SaveOneCSV.AutoSize = true;
            this.CheckBox_SaveOneCSV.Location = new System.Drawing.Point(20, 67);
            this.CheckBox_SaveOneCSV.Name = "CheckBox_SaveOneCSV";
            this.CheckBox_SaveOneCSV.Size = new System.Drawing.Size(103, 18);
            this.CheckBox_SaveOneCSV.TabIndex = 2;
            this.CheckBox_SaveOneCSV.TabStop = false;
            this.CheckBox_SaveOneCSV.Text = "-combined.csv";
            this.CheckBox_SaveOneCSV.UseVisualStyleBackColor = true;
            this.CheckBox_SaveOneCSV.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // TextBox_FileNamePrefix1
            // 
            this.TextBox_FileNamePrefix1.Location = new System.Drawing.Point(134, 117);
            this.TextBox_FileNamePrefix1.Name = "TextBox_FileNamePrefix1";
            this.TextBox_FileNamePrefix1.Size = new System.Drawing.Size(60, 22);
            this.TextBox_FileNamePrefix1.TabIndex = 12;
            this.TextBox_FileNamePrefix1.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_FileNamePrefix1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // CheckBox_FileNamePrefix
            // 
            this.CheckBox_FileNamePrefix.AutoSize = true;
            this.CheckBox_FileNamePrefix.Location = new System.Drawing.Point(20, 121);
            this.CheckBox_FileNamePrefix.Name = "CheckBox_FileNamePrefix";
            this.CheckBox_FileNamePrefix.Size = new System.Drawing.Size(113, 18);
            this.CheckBox_FileNamePrefix.TabIndex = 11;
            this.CheckBox_FileNamePrefix.TabStop = false;
            this.CheckBox_FileNamePrefix.Text = "File Name Prefix";
            this.CheckBox_FileNamePrefix.UseVisualStyleBackColor = true;
            this.CheckBox_FileNamePrefix.CheckedChanged += new System.EventHandler(this.CheckBox_FileNamePrefix_CheckedChanged);
            // 
            // TextBox_SaveDirPath
            // 
            this.TextBox_SaveDirPath.Location = new System.Drawing.Point(20, 89);
            this.TextBox_SaveDirPath.Name = "TextBox_SaveDirPath";
            this.TextBox_SaveDirPath.ReadOnly = true;
            this.TextBox_SaveDirPath.Size = new System.Drawing.Size(269, 22);
            this.TextBox_SaveDirPath.TabIndex = 9;
            this.TextBox_SaveDirPath.TabStop = false;
            // 
            // Button_SaveDirChange
            // 
            this.Button_SaveDirChange.Location = new System.Drawing.Point(295, 88);
            this.Button_SaveDirChange.Name = "Button_SaveDirChange";
            this.Button_SaveDirChange.Size = new System.Drawing.Size(75, 23);
            this.Button_SaveDirChange.TabIndex = 10;
            this.Button_SaveDirChange.Text = "Directory";
            this.Button_SaveDirChange.UseVisualStyleBackColor = true;
            this.Button_SaveDirChange.Click += new System.EventHandler(this.Button_SaveDirChange_Click);
            // 
            // CheckBox_SaveRCSV
            // 
            this.CheckBox_SaveRCSV.AutoSize = true;
            this.CheckBox_SaveRCSV.Location = new System.Drawing.Point(124, 65);
            this.CheckBox_SaveRCSV.Name = "CheckBox_SaveRCSV";
            this.CheckBox_SaveRCSV.Size = new System.Drawing.Size(108, 18);
            this.CheckBox_SaveRCSV.TabIndex = 5;
            this.CheckBox_SaveRCSV.TabStop = false;
            this.CheckBox_SaveRCSV.Text = "-reflectance.csv";
            this.CheckBox_SaveRCSV.UseVisualStyleBackColor = true;
            this.CheckBox_SaveRCSV.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveACSV
            // 
            this.CheckBox_SaveACSV.AutoSize = true;
            this.CheckBox_SaveACSV.Location = new System.Drawing.Point(124, 43);
            this.CheckBox_SaveACSV.Name = "CheckBox_SaveACSV";
            this.CheckBox_SaveACSV.Size = new System.Drawing.Size(113, 18);
            this.CheckBox_SaveACSV.TabIndex = 4;
            this.CheckBox_SaveACSV.TabStop = false;
            this.CheckBox_SaveACSV.Text = "-absorbance.csv";
            this.CheckBox_SaveACSV.UseVisualStyleBackColor = true;
            this.CheckBox_SaveACSV.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveICSV
            // 
            this.CheckBox_SaveICSV.AutoSize = true;
            this.CheckBox_SaveICSV.Location = new System.Drawing.Point(124, 21);
            this.CheckBox_SaveICSV.Name = "CheckBox_SaveICSV";
            this.CheckBox_SaveICSV.Size = new System.Drawing.Size(96, 18);
            this.CheckBox_SaveICSV.TabIndex = 3;
            this.CheckBox_SaveICSV.TabStop = false;
            this.CheckBox_SaveICSV.Text = "-intensity.csv";
            this.CheckBox_SaveICSV.UseVisualStyleBackColor = true;
            this.CheckBox_SaveICSV.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveDAT
            // 
            this.CheckBox_SaveDAT.AutoSize = true;
            this.CheckBox_SaveDAT.Checked = true;
            this.CheckBox_SaveDAT.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CheckBox_SaveDAT.Location = new System.Drawing.Point(20, 43);
            this.CheckBox_SaveDAT.Name = "CheckBox_SaveDAT";
            this.CheckBox_SaveDAT.Size = new System.Drawing.Size(53, 18);
            this.CheckBox_SaveDAT.TabIndex = 1;
            this.CheckBox_SaveDAT.TabStop = false;
            this.CheckBox_SaveDAT.Text = "*.dat";
            this.CheckBox_SaveDAT.UseVisualStyleBackColor = true;
            this.CheckBox_SaveDAT.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // CheckBox_SaveCombCSV
            // 
            this.CheckBox_SaveCombCSV.AutoSize = true;
            this.CheckBox_SaveCombCSV.Checked = true;
            this.CheckBox_SaveCombCSV.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CheckBox_SaveCombCSV.Location = new System.Drawing.Point(20, 21);
            this.CheckBox_SaveCombCSV.Name = "CheckBox_SaveCombCSV";
            this.CheckBox_SaveCombCSV.Size = new System.Drawing.Size(51, 18);
            this.CheckBox_SaveCombCSV.TabIndex = 0;
            this.CheckBox_SaveCombCSV.TabStop = false;
            this.CheckBox_SaveCombCSV.Text = "*.csv";
            this.CheckBox_SaveCombCSV.UseVisualStyleBackColor = true;
            this.CheckBox_SaveCombCSV.CheckedChanged += new System.EventHandler(this.CheckBox_SaveFileFormat_Click);
            // 
            // GroupBox_ScanAvg
            // 
            this.GroupBox_ScanAvg.Controls.Add(this.textBox_ScanAvg);
            this.GroupBox_ScanAvg.Controls.Add(this.label34);
            this.GroupBox_ScanAvg.Location = new System.Drawing.Point(3, 222);
            this.GroupBox_ScanAvg.Name = "GroupBox_ScanAvg";
            this.GroupBox_ScanAvg.Size = new System.Drawing.Size(381, 56);
            this.GroupBox_ScanAvg.TabIndex = 4;
            this.GroupBox_ScanAvg.TabStop = false;
            this.GroupBox_ScanAvg.Text = "Scan Average";
            // 
            // textBox_ScanAvg
            // 
            this.textBox_ScanAvg.Location = new System.Drawing.Point(159, 26);
            this.textBox_ScanAvg.Name = "textBox_ScanAvg";
            this.textBox_ScanAvg.Size = new System.Drawing.Size(100, 22);
            this.textBox_ScanAvg.TabIndex = 1;
            this.textBox_ScanAvg.Validated += new System.EventHandler(this.textBox_ScanAvg_Validated);
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(18, 29);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(135, 14);
            this.label34.TabIndex = 0;
            this.label34.Text = "Num Scans of Average : ";
            // 
            // GroupBox_LampControl
            // 
            this.GroupBox_LampControl.Controls.Add(this.CheckBox_LampOn);
            this.GroupBox_LampControl.Controls.Add(this.TextBox_LampStableTime);
            this.GroupBox_LampControl.Controls.Add(this.RadioButton_LampStableTime);
            this.GroupBox_LampControl.Controls.Add(this.RadioButton_LampOff);
            this.GroupBox_LampControl.Controls.Add(this.RadioButton_LampOn);
            this.GroupBox_LampControl.Location = new System.Drawing.Point(3, 85);
            this.GroupBox_LampControl.Name = "GroupBox_LampControl";
            this.GroupBox_LampControl.Size = new System.Drawing.Size(381, 70);
            this.GroupBox_LampControl.TabIndex = 2;
            this.GroupBox_LampControl.TabStop = false;
            this.GroupBox_LampControl.Text = "Lamp Control";
            // 
            // CheckBox_LampOn
            // 
            this.CheckBox_LampOn.AutoSize = true;
            this.CheckBox_LampOn.Location = new System.Drawing.Point(20, 22);
            this.CheckBox_LampOn.Name = "CheckBox_LampOn";
            this.CheckBox_LampOn.Size = new System.Drawing.Size(103, 18);
            this.CheckBox_LampOn.TabIndex = 0;
            this.CheckBox_LampOn.TabStop = false;
            this.CheckBox_LampOn.Text = "Keep Lamp On";
            this.CheckBox_LampOn.UseVisualStyleBackColor = true;
            this.CheckBox_LampOn.CheckedChanged += new System.EventHandler(this.CheckBox_LampOn_CheckedChanged);
            // 
            // TextBox_LampStableTime
            // 
            this.TextBox_LampStableTime.Location = new System.Drawing.Point(278, 42);
            this.TextBox_LampStableTime.Name = "TextBox_LampStableTime";
            this.TextBox_LampStableTime.Size = new System.Drawing.Size(92, 22);
            this.TextBox_LampStableTime.TabIndex = 3;
            this.TextBox_LampStableTime.Text = "625";
            this.TextBox_LampStableTime.TextChanged += new System.EventHandler(this.TextBox_LampStableTime_TextChanged);
            // 
            // RadioButton_LampStableTime
            // 
            this.RadioButton_LampStableTime.AutoSize = true;
            this.RadioButton_LampStableTime.Checked = true;
            this.RadioButton_LampStableTime.Location = new System.Drawing.Point(20, 43);
            this.RadioButton_LampStableTime.Name = "RadioButton_LampStableTime";
            this.RadioButton_LampStableTime.Size = new System.Drawing.Size(252, 18);
            this.RadioButton_LampStableTime.TabIndex = 2;
            this.RadioButton_LampStableTime.TabStop = true;
            this.RadioButton_LampStableTime.Text = "Lamp Stable Time  (Unit: ms, Default: 625)";
            this.RadioButton_LampStableTime.UseVisualStyleBackColor = true;
            this.RadioButton_LampStableTime.CheckedChanged += new System.EventHandler(this.RadioButton_LampStableTime_CheckedChanged);
            // 
            // RadioButton_LampOff
            // 
            this.RadioButton_LampOff.AutoSize = true;
            this.RadioButton_LampOff.Location = new System.Drawing.Point(146, 21);
            this.RadioButton_LampOff.Name = "RadioButton_LampOff";
            this.RadioButton_LampOff.Size = new System.Drawing.Size(102, 18);
            this.RadioButton_LampOff.TabIndex = 1;
            this.RadioButton_LampOff.Text = "Keep Lamp Off";
            this.RadioButton_LampOff.UseVisualStyleBackColor = true;
            this.RadioButton_LampOff.CheckedChanged += new System.EventHandler(this.RadioButton_LampOff_CheckedChanged);
            // 
            // RadioButton_LampOn
            // 
            this.RadioButton_LampOn.AutoSize = true;
            this.RadioButton_LampOn.Location = new System.Drawing.Point(20, 21);
            this.RadioButton_LampOn.Name = "RadioButton_LampOn";
            this.RadioButton_LampOn.Size = new System.Drawing.Size(102, 18);
            this.RadioButton_LampOn.TabIndex = 0;
            this.RadioButton_LampOn.Text = "Keep Lamp On";
            this.RadioButton_LampOn.UseVisualStyleBackColor = true;
            this.RadioButton_LampOn.CheckedChanged += new System.EventHandler(this.RadioButton_LampOn_CheckedChanged);
            // 
            // GroupBox_RefSelect
            // 
            this.GroupBox_RefSelect.Controls.Add(this.label_ref);
            this.GroupBox_RefSelect.Controls.Add(this.RadioButton_RefFac);
            this.GroupBox_RefSelect.Controls.Add(this.RadioButton_RefPre);
            this.GroupBox_RefSelect.Controls.Add(this.RadioButton_RefNew);
            this.GroupBox_RefSelect.Location = new System.Drawing.Point(3, 6);
            this.GroupBox_RefSelect.Name = "GroupBox_RefSelect";
            this.GroupBox_RefSelect.Size = new System.Drawing.Size(381, 73);
            this.GroupBox_RefSelect.TabIndex = 1;
            this.GroupBox_RefSelect.TabStop = false;
            this.GroupBox_RefSelect.Text = "Reference Select";
            // 
            // label_ref
            // 
            this.label_ref.AutoSize = true;
            this.label_ref.Location = new System.Drawing.Point(21, 50);
            this.label_ref.Name = "label_ref";
            this.label_ref.Size = new System.Drawing.Size(57, 14);
            this.label_ref.TabIndex = 3;
            this.label_ref.Text = "label_ref";
            // 
            // RadioButton_RefFac
            // 
            this.RadioButton_RefFac.AutoSize = true;
            this.RadioButton_RefFac.Location = new System.Drawing.Point(281, 21);
            this.RadioButton_RefFac.Name = "RadioButton_RefFac";
            this.RadioButton_RefFac.Size = new System.Drawing.Size(66, 18);
            this.RadioButton_RefFac.TabIndex = 2;
            this.RadioButton_RefFac.Text = "Built-in";
            this.RadioButton_RefFac.UseVisualStyleBackColor = true;
            this.RadioButton_RefFac.CheckedChanged += new System.EventHandler(this.RadioButton_RefFac_CheckedChanged);
            // 
            // RadioButton_RefPre
            // 
            this.RadioButton_RefPre.AutoSize = true;
            this.RadioButton_RefPre.Location = new System.Drawing.Point(146, 21);
            this.RadioButton_RefPre.Name = "RadioButton_RefPre";
            this.RadioButton_RefPre.Size = new System.Drawing.Size(71, 18);
            this.RadioButton_RefPre.TabIndex = 1;
            this.RadioButton_RefPre.Text = "Previous";
            this.RadioButton_RefPre.UseVisualStyleBackColor = true;
            this.RadioButton_RefPre.CheckedChanged += new System.EventHandler(this.RadioButton_RefPre_CheckedChanged);
            // 
            // RadioButton_RefNew
            // 
            this.RadioButton_RefNew.AutoSize = true;
            this.RadioButton_RefNew.Checked = true;
            this.RadioButton_RefNew.Location = new System.Drawing.Point(20, 21);
            this.RadioButton_RefNew.Name = "RadioButton_RefNew";
            this.RadioButton_RefNew.Size = new System.Drawing.Size(49, 18);
            this.RadioButton_RefNew.TabIndex = 0;
            this.RadioButton_RefNew.TabStop = true;
            this.RadioButton_RefNew.Text = "New";
            this.RadioButton_RefNew.UseVisualStyleBackColor = true;
            this.RadioButton_RefNew.CheckedChanged += new System.EventHandler(this.RadioButton_RefNew_CheckedChanged);
            // 
            // tabPage_ScanConfig
            // 
            this.tabPage_ScanConfig.Controls.Add(this.label16);
            this.tabPage_ScanConfig.Controls.Add(this.label15);
            this.tabPage_ScanConfig.Controls.Add(this.Button_MoveCfgT2L);
            this.tabPage_ScanConfig.Controls.Add(this.Button_MoveCfgL2T);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CopyCfgT2L);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CopyCfgL2T);
            this.tabPage_ScanConfig.Controls.Add(this.ListBox_LocalCfgs);
            this.tabPage_ScanConfig.Controls.Add(this.label_ActiveConfig);
            this.tabPage_ScanConfig.Controls.Add(this.ListBox_TargetCfgs);
            this.tabPage_ScanConfig.Controls.Add(this.Button_SetActive);
            this.tabPage_ScanConfig.Controls.Add(this.label94);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CfgCancel);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CfgSave);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CfgDelete);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CfgEdit);
            this.tabPage_ScanConfig.Controls.Add(this.Button_CfgNew);
            this.tabPage_ScanConfig.Controls.Add(this.GroupBox_CfgDetails);
            this.tabPage_ScanConfig.Location = new System.Drawing.Point(4, 23);
            this.tabPage_ScanConfig.Name = "tabPage_ScanConfig";
            this.tabPage_ScanConfig.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_ScanConfig.Size = new System.Drawing.Size(393, 600);
            this.tabPage_ScanConfig.TabIndex = 1;
            this.tabPage_ScanConfig.Text = "Scan Config";
            this.tabPage_ScanConfig.UseVisualStyleBackColor = true;
            // 
            // label16
            // 
            this.label16.BackColor = System.Drawing.Color.LightGray;
            this.label16.Location = new System.Drawing.Point(223, 38);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(161, 14);
            this.label16.TabIndex = 47;
            this.label16.Text = "Device configurations";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.Color.LightGray;
            this.label15.Location = new System.Drawing.Point(6, 38);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(161, 14);
            this.label15.TabIndex = 46;
            this.label15.Text = "Local configurations";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Button_MoveCfgT2L
            // 
            this.Button_MoveCfgT2L.Location = new System.Drawing.Point(173, 189);
            this.Button_MoveCfgT2L.Name = "Button_MoveCfgT2L";
            this.Button_MoveCfgT2L.Size = new System.Drawing.Size(44, 37);
            this.Button_MoveCfgT2L.TabIndex = 45;
            this.Button_MoveCfgT2L.Text = "Move <<";
            this.Button_MoveCfgT2L.UseVisualStyleBackColor = true;
            this.Button_MoveCfgT2L.Click += new System.EventHandler(this.Button_MoveCfgT2L_Click);
            // 
            // Button_MoveCfgL2T
            // 
            this.Button_MoveCfgL2T.Location = new System.Drawing.Point(173, 145);
            this.Button_MoveCfgL2T.Name = "Button_MoveCfgL2T";
            this.Button_MoveCfgL2T.Size = new System.Drawing.Size(44, 38);
            this.Button_MoveCfgL2T.TabIndex = 44;
            this.Button_MoveCfgL2T.Text = "Move>>";
            this.Button_MoveCfgL2T.UseVisualStyleBackColor = true;
            this.Button_MoveCfgL2T.Click += new System.EventHandler(this.Button_MoveCfgL2T_Click);
            // 
            // Button_CopyCfgT2L
            // 
            this.Button_CopyCfgT2L.Location = new System.Drawing.Point(173, 101);
            this.Button_CopyCfgT2L.Name = "Button_CopyCfgT2L";
            this.Button_CopyCfgT2L.Size = new System.Drawing.Size(44, 38);
            this.Button_CopyCfgT2L.TabIndex = 43;
            this.Button_CopyCfgT2L.Text = "Copy<<";
            this.Button_CopyCfgT2L.UseVisualStyleBackColor = true;
            this.Button_CopyCfgT2L.Click += new System.EventHandler(this.Button_CopyCfgT2L_Click);
            // 
            // Button_CopyCfgL2T
            // 
            this.Button_CopyCfgL2T.Location = new System.Drawing.Point(173, 55);
            this.Button_CopyCfgL2T.Name = "Button_CopyCfgL2T";
            this.Button_CopyCfgL2T.Size = new System.Drawing.Size(44, 40);
            this.Button_CopyCfgL2T.TabIndex = 42;
            this.Button_CopyCfgL2T.Text = "Copy>>";
            this.Button_CopyCfgL2T.UseVisualStyleBackColor = true;
            this.Button_CopyCfgL2T.Click += new System.EventHandler(this.Button_CopyCfgL2T_Click);
            // 
            // ListBox_LocalCfgs
            // 
            this.ListBox_LocalCfgs.FormattingEnabled = true;
            this.ListBox_LocalCfgs.ItemHeight = 14;
            this.ListBox_LocalCfgs.Location = new System.Drawing.Point(6, 54);
            this.ListBox_LocalCfgs.Name = "ListBox_LocalCfgs";
            this.ListBox_LocalCfgs.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.ListBox_LocalCfgs.Size = new System.Drawing.Size(161, 172);
            this.ListBox_LocalCfgs.TabIndex = 40;
            this.toolTip1.SetToolTip(this.ListBox_LocalCfgs, "Local Config List");
            this.ListBox_LocalCfgs.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ListBox_LocalCfgs_MouseClick);
            this.ListBox_LocalCfgs.SelectedIndexChanged += new System.EventHandler(this.ListBox_LocalCfgs_SelectedIndexChanged);
            this.ListBox_LocalCfgs.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ListBox_LocalCfgs_MouseDoubleClick);
            // 
            // label_ActiveConfig
            // 
            this.label_ActiveConfig.AutoSize = true;
            this.label_ActiveConfig.Font = new System.Drawing.Font("Calibri", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_ActiveConfig.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label_ActiveConfig.Location = new System.Drawing.Point(224, 21);
            this.label_ActiveConfig.Name = "label_ActiveConfig";
            this.label_ActiveConfig.Size = new System.Drawing.Size(53, 14);
            this.label_ActiveConfig.TabIndex = 10;
            this.label_ActiveConfig.Text = "Column 1";
            this.label_ActiveConfig.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ListBox_TargetCfgs
            // 
            this.ListBox_TargetCfgs.FormattingEnabled = true;
            this.ListBox_TargetCfgs.ItemHeight = 14;
            this.ListBox_TargetCfgs.Location = new System.Drawing.Point(223, 54);
            this.ListBox_TargetCfgs.Name = "ListBox_TargetCfgs";
            this.ListBox_TargetCfgs.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.ListBox_TargetCfgs.Size = new System.Drawing.Size(161, 172);
            this.ListBox_TargetCfgs.TabIndex = 41;
            this.toolTip1.SetToolTip(this.ListBox_TargetCfgs, "Device Config List");
            this.ListBox_TargetCfgs.SelectedIndexChanged += new System.EventHandler(this.ListBox_TargetCfgs_SelectedIndexChanged);
            this.ListBox_TargetCfgs.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.ListBox_TargetCfgs_MouseDoubleClick);
            // 
            // Button_SetActive
            // 
            this.Button_SetActive.Location = new System.Drawing.Point(223, 232);
            this.Button_SetActive.Name = "Button_SetActive";
            this.Button_SetActive.Size = new System.Drawing.Size(161, 23);
            this.Button_SetActive.TabIndex = 6;
            this.Button_SetActive.Text = "Set Device Default Config";
            this.Button_SetActive.UseVisualStyleBackColor = true;
            this.Button_SetActive.Click += new System.EventHandler(this.Button_SetActive_Click);
            // 
            // label94
            // 
            this.label94.AutoSize = true;
            this.label94.Font = new System.Drawing.Font("Calibri", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label94.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label94.Location = new System.Drawing.Point(224, 7);
            this.label94.Name = "label94";
            this.label94.Size = new System.Drawing.Size(159, 14);
            this.label94.TabIndex = 7;
            this.label94.Text = "Device Default Configuration : ";
            this.label94.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Button_CfgCancel
            // 
            this.Button_CfgCancel.Location = new System.Drawing.Point(314, 544);
            this.Button_CfgCancel.Name = "Button_CfgCancel";
            this.Button_CfgCancel.Size = new System.Drawing.Size(70, 25);
            this.Button_CfgCancel.TabIndex = 4;
            this.Button_CfgCancel.Text = "Cancel";
            this.Button_CfgCancel.UseVisualStyleBackColor = true;
            this.Button_CfgCancel.Click += new System.EventHandler(this.Button_CfgCancel_Click);
            // 
            // Button_CfgSave
            // 
            this.Button_CfgSave.Location = new System.Drawing.Point(236, 544);
            this.Button_CfgSave.Name = "Button_CfgSave";
            this.Button_CfgSave.Size = new System.Drawing.Size(70, 25);
            this.Button_CfgSave.TabIndex = 3;
            this.Button_CfgSave.Text = "Save";
            this.Button_CfgSave.UseVisualStyleBackColor = true;
            this.Button_CfgSave.Click += new System.EventHandler(this.Button_CfgSave_Click);
            // 
            // Button_CfgDelete
            // 
            this.Button_CfgDelete.Location = new System.Drawing.Point(159, 544);
            this.Button_CfgDelete.Name = "Button_CfgDelete";
            this.Button_CfgDelete.Size = new System.Drawing.Size(70, 25);
            this.Button_CfgDelete.TabIndex = 2;
            this.Button_CfgDelete.Text = "Delete";
            this.Button_CfgDelete.UseVisualStyleBackColor = true;
            this.Button_CfgDelete.Click += new System.EventHandler(this.Button_CfgDelete_Click);
            // 
            // Button_CfgEdit
            // 
            this.Button_CfgEdit.Location = new System.Drawing.Point(82, 544);
            this.Button_CfgEdit.Name = "Button_CfgEdit";
            this.Button_CfgEdit.Size = new System.Drawing.Size(70, 25);
            this.Button_CfgEdit.TabIndex = 1;
            this.Button_CfgEdit.Text = "Edit";
            this.Button_CfgEdit.UseVisualStyleBackColor = true;
            this.Button_CfgEdit.Click += new System.EventHandler(this.Button_CfgEdit_Click);
            // 
            // Button_CfgNew
            // 
            this.Button_CfgNew.Location = new System.Drawing.Point(6, 544);
            this.Button_CfgNew.Name = "Button_CfgNew";
            this.Button_CfgNew.Size = new System.Drawing.Size(70, 25);
            this.Button_CfgNew.TabIndex = 0;
            this.Button_CfgNew.Text = "New";
            this.Button_CfgNew.UseVisualStyleBackColor = true;
            this.Button_CfgNew.Click += new System.EventHandler(this.Button_CfgNew_Click);
            // 
            // GroupBox_CfgDetails
            // 
            this.GroupBox_CfgDetails.Controls.Add(this.label_totalPatterns);
            this.GroupBox_CfgDetails.Controls.Add(this.label7);
            this.GroupBox_CfgDetails.Controls.Add(this.label_maxPattern5);
            this.GroupBox_CfgDetails.Controls.Add(this.label_maxPattern4);
            this.GroupBox_CfgDetails.Controls.Add(this.label_maxPattern3);
            this.GroupBox_CfgDetails.Controls.Add(this.label_maxPattern2);
            this.GroupBox_CfgDetails.Controls.Add(this.label_maxPattern1);
            this.GroupBox_CfgDetails.Controls.Add(this.label14);
            this.GroupBox_CfgDetails.Controls.Add(this.label_pattern5);
            this.GroupBox_CfgDetails.Controls.Add(this.label_pattern4);
            this.GroupBox_CfgDetails.Controls.Add(this.label_pattern3);
            this.GroupBox_CfgDetails.Controls.Add(this.label_pattern2);
            this.GroupBox_CfgDetails.Controls.Add(this.label_pattern1);
            this.GroupBox_CfgDetails.Controls.Add(this.comboBox_cfgNumSec);
            this.GroupBox_CfgDetails.Controls.Add(this.Label_CfgNumPatterns);
            this.GroupBox_CfgDetails.Controls.Add(this.label93);
            this.GroupBox_CfgDetails.Controls.Add(this.label92);
            this.GroupBox_CfgDetails.Controls.Add(this.label91);
            this.GroupBox_CfgDetails.Controls.Add(this.label90);
            this.GroupBox_CfgDetails.Controls.Add(this.label89);
            this.GroupBox_CfgDetails.Controls.Add(this.label88);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgDigRes1);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgDigRes2);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgDigRes3);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgDigRes4);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgDigRes5);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgExposure1);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgExposure2);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgExposure3);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgExposure4);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgExposure5);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgWidth1);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgWidth2);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgWidth3);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgWidth4);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgWidth5);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeEnd1);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeEnd2);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeEnd3);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeEnd4);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeEnd5);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeStart1);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeStart2);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeStart3);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeStart4);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgScanType1);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgScanType2);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgScanType3);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgScanType4);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgRangeStart5);
            this.GroupBox_CfgDetails.Controls.Add(this.ComboBox_CfgScanType5);
            this.GroupBox_CfgDetails.Controls.Add(this.label87);
            this.GroupBox_CfgDetails.Controls.Add(this.label86);
            this.GroupBox_CfgDetails.Controls.Add(this.label85);
            this.GroupBox_CfgDetails.Controls.Add(this.label84);
            this.GroupBox_CfgDetails.Controls.Add(this.label83);
            this.GroupBox_CfgDetails.Controls.Add(this.label82);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgAvg);
            this.GroupBox_CfgDetails.Controls.Add(this.label81);
            this.GroupBox_CfgDetails.Controls.Add(this.TextBox_CfgName);
            this.GroupBox_CfgDetails.Controls.Add(this.label80);
            this.GroupBox_CfgDetails.Location = new System.Drawing.Point(6, 261);
            this.GroupBox_CfgDetails.Name = "GroupBox_CfgDetails";
            this.GroupBox_CfgDetails.Size = new System.Drawing.Size(381, 269);
            this.GroupBox_CfgDetails.TabIndex = 0;
            this.GroupBox_CfgDetails.TabStop = false;
            this.GroupBox_CfgDetails.Text = "Details";
            // 
            // label_totalPatterns
            // 
            this.label_totalPatterns.BackColor = System.Drawing.Color.LightGray;
            this.label_totalPatterns.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_totalPatterns.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_totalPatterns.Location = new System.Drawing.Point(132, 249);
            this.label_totalPatterns.Name = "label_totalPatterns";
            this.label_totalPatterns.Size = new System.Drawing.Size(48, 16);
            this.label_totalPatterns.TabIndex = 66;
            this.label_totalPatterns.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.Color.LightGray;
            this.label7.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label7.Location = new System.Drawing.Point(24, 249);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(101, 16);
            this.label7.TabIndex = 65;
            this.label7.Text = "Total Pattern Used";
            this.toolTip1.SetToolTip(this.label7, "Total DMD patterns used with all configuration sections, the maximum value is 624" +
        ".");
            // 
            // label_maxPattern5
            // 
            this.label_maxPattern5.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_maxPattern5.Location = new System.Drawing.Point(331, 214);
            this.label_maxPattern5.Name = "label_maxPattern5";
            this.label_maxPattern5.Size = new System.Drawing.Size(42, 14);
            this.label_maxPattern5.TabIndex = 64;
            this.label_maxPattern5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_maxPattern4
            // 
            this.label_maxPattern4.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_maxPattern4.Location = new System.Drawing.Point(283, 214);
            this.label_maxPattern4.Name = "label_maxPattern4";
            this.label_maxPattern4.Size = new System.Drawing.Size(42, 14);
            this.label_maxPattern4.TabIndex = 63;
            this.label_maxPattern4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_maxPattern3
            // 
            this.label_maxPattern3.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_maxPattern3.Location = new System.Drawing.Point(234, 214);
            this.label_maxPattern3.Name = "label_maxPattern3";
            this.label_maxPattern3.Size = new System.Drawing.Size(42, 14);
            this.label_maxPattern3.TabIndex = 62;
            this.label_maxPattern3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_maxPattern2
            // 
            this.label_maxPattern2.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_maxPattern2.Location = new System.Drawing.Point(184, 214);
            this.label_maxPattern2.Name = "label_maxPattern2";
            this.label_maxPattern2.Size = new System.Drawing.Size(42, 14);
            this.label_maxPattern2.TabIndex = 61;
            this.label_maxPattern2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_maxPattern1
            // 
            this.label_maxPattern1.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_maxPattern1.Location = new System.Drawing.Point(136, 214);
            this.label_maxPattern1.Name = "label_maxPattern1";
            this.label_maxPattern1.Size = new System.Drawing.Size(42, 14);
            this.label_maxPattern1.TabIndex = 60;
            this.label_maxPattern1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(41, 213);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(84, 14);
            this.label14.TabIndex = 59;
            this.label14.Text = "Max Resolution";
            this.toolTip1.SetToolTip(this.label14, "The maximum resolution is limited by 3x pattern overlap.");
            // 
            // label_pattern5
            // 
            this.label_pattern5.BackColor = System.Drawing.Color.LightGray;
            this.label_pattern5.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_pattern5.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_pattern5.Location = new System.Drawing.Point(328, 231);
            this.label_pattern5.Name = "label_pattern5";
            this.label_pattern5.Size = new System.Drawing.Size(48, 16);
            this.label_pattern5.TabIndex = 58;
            this.label_pattern5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_pattern4
            // 
            this.label_pattern4.BackColor = System.Drawing.Color.LightGray;
            this.label_pattern4.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_pattern4.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_pattern4.Location = new System.Drawing.Point(279, 231);
            this.label_pattern4.Name = "label_pattern4";
            this.label_pattern4.Size = new System.Drawing.Size(48, 16);
            this.label_pattern4.TabIndex = 57;
            this.label_pattern4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_pattern3
            // 
            this.label_pattern3.BackColor = System.Drawing.Color.LightGray;
            this.label_pattern3.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_pattern3.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_pattern3.Location = new System.Drawing.Point(230, 231);
            this.label_pattern3.Name = "label_pattern3";
            this.label_pattern3.Size = new System.Drawing.Size(48, 16);
            this.label_pattern3.TabIndex = 56;
            this.label_pattern3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_pattern2
            // 
            this.label_pattern2.BackColor = System.Drawing.Color.LightGray;
            this.label_pattern2.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_pattern2.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_pattern2.Location = new System.Drawing.Point(181, 231);
            this.label_pattern2.Name = "label_pattern2";
            this.label_pattern2.Size = new System.Drawing.Size(48, 16);
            this.label_pattern2.TabIndex = 55;
            this.label_pattern2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_pattern1
            // 
            this.label_pattern1.BackColor = System.Drawing.Color.LightGray;
            this.label_pattern1.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_pattern1.ForeColor = System.Drawing.Color.MidnightBlue;
            this.label_pattern1.Location = new System.Drawing.Point(132, 231);
            this.label_pattern1.Name = "label_pattern1";
            this.label_pattern1.Size = new System.Drawing.Size(48, 16);
            this.label_pattern1.TabIndex = 54;
            this.label_pattern1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // comboBox_cfgNumSec
            // 
            this.comboBox_cfgNumSec.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_cfgNumSec.FormattingEnabled = true;
            this.comboBox_cfgNumSec.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_cfgNumSec.Location = new System.Drawing.Point(92, 49);
            this.comboBox_cfgNumSec.Name = "comboBox_cfgNumSec";
            this.comboBox_cfgNumSec.Size = new System.Drawing.Size(42, 22);
            this.comboBox_cfgNumSec.TabIndex = 9;
            this.comboBox_cfgNumSec.SelectedIndexChanged += new System.EventHandler(this.CfgSection_SelectionChanged);
            this.comboBox_cfgNumSec.Enter += new System.EventHandler(this.CfgField_Enter);
            this.comboBox_cfgNumSec.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // Label_CfgNumPatterns
            // 
            this.Label_CfgNumPatterns.BackColor = System.Drawing.Color.LightGray;
            this.Label_CfgNumPatterns.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label_CfgNumPatterns.ForeColor = System.Drawing.Color.MidnightBlue;
            this.Label_CfgNumPatterns.Location = new System.Drawing.Point(52, 231);
            this.Label_CfgNumPatterns.Name = "Label_CfgNumPatterns";
            this.Label_CfgNumPatterns.Size = new System.Drawing.Size(73, 16);
            this.Label_CfgNumPatterns.TabIndex = 52;
            this.Label_CfgNumPatterns.Text = "Pattern Used";
            this.toolTip1.SetToolTip(this.Label_CfgNumPatterns, "DMD pattern number used by configuration.");
            // 
            // label93
            // 
            this.label93.AutoSize = true;
            this.label93.Location = new System.Drawing.Point(36, 190);
            this.label93.Name = "label93";
            this.label93.Size = new System.Drawing.Size(90, 14);
            this.label93.TabIndex = 46;
            this.label93.Text = "Dig. Resolution";
            // 
            // label92
            // 
            this.label92.AutoSize = true;
            this.label92.Location = new System.Drawing.Point(15, 168);
            this.label92.Name = "label92";
            this.label92.Size = new System.Drawing.Size(113, 14);
            this.label92.TabIndex = 45;
            this.label92.Text = "Exposure Time (ms)";
            // 
            // label91
            // 
            this.label91.AutoSize = true;
            this.label91.Location = new System.Drawing.Point(60, 146);
            this.label91.Name = "label91";
            this.label91.Size = new System.Drawing.Size(68, 14);
            this.label91.TabIndex = 44;
            this.label91.Text = "Width (nm)";
            // 
            // label90
            // 
            this.label90.AutoSize = true;
            this.label90.Location = new System.Drawing.Point(17, 124);
            this.label90.Name = "label90";
            this.label90.Size = new System.Drawing.Size(111, 14);
            this.label90.TabIndex = 43;
            this.label90.Text = "Spectral Range End";
            // 
            // label89
            // 
            this.label89.AutoSize = true;
            this.label89.Location = new System.Drawing.Point(12, 102);
            this.label89.Name = "label89";
            this.label89.Size = new System.Drawing.Size(116, 14);
            this.label89.TabIndex = 42;
            this.label89.Text = "Spectral Range Start";
            // 
            // label88
            // 
            this.label88.AutoSize = true;
            this.label88.Location = new System.Drawing.Point(69, 80);
            this.label88.Name = "label88";
            this.label88.Size = new System.Drawing.Size(59, 14);
            this.label88.TabIndex = 41;
            this.label88.Text = "Scan Type";
            // 
            // TextBox_CfgDigRes1
            // 
            this.TextBox_CfgDigRes1.Location = new System.Drawing.Point(131, 187);
            this.TextBox_CfgDigRes1.Name = "TextBox_CfgDigRes1";
            this.TextBox_CfgDigRes1.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgDigRes1.TabIndex = 15;
            this.TextBox_CfgDigRes1.Text = "3";
            this.TextBox_CfgDigRes1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgDigRes1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgDigRes1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgDigRes1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgDigRes1.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgDigRes2
            // 
            this.TextBox_CfgDigRes2.Location = new System.Drawing.Point(180, 187);
            this.TextBox_CfgDigRes2.Name = "TextBox_CfgDigRes2";
            this.TextBox_CfgDigRes2.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgDigRes2.TabIndex = 21;
            this.TextBox_CfgDigRes2.Text = "3";
            this.TextBox_CfgDigRes2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgDigRes2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgDigRes2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgDigRes2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgDigRes2.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgDigRes3
            // 
            this.TextBox_CfgDigRes3.Location = new System.Drawing.Point(229, 187);
            this.TextBox_CfgDigRes3.Name = "TextBox_CfgDigRes3";
            this.TextBox_CfgDigRes3.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgDigRes3.TabIndex = 27;
            this.TextBox_CfgDigRes3.Text = "3";
            this.TextBox_CfgDigRes3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgDigRes3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgDigRes3.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgDigRes3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgDigRes3.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgDigRes4
            // 
            this.TextBox_CfgDigRes4.Location = new System.Drawing.Point(278, 187);
            this.TextBox_CfgDigRes4.Name = "TextBox_CfgDigRes4";
            this.TextBox_CfgDigRes4.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgDigRes4.TabIndex = 33;
            this.TextBox_CfgDigRes4.Text = "3";
            this.TextBox_CfgDigRes4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgDigRes4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgDigRes4.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgDigRes4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgDigRes4.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgDigRes5
            // 
            this.TextBox_CfgDigRes5.Location = new System.Drawing.Point(327, 187);
            this.TextBox_CfgDigRes5.Name = "TextBox_CfgDigRes5";
            this.TextBox_CfgDigRes5.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgDigRes5.TabIndex = 39;
            this.TextBox_CfgDigRes5.Text = "3";
            this.TextBox_CfgDigRes5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgDigRes5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgDigRes5.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgDigRes5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgDigRes5.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // ComboBox_CfgExposure1
            // 
            this.ComboBox_CfgExposure1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgExposure1.FormattingEnabled = true;
            this.ComboBox_CfgExposure1.Location = new System.Drawing.Point(131, 165);
            this.ComboBox_CfgExposure1.Name = "ComboBox_CfgExposure1";
            this.ComboBox_CfgExposure1.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgExposure1.TabIndex = 14;
            this.ComboBox_CfgExposure1.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgExposure1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgExposure1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgExposure2
            // 
            this.ComboBox_CfgExposure2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgExposure2.FormattingEnabled = true;
            this.ComboBox_CfgExposure2.Location = new System.Drawing.Point(180, 165);
            this.ComboBox_CfgExposure2.Name = "ComboBox_CfgExposure2";
            this.ComboBox_CfgExposure2.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgExposure2.TabIndex = 20;
            this.ComboBox_CfgExposure2.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgExposure2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgExposure2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgExposure3
            // 
            this.ComboBox_CfgExposure3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgExposure3.FormattingEnabled = true;
            this.ComboBox_CfgExposure3.Location = new System.Drawing.Point(229, 165);
            this.ComboBox_CfgExposure3.Name = "ComboBox_CfgExposure3";
            this.ComboBox_CfgExposure3.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgExposure3.TabIndex = 26;
            this.ComboBox_CfgExposure3.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgExposure3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgExposure3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgExposure4
            // 
            this.ComboBox_CfgExposure4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgExposure4.FormattingEnabled = true;
            this.ComboBox_CfgExposure4.Location = new System.Drawing.Point(278, 165);
            this.ComboBox_CfgExposure4.Name = "ComboBox_CfgExposure4";
            this.ComboBox_CfgExposure4.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgExposure4.TabIndex = 32;
            this.ComboBox_CfgExposure4.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgExposure4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgExposure4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgExposure5
            // 
            this.ComboBox_CfgExposure5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgExposure5.FormattingEnabled = true;
            this.ComboBox_CfgExposure5.Location = new System.Drawing.Point(327, 165);
            this.ComboBox_CfgExposure5.Name = "ComboBox_CfgExposure5";
            this.ComboBox_CfgExposure5.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgExposure5.TabIndex = 38;
            this.ComboBox_CfgExposure5.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgExposure5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgExposure5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgWidth1
            // 
            this.ComboBox_CfgWidth1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgWidth1.FormattingEnabled = true;
            this.ComboBox_CfgWidth1.Location = new System.Drawing.Point(131, 143);
            this.ComboBox_CfgWidth1.Name = "ComboBox_CfgWidth1";
            this.ComboBox_CfgWidth1.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgWidth1.TabIndex = 13;
            this.ComboBox_CfgWidth1.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgWidth1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgWidth1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgWidth2
            // 
            this.ComboBox_CfgWidth2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgWidth2.FormattingEnabled = true;
            this.ComboBox_CfgWidth2.Location = new System.Drawing.Point(180, 143);
            this.ComboBox_CfgWidth2.Name = "ComboBox_CfgWidth2";
            this.ComboBox_CfgWidth2.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgWidth2.TabIndex = 19;
            this.ComboBox_CfgWidth2.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgWidth2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgWidth2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgWidth3
            // 
            this.ComboBox_CfgWidth3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgWidth3.FormattingEnabled = true;
            this.ComboBox_CfgWidth3.Location = new System.Drawing.Point(229, 143);
            this.ComboBox_CfgWidth3.Name = "ComboBox_CfgWidth3";
            this.ComboBox_CfgWidth3.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgWidth3.TabIndex = 25;
            this.ComboBox_CfgWidth3.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgWidth3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgWidth3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgWidth4
            // 
            this.ComboBox_CfgWidth4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgWidth4.FormattingEnabled = true;
            this.ComboBox_CfgWidth4.Location = new System.Drawing.Point(278, 143);
            this.ComboBox_CfgWidth4.Name = "ComboBox_CfgWidth4";
            this.ComboBox_CfgWidth4.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgWidth4.TabIndex = 31;
            this.ComboBox_CfgWidth4.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgWidth4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgWidth4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgWidth5
            // 
            this.ComboBox_CfgWidth5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgWidth5.FormattingEnabled = true;
            this.ComboBox_CfgWidth5.Location = new System.Drawing.Point(327, 143);
            this.ComboBox_CfgWidth5.Name = "ComboBox_CfgWidth5";
            this.ComboBox_CfgWidth5.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgWidth5.TabIndex = 37;
            this.ComboBox_CfgWidth5.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgWidth5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgWidth5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // TextBox_CfgRangeEnd1
            // 
            this.TextBox_CfgRangeEnd1.Location = new System.Drawing.Point(131, 121);
            this.TextBox_CfgRangeEnd1.Name = "TextBox_CfgRangeEnd1";
            this.TextBox_CfgRangeEnd1.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeEnd1.TabIndex = 12;
            this.TextBox_CfgRangeEnd1.Text = "1700";
            this.TextBox_CfgRangeEnd1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeEnd1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeEnd1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeEnd1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeEnd1.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeEnd2
            // 
            this.TextBox_CfgRangeEnd2.Location = new System.Drawing.Point(180, 121);
            this.TextBox_CfgRangeEnd2.Name = "TextBox_CfgRangeEnd2";
            this.TextBox_CfgRangeEnd2.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeEnd2.TabIndex = 18;
            this.TextBox_CfgRangeEnd2.Text = "1700";
            this.TextBox_CfgRangeEnd2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeEnd2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeEnd2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeEnd2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeEnd2.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeEnd3
            // 
            this.TextBox_CfgRangeEnd3.Location = new System.Drawing.Point(229, 121);
            this.TextBox_CfgRangeEnd3.Name = "TextBox_CfgRangeEnd3";
            this.TextBox_CfgRangeEnd3.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeEnd3.TabIndex = 24;
            this.TextBox_CfgRangeEnd3.Text = "1700";
            this.TextBox_CfgRangeEnd3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeEnd3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeEnd3.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeEnd3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeEnd3.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeEnd4
            // 
            this.TextBox_CfgRangeEnd4.Location = new System.Drawing.Point(278, 121);
            this.TextBox_CfgRangeEnd4.Name = "TextBox_CfgRangeEnd4";
            this.TextBox_CfgRangeEnd4.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeEnd4.TabIndex = 30;
            this.TextBox_CfgRangeEnd4.Text = "1700";
            this.TextBox_CfgRangeEnd4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeEnd4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeEnd4.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeEnd4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeEnd4.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeEnd5
            // 
            this.TextBox_CfgRangeEnd5.Location = new System.Drawing.Point(327, 121);
            this.TextBox_CfgRangeEnd5.Name = "TextBox_CfgRangeEnd5";
            this.TextBox_CfgRangeEnd5.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeEnd5.TabIndex = 36;
            this.TextBox_CfgRangeEnd5.Text = "1700";
            this.TextBox_CfgRangeEnd5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeEnd5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeEnd5.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeEnd5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeEnd5.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeStart1
            // 
            this.TextBox_CfgRangeStart1.Location = new System.Drawing.Point(131, 99);
            this.TextBox_CfgRangeStart1.Name = "TextBox_CfgRangeStart1";
            this.TextBox_CfgRangeStart1.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeStart1.TabIndex = 11;
            this.TextBox_CfgRangeStart1.Text = "900";
            this.TextBox_CfgRangeStart1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeStart1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeStart1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeStart1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeStart1.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeStart2
            // 
            this.TextBox_CfgRangeStart2.Location = new System.Drawing.Point(180, 99);
            this.TextBox_CfgRangeStart2.Name = "TextBox_CfgRangeStart2";
            this.TextBox_CfgRangeStart2.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeStart2.TabIndex = 17;
            this.TextBox_CfgRangeStart2.Text = "900";
            this.TextBox_CfgRangeStart2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeStart2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeStart2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeStart2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeStart2.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeStart3
            // 
            this.TextBox_CfgRangeStart3.Location = new System.Drawing.Point(229, 99);
            this.TextBox_CfgRangeStart3.Name = "TextBox_CfgRangeStart3";
            this.TextBox_CfgRangeStart3.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeStart3.TabIndex = 23;
            this.TextBox_CfgRangeStart3.Text = "900";
            this.TextBox_CfgRangeStart3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeStart3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeStart3.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeStart3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeStart3.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // TextBox_CfgRangeStart4
            // 
            this.TextBox_CfgRangeStart4.Location = new System.Drawing.Point(278, 99);
            this.TextBox_CfgRangeStart4.Name = "TextBox_CfgRangeStart4";
            this.TextBox_CfgRangeStart4.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeStart4.TabIndex = 29;
            this.TextBox_CfgRangeStart4.Text = "900";
            this.TextBox_CfgRangeStart4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeStart4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeStart4.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeStart4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeStart4.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // ComboBox_CfgScanType1
            // 
            this.ComboBox_CfgScanType1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgScanType1.FormattingEnabled = true;
            this.ComboBox_CfgScanType1.Location = new System.Drawing.Point(131, 77);
            this.ComboBox_CfgScanType1.Name = "ComboBox_CfgScanType1";
            this.ComboBox_CfgScanType1.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgScanType1.TabIndex = 10;
            this.ComboBox_CfgScanType1.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgScanType1.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgScanType1.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgScanType2
            // 
            this.ComboBox_CfgScanType2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgScanType2.FormattingEnabled = true;
            this.ComboBox_CfgScanType2.Location = new System.Drawing.Point(180, 77);
            this.ComboBox_CfgScanType2.Name = "ComboBox_CfgScanType2";
            this.ComboBox_CfgScanType2.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgScanType2.TabIndex = 16;
            this.ComboBox_CfgScanType2.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgScanType2.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgScanType2.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgScanType3
            // 
            this.ComboBox_CfgScanType3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgScanType3.FormattingEnabled = true;
            this.ComboBox_CfgScanType3.Location = new System.Drawing.Point(229, 77);
            this.ComboBox_CfgScanType3.Name = "ComboBox_CfgScanType3";
            this.ComboBox_CfgScanType3.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgScanType3.TabIndex = 22;
            this.ComboBox_CfgScanType3.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgScanType3.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgScanType3.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // ComboBox_CfgScanType4
            // 
            this.ComboBox_CfgScanType4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgScanType4.FormattingEnabled = true;
            this.ComboBox_CfgScanType4.Location = new System.Drawing.Point(278, 77);
            this.ComboBox_CfgScanType4.Name = "ComboBox_CfgScanType4";
            this.ComboBox_CfgScanType4.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgScanType4.TabIndex = 28;
            this.ComboBox_CfgScanType4.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgScanType4.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgScanType4.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // TextBox_CfgRangeStart5
            // 
            this.TextBox_CfgRangeStart5.Location = new System.Drawing.Point(327, 99);
            this.TextBox_CfgRangeStart5.Name = "TextBox_CfgRangeStart5";
            this.TextBox_CfgRangeStart5.Size = new System.Drawing.Size(50, 22);
            this.TextBox_CfgRangeStart5.TabIndex = 35;
            this.TextBox_CfgRangeStart5.Text = "900";
            this.TextBox_CfgRangeStart5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgRangeStart5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgRangeStart5.KeyUp += new System.Windows.Forms.KeyEventHandler(this.CfgDetails_KeyPressed);
            this.TextBox_CfgRangeStart5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgRangeStart5.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // ComboBox_CfgScanType5
            // 
            this.ComboBox_CfgScanType5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ComboBox_CfgScanType5.FormattingEnabled = true;
            this.ComboBox_CfgScanType5.Location = new System.Drawing.Point(327, 77);
            this.ComboBox_CfgScanType5.Name = "ComboBox_CfgScanType5";
            this.ComboBox_CfgScanType5.Size = new System.Drawing.Size(50, 22);
            this.ComboBox_CfgScanType5.TabIndex = 34;
            this.ComboBox_CfgScanType5.SelectedIndexChanged += new System.EventHandler(this.CfgDetails_Validated);
            this.ComboBox_CfgScanType5.Enter += new System.EventHandler(this.CfgField_Enter);
            this.ComboBox_CfgScanType5.Leave += new System.EventHandler(this.CfgField_LostFocus);
            // 
            // label87
            // 
            this.label87.AutoSize = true;
            this.label87.Location = new System.Drawing.Point(6, 52);
            this.label87.Name = "label87";
            this.label87.Size = new System.Drawing.Size(80, 14);
            this.label87.TabIndex = 9;
            this.label87.Text = "Num Sections";
            // 
            // label86
            // 
            this.label86.Location = new System.Drawing.Point(137, 52);
            this.label86.Name = "label86";
            this.label86.Size = new System.Drawing.Size(39, 12);
            this.label86.TabIndex = 8;
            this.label86.Text = "1";
            this.label86.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label85
            // 
            this.label85.Location = new System.Drawing.Point(183, 52);
            this.label85.Name = "label85";
            this.label85.Size = new System.Drawing.Size(39, 12);
            this.label85.TabIndex = 7;
            this.label85.Text = "2";
            this.label85.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label84
            // 
            this.label84.Location = new System.Drawing.Point(232, 52);
            this.label84.Name = "label84";
            this.label84.Size = new System.Drawing.Size(39, 12);
            this.label84.TabIndex = 6;
            this.label84.Text = "3";
            this.label84.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label83
            // 
            this.label83.Location = new System.Drawing.Point(283, 52);
            this.label83.Name = "label83";
            this.label83.Size = new System.Drawing.Size(39, 12);
            this.label83.TabIndex = 5;
            this.label83.Text = "4";
            this.label83.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label82
            // 
            this.label82.Location = new System.Drawing.Point(331, 52);
            this.label82.Name = "label82";
            this.label82.Size = new System.Drawing.Size(39, 12);
            this.label82.TabIndex = 4;
            this.label82.Text = "5";
            this.label82.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // TextBox_CfgAvg
            // 
            this.TextBox_CfgAvg.Location = new System.Drawing.Point(327, 15);
            this.TextBox_CfgAvg.Name = "TextBox_CfgAvg";
            this.TextBox_CfgAvg.Size = new System.Drawing.Size(48, 22);
            this.TextBox_CfgAvg.TabIndex = 8;
            this.TextBox_CfgAvg.Text = "6";
            this.TextBox_CfgAvg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.TextBox_CfgAvg.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgAvg.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgAvg.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // label81
            // 
            this.label81.AutoSize = true;
            this.label81.Location = new System.Drawing.Point(222, 18);
            this.label81.Name = "label81";
            this.label81.Size = new System.Drawing.Size(104, 14);
            this.label81.TabIndex = 2;
            this.label81.Text = "Num Scans to Avg.";
            // 
            // TextBox_CfgName
            // 
            this.TextBox_CfgName.Location = new System.Drawing.Point(51, 15);
            this.TextBox_CfgName.Name = "TextBox_CfgName";
            this.TextBox_CfgName.Size = new System.Drawing.Size(129, 22);
            this.TextBox_CfgName.TabIndex = 7;
            this.TextBox_CfgName.Text = "Config1";
            this.TextBox_CfgName.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_CfgName.Enter += new System.EventHandler(this.CfgField_Enter);
            this.TextBox_CfgName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            this.TextBox_CfgName.Leave += new System.EventHandler(this.CfgField_LostFocus);
            this.TextBox_CfgName.Validated += new System.EventHandler(this.CfgDetails_Validated);
            // 
            // label80
            // 
            this.label80.AutoSize = true;
            this.label80.Location = new System.Drawing.Point(6, 18);
            this.label80.Name = "label80";
            this.label80.Size = new System.Drawing.Size(39, 14);
            this.label80.TabIndex = 0;
            this.label80.Text = "Name";
            // 
            // tabPage_SaveScans
            // 
            this.tabPage_SaveScans.Controls.Add(this.panel_Saved_Scan);
            this.tabPage_SaveScans.Location = new System.Drawing.Point(4, 23);
            this.tabPage_SaveScans.Name = "tabPage_SaveScans";
            this.tabPage_SaveScans.Size = new System.Drawing.Size(393, 600);
            this.tabPage_SaveScans.TabIndex = 2;
            this.tabPage_SaveScans.Text = "Saved Scans";
            this.tabPage_SaveScans.UseVisualStyleBackColor = true;
            // 
            // panel_Saved_Scan
            // 
            this.panel_Saved_Scan.Controls.Add(this.button_clear);
            this.panel_Saved_Scan.Controls.Add(this.textBox_filter);
            this.panel_Saved_Scan.Controls.Add(this.label13);
            this.panel_Saved_Scan.Controls.Add(this.dataGridView_savescan);
            this.panel_Saved_Scan.Controls.Add(this.label33);
            this.panel_Saved_Scan.Controls.Add(this.label30);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedAvg);
            this.panel_Saved_Scan.Controls.Add(this.label77);
            this.panel_Saved_Scan.Controls.Add(this.label76);
            this.panel_Saved_Scan.Controls.Add(this.label75);
            this.panel_Saved_Scan.Controls.Add(this.label74);
            this.panel_Saved_Scan.Controls.Add(this.label73);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedExposure1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedExposure2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedExposure3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedExposure4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedExposure5);
            this.panel_Saved_Scan.Controls.Add(this.label67);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedDigRes1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedDigRes2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedDigRes3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedDigRes4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedDigRes5);
            this.panel_Saved_Scan.Controls.Add(this.label61);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedWidth1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedWidth2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedWidth3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedWidth4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedWidth5);
            this.panel_Saved_Scan.Controls.Add(this.label55);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeEnd1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeEnd2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeEnd3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeEnd4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeEnd5);
            this.panel_Saved_Scan.Controls.Add(this.label49);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeStart1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeStart2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeStart3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeStart4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedRangeStart5);
            this.panel_Saved_Scan.Controls.Add(this.label43);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedScanType1);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedScanType2);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedScanType3);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedScanType4);
            this.panel_Saved_Scan.Controls.Add(this.Label_SavedScanType5);
            this.panel_Saved_Scan.Controls.Add(this.label37);
            this.panel_Saved_Scan.Controls.Add(this.label36);
            this.panel_Saved_Scan.Controls.Add(this.label35);
            this.panel_Saved_Scan.Controls.Add(this.Button_DisplayDirChange);
            this.panel_Saved_Scan.Controls.Add(this.TextBox_SavedFileDirPath);
            this.panel_Saved_Scan.Controls.Add(this.label79);
            this.panel_Saved_Scan.Location = new System.Drawing.Point(1, 1);
            this.panel_Saved_Scan.Name = "panel_Saved_Scan";
            this.panel_Saved_Scan.Size = new System.Drawing.Size(392, 584);
            this.panel_Saved_Scan.TabIndex = 0;
            // 
            // button_clear
            // 
            this.button_clear.Location = new System.Drawing.Point(315, 43);
            this.button_clear.Name = "button_clear";
            this.button_clear.Size = new System.Drawing.Size(75, 23);
            this.button_clear.TabIndex = 5;
            this.button_clear.Text = "Clear";
            this.button_clear.UseVisualStyleBackColor = true;
            this.button_clear.Click += new System.EventHandler(this.button_clear_Click);
            // 
            // textBox_filter
            // 
            this.textBox_filter.Location = new System.Drawing.Point(49, 44);
            this.textBox_filter.Name = "textBox_filter";
            this.textBox_filter.Size = new System.Drawing.Size(260, 22);
            this.textBox_filter.TabIndex = 3;
            this.textBox_filter.TextChanged += new System.EventHandler(this.textBox_filter_TextChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(7, 47);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(45, 14);
            this.label13.TabIndex = 104;
            this.label13.Text = "Filter : ";
            // 
            // dataGridView_savescan
            // 
            this.dataGridView_savescan.AllowUserToAddRows = false;
            this.dataGridView_savescan.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView_savescan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_savescan.Location = new System.Drawing.Point(3, 72);
            this.dataGridView_savescan.Name = "dataGridView_savescan";
            this.dataGridView_savescan.ReadOnly = true;
            this.dataGridView_savescan.RowTemplate.Height = 24;
            this.dataGridView_savescan.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_savescan.Size = new System.Drawing.Size(386, 278);
            this.dataGridView_savescan.TabIndex = 6;
            this.dataGridView_savescan.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView_savescan_KeyDown);
            this.dataGridView_savescan.MouseClick += new System.Windows.Forms.MouseEventHandler(this.dataGridView_savescan_MouseClick);
            // 
            // label33
            // 
            this.label33.Location = new System.Drawing.Point(342, 387);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(39, 12);
            this.label33.TabIndex = 101;
            this.label33.Text = "5";
            this.label33.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label30
            // 
            this.label30.Location = new System.Drawing.Point(288, 387);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(39, 12);
            this.label30.TabIndex = 100;
            this.label30.Text = "4";
            this.label30.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedAvg
            // 
            this.Label_SavedAvg.Location = new System.Drawing.Point(126, 555);
            this.Label_SavedAvg.Name = "Label_SavedAvg";
            this.Label_SavedAvg.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedAvg.TabIndex = 98;
            this.Label_SavedAvg.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label77
            // 
            this.label77.Location = new System.Drawing.Point(180, 555);
            this.label77.Name = "label77";
            this.label77.Size = new System.Drawing.Size(39, 12);
            this.label77.TabIndex = 97;
            this.label77.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label76
            // 
            this.label76.Location = new System.Drawing.Point(234, 554);
            this.label76.Name = "label76";
            this.label76.Size = new System.Drawing.Size(39, 12);
            this.label76.TabIndex = 96;
            this.label76.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label75
            // 
            this.label75.Location = new System.Drawing.Point(288, 554);
            this.label75.Name = "label75";
            this.label75.Size = new System.Drawing.Size(39, 12);
            this.label75.TabIndex = 95;
            this.label75.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label74
            // 
            this.label74.Location = new System.Drawing.Point(342, 554);
            this.label74.Name = "label74";
            this.label74.Size = new System.Drawing.Size(39, 12);
            this.label74.TabIndex = 94;
            this.label74.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label73
            // 
            this.label73.AutoSize = true;
            this.label73.Location = new System.Drawing.Point(6, 530);
            this.label73.Name = "label73";
            this.label73.Size = new System.Drawing.Size(113, 14);
            this.label73.TabIndex = 93;
            this.label73.Text = "Exposure Time (ms)";
            // 
            // Label_SavedExposure1
            // 
            this.Label_SavedExposure1.Location = new System.Drawing.Point(126, 530);
            this.Label_SavedExposure1.Name = "Label_SavedExposure1";
            this.Label_SavedExposure1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedExposure1.TabIndex = 92;
            this.Label_SavedExposure1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedExposure2
            // 
            this.Label_SavedExposure2.Location = new System.Drawing.Point(180, 530);
            this.Label_SavedExposure2.Name = "Label_SavedExposure2";
            this.Label_SavedExposure2.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedExposure2.TabIndex = 91;
            this.Label_SavedExposure2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedExposure3
            // 
            this.Label_SavedExposure3.Location = new System.Drawing.Point(234, 530);
            this.Label_SavedExposure3.Name = "Label_SavedExposure3";
            this.Label_SavedExposure3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedExposure3.TabIndex = 90;
            this.Label_SavedExposure3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedExposure4
            // 
            this.Label_SavedExposure4.Location = new System.Drawing.Point(288, 530);
            this.Label_SavedExposure4.Name = "Label_SavedExposure4";
            this.Label_SavedExposure4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedExposure4.TabIndex = 89;
            this.Label_SavedExposure4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedExposure5
            // 
            this.Label_SavedExposure5.Location = new System.Drawing.Point(342, 530);
            this.Label_SavedExposure5.Name = "Label_SavedExposure5";
            this.Label_SavedExposure5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedExposure5.TabIndex = 88;
            this.Label_SavedExposure5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label67
            // 
            this.label67.AutoSize = true;
            this.label67.Location = new System.Drawing.Point(6, 506);
            this.label67.Name = "label67";
            this.label67.Size = new System.Drawing.Size(106, 14);
            this.label67.TabIndex = 87;
            this.label67.Text = "Digital Resolution";
            // 
            // Label_SavedDigRes1
            // 
            this.Label_SavedDigRes1.Location = new System.Drawing.Point(126, 506);
            this.Label_SavedDigRes1.Name = "Label_SavedDigRes1";
            this.Label_SavedDigRes1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedDigRes1.TabIndex = 86;
            this.Label_SavedDigRes1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedDigRes2
            // 
            this.Label_SavedDigRes2.Location = new System.Drawing.Point(180, 506);
            this.Label_SavedDigRes2.Name = "Label_SavedDigRes2";
            this.Label_SavedDigRes2.Size = new System.Drawing.Size(39, 14);
            this.Label_SavedDigRes2.TabIndex = 85;
            this.Label_SavedDigRes2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedDigRes3
            // 
            this.Label_SavedDigRes3.Location = new System.Drawing.Point(234, 506);
            this.Label_SavedDigRes3.Name = "Label_SavedDigRes3";
            this.Label_SavedDigRes3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedDigRes3.TabIndex = 84;
            this.Label_SavedDigRes3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedDigRes4
            // 
            this.Label_SavedDigRes4.Location = new System.Drawing.Point(288, 506);
            this.Label_SavedDigRes4.Name = "Label_SavedDigRes4";
            this.Label_SavedDigRes4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedDigRes4.TabIndex = 83;
            this.Label_SavedDigRes4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedDigRes5
            // 
            this.Label_SavedDigRes5.Location = new System.Drawing.Point(342, 506);
            this.Label_SavedDigRes5.Name = "Label_SavedDigRes5";
            this.Label_SavedDigRes5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedDigRes5.TabIndex = 82;
            this.Label_SavedDigRes5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label61
            // 
            this.label61.AutoSize = true;
            this.label61.Location = new System.Drawing.Point(6, 482);
            this.label61.Name = "label61";
            this.label61.Size = new System.Drawing.Size(68, 14);
            this.label61.TabIndex = 81;
            this.label61.Text = "Width (nm)";
            // 
            // Label_SavedWidth1
            // 
            this.Label_SavedWidth1.Location = new System.Drawing.Point(126, 482);
            this.Label_SavedWidth1.Name = "Label_SavedWidth1";
            this.Label_SavedWidth1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedWidth1.TabIndex = 80;
            this.Label_SavedWidth1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedWidth2
            // 
            this.Label_SavedWidth2.Location = new System.Drawing.Point(180, 482);
            this.Label_SavedWidth2.Name = "Label_SavedWidth2";
            this.Label_SavedWidth2.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedWidth2.TabIndex = 79;
            this.Label_SavedWidth2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedWidth3
            // 
            this.Label_SavedWidth3.Location = new System.Drawing.Point(234, 482);
            this.Label_SavedWidth3.Name = "Label_SavedWidth3";
            this.Label_SavedWidth3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedWidth3.TabIndex = 78;
            this.Label_SavedWidth3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedWidth4
            // 
            this.Label_SavedWidth4.Location = new System.Drawing.Point(288, 482);
            this.Label_SavedWidth4.Name = "Label_SavedWidth4";
            this.Label_SavedWidth4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedWidth4.TabIndex = 77;
            this.Label_SavedWidth4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedWidth5
            // 
            this.Label_SavedWidth5.Location = new System.Drawing.Point(342, 482);
            this.Label_SavedWidth5.Name = "Label_SavedWidth5";
            this.Label_SavedWidth5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedWidth5.TabIndex = 76;
            this.Label_SavedWidth5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label55
            // 
            this.label55.AutoSize = true;
            this.label55.Location = new System.Drawing.Point(6, 458);
            this.label55.Name = "label55";
            this.label55.Size = new System.Drawing.Size(111, 14);
            this.label55.TabIndex = 75;
            this.label55.Text = "Spectral Range End";
            // 
            // Label_SavedRangeEnd1
            // 
            this.Label_SavedRangeEnd1.Location = new System.Drawing.Point(126, 458);
            this.Label_SavedRangeEnd1.Name = "Label_SavedRangeEnd1";
            this.Label_SavedRangeEnd1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeEnd1.TabIndex = 74;
            this.Label_SavedRangeEnd1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeEnd2
            // 
            this.Label_SavedRangeEnd2.Location = new System.Drawing.Point(180, 458);
            this.Label_SavedRangeEnd2.Name = "Label_SavedRangeEnd2";
            this.Label_SavedRangeEnd2.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeEnd2.TabIndex = 73;
            this.Label_SavedRangeEnd2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeEnd3
            // 
            this.Label_SavedRangeEnd3.Location = new System.Drawing.Point(234, 458);
            this.Label_SavedRangeEnd3.Name = "Label_SavedRangeEnd3";
            this.Label_SavedRangeEnd3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeEnd3.TabIndex = 72;
            this.Label_SavedRangeEnd3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeEnd4
            // 
            this.Label_SavedRangeEnd4.Location = new System.Drawing.Point(288, 458);
            this.Label_SavedRangeEnd4.Name = "Label_SavedRangeEnd4";
            this.Label_SavedRangeEnd4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeEnd4.TabIndex = 71;
            this.Label_SavedRangeEnd4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeEnd5
            // 
            this.Label_SavedRangeEnd5.Location = new System.Drawing.Point(342, 458);
            this.Label_SavedRangeEnd5.Name = "Label_SavedRangeEnd5";
            this.Label_SavedRangeEnd5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeEnd5.TabIndex = 70;
            this.Label_SavedRangeEnd5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label49
            // 
            this.label49.AutoSize = true;
            this.label49.Location = new System.Drawing.Point(6, 434);
            this.label49.Name = "label49";
            this.label49.Size = new System.Drawing.Size(116, 14);
            this.label49.TabIndex = 69;
            this.label49.Text = "Spectral Range Start";
            // 
            // Label_SavedRangeStart1
            // 
            this.Label_SavedRangeStart1.Location = new System.Drawing.Point(126, 434);
            this.Label_SavedRangeStart1.Name = "Label_SavedRangeStart1";
            this.Label_SavedRangeStart1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeStart1.TabIndex = 68;
            this.Label_SavedRangeStart1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeStart2
            // 
            this.Label_SavedRangeStart2.Location = new System.Drawing.Point(180, 434);
            this.Label_SavedRangeStart2.Name = "Label_SavedRangeStart2";
            this.Label_SavedRangeStart2.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeStart2.TabIndex = 67;
            this.Label_SavedRangeStart2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeStart3
            // 
            this.Label_SavedRangeStart3.Location = new System.Drawing.Point(234, 434);
            this.Label_SavedRangeStart3.Name = "Label_SavedRangeStart3";
            this.Label_SavedRangeStart3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeStart3.TabIndex = 66;
            this.Label_SavedRangeStart3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeStart4
            // 
            this.Label_SavedRangeStart4.Location = new System.Drawing.Point(288, 434);
            this.Label_SavedRangeStart4.Name = "Label_SavedRangeStart4";
            this.Label_SavedRangeStart4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeStart4.TabIndex = 65;
            this.Label_SavedRangeStart4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedRangeStart5
            // 
            this.Label_SavedRangeStart5.Location = new System.Drawing.Point(342, 434);
            this.Label_SavedRangeStart5.Name = "Label_SavedRangeStart5";
            this.Label_SavedRangeStart5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedRangeStart5.TabIndex = 64;
            this.Label_SavedRangeStart5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label43
            // 
            this.label43.AutoSize = true;
            this.label43.Location = new System.Drawing.Point(6, 410);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(59, 14);
            this.label43.TabIndex = 63;
            this.label43.Text = "Scan Type";
            // 
            // Label_SavedScanType1
            // 
            this.Label_SavedScanType1.Location = new System.Drawing.Point(126, 410);
            this.Label_SavedScanType1.Name = "Label_SavedScanType1";
            this.Label_SavedScanType1.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedScanType1.TabIndex = 62;
            this.Label_SavedScanType1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedScanType2
            // 
            this.Label_SavedScanType2.Location = new System.Drawing.Point(180, 410);
            this.Label_SavedScanType2.Name = "Label_SavedScanType2";
            this.Label_SavedScanType2.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedScanType2.TabIndex = 61;
            this.Label_SavedScanType2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedScanType3
            // 
            this.Label_SavedScanType3.Location = new System.Drawing.Point(234, 410);
            this.Label_SavedScanType3.Name = "Label_SavedScanType3";
            this.Label_SavedScanType3.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedScanType3.TabIndex = 60;
            this.Label_SavedScanType3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedScanType4
            // 
            this.Label_SavedScanType4.Location = new System.Drawing.Point(288, 410);
            this.Label_SavedScanType4.Name = "Label_SavedScanType4";
            this.Label_SavedScanType4.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedScanType4.TabIndex = 59;
            this.Label_SavedScanType4.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Label_SavedScanType5
            // 
            this.Label_SavedScanType5.Location = new System.Drawing.Point(342, 410);
            this.Label_SavedScanType5.Name = "Label_SavedScanType5";
            this.Label_SavedScanType5.Size = new System.Drawing.Size(39, 12);
            this.Label_SavedScanType5.TabIndex = 58;
            this.Label_SavedScanType5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label37
            // 
            this.label37.Location = new System.Drawing.Point(126, 387);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(39, 12);
            this.label37.TabIndex = 57;
            this.label37.Text = "1";
            this.label37.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label36
            // 
            this.label36.Location = new System.Drawing.Point(180, 387);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(39, 12);
            this.label36.TabIndex = 56;
            this.label36.Text = "2";
            this.label36.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // label35
            // 
            this.label35.Location = new System.Drawing.Point(234, 387);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(39, 12);
            this.label35.TabIndex = 55;
            this.label35.Text = "3";
            this.label35.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Button_DisplayDirChange
            // 
            this.Button_DisplayDirChange.Location = new System.Drawing.Point(315, 16);
            this.Button_DisplayDirChange.Name = "Button_DisplayDirChange";
            this.Button_DisplayDirChange.Size = new System.Drawing.Size(75, 23);
            this.Button_DisplayDirChange.TabIndex = 2;
            this.Button_DisplayDirChange.Text = "Directory";
            this.Button_DisplayDirChange.UseVisualStyleBackColor = true;
            this.Button_DisplayDirChange.Click += new System.EventHandler(this.Button_DisplayDirChange_Click);
            // 
            // TextBox_SavedFileDirPath
            // 
            this.TextBox_SavedFileDirPath.Location = new System.Drawing.Point(6, 16);
            this.TextBox_SavedFileDirPath.Name = "TextBox_SavedFileDirPath";
            this.TextBox_SavedFileDirPath.ReadOnly = true;
            this.TextBox_SavedFileDirPath.Size = new System.Drawing.Size(303, 22);
            this.TextBox_SavedFileDirPath.TabIndex = 1;
            // 
            // label79
            // 
            this.label79.AutoSize = true;
            this.label79.Location = new System.Drawing.Point(6, 555);
            this.label79.Name = "label79";
            this.label79.Size = new System.Drawing.Size(104, 14);
            this.label79.TabIndex = 99;
            this.label79.Text = "Num Scans to Avg.";
            // 
            // tabPage_Utility
            // 
            this.tabPage_Utility.Controls.Add(this.GroupBox_BleName);
            this.tabPage_Utility.Controls.Add(this.groupBox_Device);
            this.tabPage_Utility.Controls.Add(this.groupBox_ActivationKey);
            this.tabPage_Utility.Controls.Add(this.groupBox_DevInfo);
            this.tabPage_Utility.Controls.Add(this.GroupBox_CalibCoeffs);
            this.tabPage_Utility.Controls.Add(this.GroupBox_Sensors);
            this.tabPage_Utility.Controls.Add(this.GroupBox_DateTime);
            this.tabPage_Utility.Controls.Add(this.GroupBox_LampUsage);
            this.tabPage_Utility.Controls.Add(this.GroupBox_DLPC150FWUpdate);
            this.tabPage_Utility.Controls.Add(this.GroupBox_TivaFWUpdate);
            this.tabPage_Utility.Controls.Add(this.GroupBox_SerialNumber);
            this.tabPage_Utility.Controls.Add(this.GroupBox_ModelName);
            this.tabPage_Utility.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPage_Utility.Location = new System.Drawing.Point(4, 23);
            this.tabPage_Utility.Name = "tabPage_Utility";
            this.tabPage_Utility.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage_Utility.Size = new System.Drawing.Size(1256, 633);
            this.tabPage_Utility.TabIndex = 1;
            this.tabPage_Utility.Text = "Utility";
            this.tabPage_Utility.UseVisualStyleBackColor = true;
            // 
            // GroupBox_BleName
            // 
            this.GroupBox_BleName.Controls.Add(this.Button_Get_BLE_Display_Name);
            this.GroupBox_BleName.Controls.Add(this.Button_Set_BLE_Display_Name);
            this.GroupBox_BleName.Controls.Add(this.Button_Clear_BLE_Display_Name);
            this.GroupBox_BleName.Controls.Add(this.TextBox_BLE_Display_Name);
            this.GroupBox_BleName.Location = new System.Drawing.Point(7, 474);
            this.GroupBox_BleName.Name = "GroupBox_BleName";
            this.GroupBox_BleName.Size = new System.Drawing.Size(353, 80);
            this.GroupBox_BleName.TabIndex = 6;
            this.GroupBox_BleName.TabStop = false;
            this.GroupBox_BleName.Text = "Bluetooth LE Advertising Name";
            // 
            // Button_Get_BLE_Display_Name
            // 
            this.Button_Get_BLE_Display_Name.Location = new System.Drawing.Point(168, 49);
            this.Button_Get_BLE_Display_Name.Name = "Button_Get_BLE_Display_Name";
            this.Button_Get_BLE_Display_Name.Size = new System.Drawing.Size(75, 23);
            this.Button_Get_BLE_Display_Name.TabIndex = 3;
            this.Button_Get_BLE_Display_Name.Text = "Get";
            this.Button_Get_BLE_Display_Name.UseVisualStyleBackColor = true;
            this.Button_Get_BLE_Display_Name.Click += new System.EventHandler(this.Button_Get_BLE_Display_Name_Click);
            // 
            // Button_Set_BLE_Display_Name
            // 
            this.Button_Set_BLE_Display_Name.Location = new System.Drawing.Point(87, 49);
            this.Button_Set_BLE_Display_Name.Name = "Button_Set_BLE_Display_Name";
            this.Button_Set_BLE_Display_Name.Size = new System.Drawing.Size(75, 23);
            this.Button_Set_BLE_Display_Name.TabIndex = 2;
            this.Button_Set_BLE_Display_Name.Text = "Set";
            this.Button_Set_BLE_Display_Name.UseVisualStyleBackColor = true;
            this.Button_Set_BLE_Display_Name.Click += new System.EventHandler(this.Button_Set_BLE_Display_Name_Click);
            // 
            // Button_Clear_BLE_Display_Name
            // 
            this.Button_Clear_BLE_Display_Name.Location = new System.Drawing.Point(6, 49);
            this.Button_Clear_BLE_Display_Name.Name = "Button_Clear_BLE_Display_Name";
            this.Button_Clear_BLE_Display_Name.Size = new System.Drawing.Size(75, 23);
            this.Button_Clear_BLE_Display_Name.TabIndex = 1;
            this.Button_Clear_BLE_Display_Name.Text = "Default";
            this.Button_Clear_BLE_Display_Name.UseVisualStyleBackColor = true;
            this.Button_Clear_BLE_Display_Name.Click += new System.EventHandler(this.Button_Clear_BLE_Display_Name_Click);
            // 
            // TextBox_BLE_Display_Name
            // 
            this.TextBox_BLE_Display_Name.Location = new System.Drawing.Point(6, 21);
            this.TextBox_BLE_Display_Name.Name = "TextBox_BLE_Display_Name";
            this.TextBox_BLE_Display_Name.Size = new System.Drawing.Size(237, 22);
            this.TextBox_BLE_Display_Name.TabIndex = 0;
            this.TextBox_BLE_Display_Name.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_BLE_Display_Name.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TextBox_KeyDown);
            this.TextBox_BLE_Display_Name.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // groupBox_Device
            // 
            this.groupBox_Device.Controls.Add(this.button_restore_fac_ref_warning);
            this.groupBox_Device.Controls.Add(this.Button_LockButton);
            this.groupBox_Device.Controls.Add(this.Label_ButtonStatus);
            this.groupBox_Device.Controls.Add(this.Button_UnlockButton);
            this.groupBox_Device.Controls.Add(this.button_DeviceRestoreFacRef);
            this.groupBox_Device.Controls.Add(this.label_RestoreFacRef);
            this.groupBox_Device.Controls.Add(this.button_DeviceUpdateRef);
            this.groupBox_Device.Controls.Add(this.label119);
            this.groupBox_Device.Controls.Add(this.button_DeviceResetSys);
            this.groupBox_Device.Controls.Add(this.label118);
            this.groupBox_Device.Location = new System.Drawing.Point(883, 385);
            this.groupBox_Device.Name = "groupBox_Device";
            this.groupBox_Device.Size = new System.Drawing.Size(348, 141);
            this.groupBox_Device.TabIndex = 12;
            this.groupBox_Device.TabStop = false;
            this.groupBox_Device.Text = "Device";
            // 
            // button_restore_fac_ref_warning
            // 
            this.button_restore_fac_ref_warning.FlatAppearance.BorderSize = 0;
            this.button_restore_fac_ref_warning.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_restore_fac_ref_warning.Image = global::ISC_Win_WinForm_GUI.Properties.Resources.warning;
            this.button_restore_fac_ref_warning.Location = new System.Drawing.Point(160, 76);
            this.button_restore_fac_ref_warning.Name = "button_restore_fac_ref_warning";
            this.button_restore_fac_ref_warning.Size = new System.Drawing.Size(21, 23);
            this.button_restore_fac_ref_warning.TabIndex = 17;
            this.button_restore_fac_ref_warning.TabStop = false;
            this.button_restore_fac_ref_warning.UseVisualStyleBackColor = true;
            this.button_restore_fac_ref_warning.Click += new System.EventHandler(this.button_restore_fac_ref_warning_Click);
            // 
            // Button_LockButton
            // 
            this.Button_LockButton.Location = new System.Drawing.Point(183, 109);
            this.Button_LockButton.Name = "Button_LockButton";
            this.Button_LockButton.Size = new System.Drawing.Size(75, 23);
            this.Button_LockButton.TabIndex = 4;
            this.Button_LockButton.Text = "Lock";
            this.Button_LockButton.UseVisualStyleBackColor = true;
            this.Button_LockButton.Click += new System.EventHandler(this.Button_LockButton_Click);
            // 
            // Label_ButtonStatus
            // 
            this.Label_ButtonStatus.AutoSize = true;
            this.Label_ButtonStatus.Location = new System.Drawing.Point(6, 113);
            this.Label_ButtonStatus.Name = "Label_ButtonStatus";
            this.Label_ButtonStatus.Size = new System.Drawing.Size(86, 14);
            this.Label_ButtonStatus.TabIndex = 12;
            this.Label_ButtonStatus.Text = "Button Status: ";
            // 
            // Button_UnlockButton
            // 
            this.Button_UnlockButton.Location = new System.Drawing.Point(264, 109);
            this.Button_UnlockButton.Name = "Button_UnlockButton";
            this.Button_UnlockButton.Size = new System.Drawing.Size(78, 23);
            this.Button_UnlockButton.TabIndex = 5;
            this.Button_UnlockButton.Text = "Unlock";
            this.Button_UnlockButton.UseVisualStyleBackColor = true;
            this.Button_UnlockButton.Click += new System.EventHandler(this.Button_UnlockButton_Click);
            // 
            // button_DeviceRestoreFacRef
            // 
            this.button_DeviceRestoreFacRef.Location = new System.Drawing.Point(264, 78);
            this.button_DeviceRestoreFacRef.Name = "button_DeviceRestoreFacRef";
            this.button_DeviceRestoreFacRef.Size = new System.Drawing.Size(78, 23);
            this.button_DeviceRestoreFacRef.TabIndex = 3;
            this.button_DeviceRestoreFacRef.Text = "Restore";
            this.button_DeviceRestoreFacRef.UseVisualStyleBackColor = true;
            this.button_DeviceRestoreFacRef.Click += new System.EventHandler(this.button_DeviceRestoreFacRef_Click);
            // 
            // label_RestoreFacRef
            // 
            this.label_RestoreFacRef.AutoSize = true;
            this.label_RestoreFacRef.Location = new System.Drawing.Point(6, 82);
            this.label_RestoreFacRef.Name = "label_RestoreFacRef";
            this.label_RestoreFacRef.Size = new System.Drawing.Size(148, 14);
            this.label_RestoreFacRef.TabIndex = 6;
            this.label_RestoreFacRef.Text = "Restore Factory Reference";
            // 
            // button_DeviceUpdateRef
            // 
            this.button_DeviceUpdateRef.Location = new System.Drawing.Point(264, 49);
            this.button_DeviceUpdateRef.Name = "button_DeviceUpdateRef";
            this.button_DeviceUpdateRef.Size = new System.Drawing.Size(78, 23);
            this.button_DeviceUpdateRef.TabIndex = 2;
            this.button_DeviceUpdateRef.Text = "Update";
            this.button_DeviceUpdateRef.UseVisualStyleBackColor = true;
            this.button_DeviceUpdateRef.Click += new System.EventHandler(this.button_DeviceUpdateRef_Click);
            // 
            // label119
            // 
            this.label119.AutoSize = true;
            this.label119.Location = new System.Drawing.Point(6, 53);
            this.label119.Name = "label119";
            this.label119.Size = new System.Drawing.Size(134, 14);
            this.label119.TabIndex = 2;
            this.label119.Text = "Update Reference Data";
            // 
            // button_DeviceResetSys
            // 
            this.button_DeviceResetSys.Location = new System.Drawing.Point(264, 21);
            this.button_DeviceResetSys.Name = "button_DeviceResetSys";
            this.button_DeviceResetSys.Size = new System.Drawing.Size(78, 23);
            this.button_DeviceResetSys.TabIndex = 1;
            this.button_DeviceResetSys.Text = "Reset";
            this.button_DeviceResetSys.UseVisualStyleBackColor = true;
            this.button_DeviceResetSys.Click += new System.EventHandler(this.button_DeviceResetSys_Click);
            // 
            // label118
            // 
            this.label118.AutoSize = true;
            this.label118.Location = new System.Drawing.Point(6, 25);
            this.label118.Name = "label118";
            this.label118.Size = new System.Drawing.Size(79, 14);
            this.label118.TabIndex = 0;
            this.label118.Text = "Reset System";
            // 
            // groupBox_ActivationKey
            // 
            this.groupBox_ActivationKey.Controls.Add(this.label_ActivateStatus);
            this.groupBox_ActivationKey.Controls.Add(this.Status);
            this.groupBox_ActivationKey.Controls.Add(this.button_KeySet);
            this.groupBox_ActivationKey.Controls.Add(this.TextBox_Key);
            this.groupBox_ActivationKey.Controls.Add(this.label117);
            this.groupBox_ActivationKey.Location = new System.Drawing.Point(883, 296);
            this.groupBox_ActivationKey.Name = "groupBox_ActivationKey";
            this.groupBox_ActivationKey.Size = new System.Drawing.Size(348, 83);
            this.groupBox_ActivationKey.TabIndex = 11;
            this.groupBox_ActivationKey.TabStop = false;
            this.groupBox_ActivationKey.Text = "Activation Key";
            // 
            // label_ActivateStatus
            // 
            this.label_ActivateStatus.AutoSize = true;
            this.label_ActivateStatus.Location = new System.Drawing.Point(70, 55);
            this.label_ActivateStatus.Name = "label_ActivateStatus";
            this.label_ActivateStatus.Size = new System.Drawing.Size(82, 14);
            this.label_ActivateStatus.TabIndex = 4;
            this.label_ActivateStatus.Text = "Not activated!";
            // 
            // Status
            // 
            this.Status.AutoSize = true;
            this.Status.Location = new System.Drawing.Point(22, 55);
            this.Status.Name = "Status";
            this.Status.Size = new System.Drawing.Size(47, 14);
            this.Status.TabIndex = 3;
            this.Status.Text = "Status :";
            // 
            // button_KeySet
            // 
            this.button_KeySet.Location = new System.Drawing.Point(260, 20);
            this.button_KeySet.Name = "button_KeySet";
            this.button_KeySet.Size = new System.Drawing.Size(78, 23);
            this.button_KeySet.TabIndex = 2;
            this.button_KeySet.Text = "Set Key";
            this.button_KeySet.UseVisualStyleBackColor = true;
            this.button_KeySet.Click += new System.EventHandler(this.button_KeySet_Click);
            // 
            // TextBox_Key
            // 
            this.TextBox_Key.Location = new System.Drawing.Point(73, 20);
            this.TextBox_Key.Name = "TextBox_Key";
            this.TextBox_Key.Size = new System.Drawing.Size(173, 22);
            this.TextBox_Key.TabIndex = 1;
            // 
            // label117
            // 
            this.label117.AutoSize = true;
            this.label117.Location = new System.Drawing.Point(6, 23);
            this.label117.Name = "label117";
            this.label117.Size = new System.Drawing.Size(63, 14);
            this.label117.TabIndex = 0;
            this.label117.Text = "Input Key :";
            // 
            // groupBox_DevInfo
            // 
            this.groupBox_DevInfo.Controls.Add(this.Label_BleNameValue);
            this.groupBox_DevInfo.Controls.Add(this.Label_Blename);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoLampUsageValue);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoLampUsage);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoUUID);
            this.groupBox_DevInfo.Controls.Add(this.label114);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoManfacSerNum);
            this.groupBox_DevInfo.Controls.Add(this.label_MFC_Seri_Num);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoDevSerNum);
            this.groupBox_DevInfo.Controls.Add(this.label104);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoModelName);
            this.groupBox_DevInfo.Controls.Add(this.label106);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoDetectorBoardVer);
            this.groupBox_DevInfo.Controls.Add(this.label108);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoMainBoardVer);
            this.groupBox_DevInfo.Controls.Add(this.label110);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoDLPCVer);
            this.groupBox_DevInfo.Controls.Add(this.label102);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoTivaSWVer);
            this.groupBox_DevInfo.Controls.Add(this.label98);
            this.groupBox_DevInfo.Controls.Add(this.label_DevInfoGUIVer);
            this.groupBox_DevInfo.Controls.Add(this.label95);
            this.groupBox_DevInfo.Location = new System.Drawing.Point(883, 6);
            this.groupBox_DevInfo.Name = "groupBox_DevInfo";
            this.groupBox_DevInfo.Size = new System.Drawing.Size(348, 284);
            this.groupBox_DevInfo.TabIndex = 10;
            this.groupBox_DevInfo.TabStop = false;
            this.groupBox_DevInfo.Text = "Device Information";
            // 
            // Label_BleNameValue
            // 
            this.Label_BleNameValue.Enabled = false;
            this.Label_BleNameValue.Location = new System.Drawing.Point(200, 249);
            this.Label_BleNameValue.Name = "Label_BleNameValue";
            this.Label_BleNameValue.Size = new System.Drawing.Size(142, 14);
            this.Label_BleNameValue.TabIndex = 23;
            this.Label_BleNameValue.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.toolTip1.SetToolTip(this.Label_BleNameValue, "Double Click to Copy Text");
            // 
            // Label_Blename
            // 
            this.Label_Blename.AutoSize = true;
            this.Label_Blename.Enabled = false;
            this.Label_Blename.Location = new System.Drawing.Point(6, 249);
            this.Label_Blename.Name = "Label_Blename";
            this.Label_Blename.Size = new System.Drawing.Size(123, 14);
            this.Label_Blename.TabIndex = 22;
            this.Label_Blename.Text = "BLE Advertising Name";
            // 
            // label_DevInfoLampUsageValue
            // 
            this.label_DevInfoLampUsageValue.Location = new System.Drawing.Point(200, 226);
            this.label_DevInfoLampUsageValue.Name = "label_DevInfoLampUsageValue";
            this.label_DevInfoLampUsageValue.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoLampUsageValue.TabIndex = 21;
            this.toolTip1.SetToolTip(this.label_DevInfoLampUsageValue, "Double Click to Copy Text");
            // 
            // label_DevInfoLampUsage
            // 
            this.label_DevInfoLampUsage.AutoSize = true;
            this.label_DevInfoLampUsage.Location = new System.Drawing.Point(6, 226);
            this.label_DevInfoLampUsage.Name = "label_DevInfoLampUsage";
            this.label_DevInfoLampUsage.Size = new System.Drawing.Size(73, 14);
            this.label_DevInfoLampUsage.TabIndex = 20;
            this.label_DevInfoLampUsage.Text = "Lamp Usage";
            // 
            // label_DevInfoUUID
            // 
            this.label_DevInfoUUID.Location = new System.Drawing.Point(200, 203);
            this.label_DevInfoUUID.Name = "label_DevInfoUUID";
            this.label_DevInfoUUID.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoUUID.TabIndex = 19;
            this.toolTip1.SetToolTip(this.label_DevInfoUUID, "Double Click to Copy Text");
            // 
            // label114
            // 
            this.label114.AutoSize = true;
            this.label114.Location = new System.Drawing.Point(6, 203);
            this.label114.Name = "label114";
            this.label114.Size = new System.Drawing.Size(74, 14);
            this.label114.TabIndex = 18;
            this.label114.Text = "Device UUID";
            // 
            // label_DevInfoManfacSerNum
            // 
            this.label_DevInfoManfacSerNum.Location = new System.Drawing.Point(200, 180);
            this.label_DevInfoManfacSerNum.Name = "label_DevInfoManfacSerNum";
            this.label_DevInfoManfacSerNum.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoManfacSerNum.TabIndex = 17;
            this.toolTip1.SetToolTip(this.label_DevInfoManfacSerNum, "Double Click to Copy Text");
            // 
            // label_MFC_Seri_Num
            // 
            this.label_MFC_Seri_Num.AutoSize = true;
            this.label_MFC_Seri_Num.Location = new System.Drawing.Point(6, 180);
            this.label_MFC_Seri_Num.Name = "label_MFC_Seri_Num";
            this.label_MFC_Seri_Num.Size = new System.Drawing.Size(167, 14);
            this.label_MFC_Seri_Num.TabIndex = 16;
            this.label_MFC_Seri_Num.Text = "Manufacturing Serial Number";
            // 
            // label_DevInfoDevSerNum
            // 
            this.label_DevInfoDevSerNum.Location = new System.Drawing.Point(200, 157);
            this.label_DevInfoDevSerNum.Name = "label_DevInfoDevSerNum";
            this.label_DevInfoDevSerNum.Size = new System.Drawing.Size(142, 13);
            this.label_DevInfoDevSerNum.TabIndex = 15;
            this.toolTip1.SetToolTip(this.label_DevInfoDevSerNum, "Double Click to Copy Text");
            // 
            // label104
            // 
            this.label104.AutoSize = true;
            this.label104.Location = new System.Drawing.Point(6, 157);
            this.label104.Name = "label104";
            this.label104.Size = new System.Drawing.Size(124, 14);
            this.label104.TabIndex = 14;
            this.label104.Text = "Device Serial Number";
            // 
            // label_DevInfoModelName
            // 
            this.label_DevInfoModelName.Location = new System.Drawing.Point(200, 133);
            this.label_DevInfoModelName.Name = "label_DevInfoModelName";
            this.label_DevInfoModelName.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoModelName.TabIndex = 13;
            this.toolTip1.SetToolTip(this.label_DevInfoModelName, "Double Click to Copy Text");
            // 
            // label106
            // 
            this.label106.AutoSize = true;
            this.label106.Location = new System.Drawing.Point(6, 133);
            this.label106.Name = "label106";
            this.label106.Size = new System.Drawing.Size(77, 14);
            this.label106.TabIndex = 12;
            this.label106.Text = "Model Name";
            // 
            // label_DevInfoDetectorBoardVer
            // 
            this.label_DevInfoDetectorBoardVer.Location = new System.Drawing.Point(200, 110);
            this.label_DevInfoDetectorBoardVer.Name = "label_DevInfoDetectorBoardVer";
            this.label_DevInfoDetectorBoardVer.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoDetectorBoardVer.TabIndex = 11;
            this.toolTip1.SetToolTip(this.label_DevInfoDetectorBoardVer, "Double Click to Copy Text");
            // 
            // label108
            // 
            this.label108.AutoSize = true;
            this.label108.Location = new System.Drawing.Point(6, 110);
            this.label108.Name = "label108";
            this.label108.Size = new System.Drawing.Size(132, 14);
            this.label108.TabIndex = 10;
            this.label108.Text = "Detector Board Version";
            // 
            // label_DevInfoMainBoardVer
            // 
            this.label_DevInfoMainBoardVer.Location = new System.Drawing.Point(200, 88);
            this.label_DevInfoMainBoardVer.Name = "label_DevInfoMainBoardVer";
            this.label_DevInfoMainBoardVer.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoMainBoardVer.TabIndex = 9;
            this.toolTip1.SetToolTip(this.label_DevInfoMainBoardVer, "Double Click to Copy Text");
            // 
            // label110
            // 
            this.label110.AutoSize = true;
            this.label110.Location = new System.Drawing.Point(6, 88);
            this.label110.Name = "label110";
            this.label110.Size = new System.Drawing.Size(114, 14);
            this.label110.TabIndex = 8;
            this.label110.Text = "Main Board Version";
            // 
            // label_DevInfoDLPCVer
            // 
            this.label_DevInfoDLPCVer.Location = new System.Drawing.Point(200, 65);
            this.label_DevInfoDLPCVer.Name = "label_DevInfoDLPCVer";
            this.label_DevInfoDLPCVer.Size = new System.Drawing.Size(142, 14);
            this.label_DevInfoDLPCVer.TabIndex = 5;
            this.toolTip1.SetToolTip(this.label_DevInfoDLPCVer, "Double Click to Copy Text");
            // 
            // label102
            // 
            this.label102.AutoSize = true;
            this.label102.Location = new System.Drawing.Point(6, 65);
            this.label102.Name = "label102";
            this.label102.Size = new System.Drawing.Size(109, 14);
            this.label102.TabIndex = 4;
            this.label102.Text = "DLPC Flash Version";
            // 
            // label_DevInfoTivaSWVer
            // 
            this.label_DevInfoTivaSWVer.Location = new System.Drawing.Point(200, 41);
            this.label_DevInfoTivaSWVer.Name = "label_DevInfoTivaSWVer";
            this.label_DevInfoTivaSWVer.Size = new System.Drawing.Size(142, 15);
            this.label_DevInfoTivaSWVer.TabIndex = 3;
            this.toolTip1.SetToolTip(this.label_DevInfoTivaSWVer, "Double Click to Copy Text");
            // 
            // label98
            // 
            this.label98.AutoSize = true;
            this.label98.Location = new System.Drawing.Point(6, 42);
            this.label98.Name = "label98";
            this.label98.Size = new System.Drawing.Size(93, 14);
            this.label98.TabIndex = 2;
            this.label98.Text = "Tiva SW Version";
            // 
            // label_DevInfoGUIVer
            // 
            this.label_DevInfoGUIVer.Location = new System.Drawing.Point(200, 18);
            this.label_DevInfoGUIVer.Name = "label_DevInfoGUIVer";
            this.label_DevInfoGUIVer.Size = new System.Drawing.Size(142, 15);
            this.label_DevInfoGUIVer.TabIndex = 1;
            this.toolTip1.SetToolTip(this.label_DevInfoGUIVer, "Double Click to Copy Text");
            // 
            // label95
            // 
            this.label95.AutoSize = true;
            this.label95.Location = new System.Drawing.Point(6, 19);
            this.label95.Name = "label95";
            this.label95.Size = new System.Drawing.Size(71, 14);
            this.label95.TabIndex = 0;
            this.label95.Text = "GUI Version";
            // 
            // GroupBox_CalibCoeffs
            // 
            this.GroupBox_CalibCoeffs.Controls.Add(this.Button_CalWriteCoeffs);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Button_CalReadCoeffs);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Button_CalRestoreDefaultCoeffs);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Button_CalWriteGenCoeffs);
            this.GroupBox_CalibCoeffs.Controls.Add(this.CheckBox_CalWriteEnable);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_ShiftVectCoeff2);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label28);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_ShiftVectCoeff1);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label31);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_ShiftVectCoeff0);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label22);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_P2WCoeff2);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label25);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Label_ScanCfgVer);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label27);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_P2WCoeff1);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label19);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Label_RefCalVer);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label21);
            this.GroupBox_CalibCoeffs.Controls.Add(this.TextBox_P2WCoeff0);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label23);
            this.GroupBox_CalibCoeffs.Controls.Add(this.Label_CalCoeffVer);
            this.GroupBox_CalibCoeffs.Controls.Add(this.label29);
            this.GroupBox_CalibCoeffs.Location = new System.Drawing.Point(366, 182);
            this.GroupBox_CalibCoeffs.Name = "GroupBox_CalibCoeffs";
            this.GroupBox_CalibCoeffs.Size = new System.Drawing.Size(511, 372);
            this.GroupBox_CalibCoeffs.TabIndex = 9;
            this.GroupBox_CalibCoeffs.TabStop = false;
            this.GroupBox_CalibCoeffs.Text = "Calibration Coefficients";
            // 
            // Button_CalWriteCoeffs
            // 
            this.Button_CalWriteCoeffs.Enabled = false;
            this.Button_CalWriteCoeffs.Location = new System.Drawing.Point(304, 253);
            this.Button_CalWriteCoeffs.Name = "Button_CalWriteCoeffs";
            this.Button_CalWriteCoeffs.Size = new System.Drawing.Size(115, 23);
            this.Button_CalWriteCoeffs.TabIndex = 10;
            this.Button_CalWriteCoeffs.Text = "Write Coeffs";
            this.Button_CalWriteCoeffs.UseVisualStyleBackColor = true;
            this.Button_CalWriteCoeffs.Click += new System.EventHandler(this.Button_CalWriteCoeffs_Click);
            // 
            // Button_CalReadCoeffs
            // 
            this.Button_CalReadCoeffs.Location = new System.Drawing.Point(206, 253);
            this.Button_CalReadCoeffs.Name = "Button_CalReadCoeffs";
            this.Button_CalReadCoeffs.Size = new System.Drawing.Size(92, 23);
            this.Button_CalReadCoeffs.TabIndex = 9;
            this.Button_CalReadCoeffs.Text = "Read Coeffs";
            this.Button_CalReadCoeffs.UseVisualStyleBackColor = true;
            this.Button_CalReadCoeffs.Click += new System.EventHandler(this.Button_CalReadCoeffs_Click);
            // 
            // Button_CalRestoreDefaultCoeffs
            // 
            this.Button_CalRestoreDefaultCoeffs.Enabled = false;
            this.Button_CalRestoreDefaultCoeffs.Location = new System.Drawing.Point(9, 253);
            this.Button_CalRestoreDefaultCoeffs.Name = "Button_CalRestoreDefaultCoeffs";
            this.Button_CalRestoreDefaultCoeffs.Size = new System.Drawing.Size(191, 23);
            this.Button_CalRestoreDefaultCoeffs.TabIndex = 8;
            this.Button_CalRestoreDefaultCoeffs.Text = "Restore Factory Calibration Data";
            this.Button_CalRestoreDefaultCoeffs.UseVisualStyleBackColor = true;
            this.Button_CalRestoreDefaultCoeffs.Click += new System.EventHandler(this.Button_CalRestoreDefaultCoeffs_Click);
            // 
            // Button_CalWriteGenCoeffs
            // 
            this.Button_CalWriteGenCoeffs.Enabled = false;
            this.Button_CalWriteGenCoeffs.Location = new System.Drawing.Point(9, 217);
            this.Button_CalWriteGenCoeffs.Name = "Button_CalWriteGenCoeffs";
            this.Button_CalWriteGenCoeffs.Size = new System.Drawing.Size(106, 23);
            this.Button_CalWriteGenCoeffs.TabIndex = 7;
            this.Button_CalWriteGenCoeffs.Text = "Write Generic Data";
            this.Button_CalWriteGenCoeffs.UseVisualStyleBackColor = true;
            this.Button_CalWriteGenCoeffs.Click += new System.EventHandler(this.Button_CalWriteGenCoeffs_Click);
            // 
            // CheckBox_CalWriteEnable
            // 
            this.CheckBox_CalWriteEnable.AutoSize = true;
            this.CheckBox_CalWriteEnable.Location = new System.Drawing.Point(9, 181);
            this.CheckBox_CalWriteEnable.Name = "CheckBox_CalWriteEnable";
            this.CheckBox_CalWriteEnable.Size = new System.Drawing.Size(97, 18);
            this.CheckBox_CalWriteEnable.TabIndex = 6;
            this.CheckBox_CalWriteEnable.Text = "Write Enable";
            this.CheckBox_CalWriteEnable.UseVisualStyleBackColor = true;
            this.CheckBox_CalWriteEnable.CheckedChanged += new System.EventHandler(this.CheckBox_CalWriteEnable_CheckedChanged);
            // 
            // TextBox_ShiftVectCoeff2
            // 
            this.TextBox_ShiftVectCoeff2.Location = new System.Drawing.Point(265, 217);
            this.TextBox_ShiftVectCoeff2.Name = "TextBox_ShiftVectCoeff2";
            this.TextBox_ShiftVectCoeff2.Size = new System.Drawing.Size(127, 22);
            this.TextBox_ShiftVectCoeff2.TabIndex = 5;
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Location = new System.Drawing.Point(136, 222);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(95, 14);
            this.label28.TabIndex = 22;
            this.label28.Text = "Shift Vect Coeff 2";
            // 
            // TextBox_ShiftVectCoeff1
            // 
            this.TextBox_ShiftVectCoeff1.Location = new System.Drawing.Point(265, 177);
            this.TextBox_ShiftVectCoeff1.Name = "TextBox_ShiftVectCoeff1";
            this.TextBox_ShiftVectCoeff1.Size = new System.Drawing.Size(127, 22);
            this.TextBox_ShiftVectCoeff1.TabIndex = 4;
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(136, 182);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(95, 14);
            this.label31.TabIndex = 18;
            this.label31.Text = "Shift Vect Coeff 1";
            // 
            // TextBox_ShiftVectCoeff0
            // 
            this.TextBox_ShiftVectCoeff0.Location = new System.Drawing.Point(265, 137);
            this.TextBox_ShiftVectCoeff0.Name = "TextBox_ShiftVectCoeff0";
            this.TextBox_ShiftVectCoeff0.Size = new System.Drawing.Size(127, 22);
            this.TextBox_ShiftVectCoeff0.TabIndex = 3;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(136, 142);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(95, 14);
            this.label22.TabIndex = 14;
            this.label22.Text = "Shift Vect Coeff 0";
            // 
            // TextBox_P2WCoeff2
            // 
            this.TextBox_P2WCoeff2.Location = new System.Drawing.Point(265, 97);
            this.TextBox_P2WCoeff2.Name = "TextBox_P2WCoeff2";
            this.TextBox_P2WCoeff2.Size = new System.Drawing.Size(127, 22);
            this.TextBox_P2WCoeff2.TabIndex = 2;
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(136, 102);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(95, 14);
            this.label25.TabIndex = 10;
            this.label25.Text = "Pix-Wave Coeff 2";
            // 
            // Label_ScanCfgVer
            // 
            this.Label_ScanCfgVer.AutoSize = true;
            this.Label_ScanCfgVer.Location = new System.Drawing.Point(104, 102);
            this.Label_ScanCfgVer.Name = "Label_ScanCfgVer";
            this.Label_ScanCfgVer.Size = new System.Drawing.Size(13, 14);
            this.Label_ScanCfgVer.TabIndex = 9;
            this.Label_ScanCfgVer.Text = "0";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(7, 102);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(77, 14);
            this.label27.TabIndex = 8;
            this.label27.Text = "Scan Cfg Ver :";
            // 
            // TextBox_P2WCoeff1
            // 
            this.TextBox_P2WCoeff1.Location = new System.Drawing.Point(265, 57);
            this.TextBox_P2WCoeff1.Name = "TextBox_P2WCoeff1";
            this.TextBox_P2WCoeff1.Size = new System.Drawing.Size(127, 22);
            this.TextBox_P2WCoeff1.TabIndex = 1;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(136, 62);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(95, 14);
            this.label19.TabIndex = 6;
            this.label19.Text = "Pix-Wave Coeff 1";
            // 
            // Label_RefCalVer
            // 
            this.Label_RefCalVer.AutoSize = true;
            this.Label_RefCalVer.Location = new System.Drawing.Point(104, 62);
            this.Label_RefCalVer.Name = "Label_RefCalVer";
            this.Label_RefCalVer.Size = new System.Drawing.Size(13, 14);
            this.Label_RefCalVer.TabIndex = 5;
            this.Label_RefCalVer.Text = "0";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(7, 62);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(71, 14);
            this.label21.TabIndex = 4;
            this.label21.Text = "Ref Cal Ver :";
            // 
            // TextBox_P2WCoeff0
            // 
            this.TextBox_P2WCoeff0.Location = new System.Drawing.Point(265, 17);
            this.TextBox_P2WCoeff0.Name = "TextBox_P2WCoeff0";
            this.TextBox_P2WCoeff0.Size = new System.Drawing.Size(127, 22);
            this.TextBox_P2WCoeff0.TabIndex = 0;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(136, 22);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(95, 14);
            this.label23.TabIndex = 2;
            this.label23.Text = "Pix-Wave Coeff 0";
            // 
            // Label_CalCoeffVer
            // 
            this.Label_CalCoeffVer.AutoSize = true;
            this.Label_CalCoeffVer.Location = new System.Drawing.Point(104, 22);
            this.Label_CalCoeffVer.Name = "Label_CalCoeffVer";
            this.Label_CalCoeffVer.Size = new System.Drawing.Size(13, 14);
            this.Label_CalCoeffVer.TabIndex = 1;
            this.Label_CalCoeffVer.Text = "0";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(7, 22);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(80, 14);
            this.label29.TabIndex = 0;
            this.label29.Text = "Cal Coeff Ver :";
            // 
            // GroupBox_Sensors
            // 
            this.GroupBox_Sensors.Controls.Add(this.Button_SensorRead);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampCM2Value);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampCM1Value);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampVM2Value);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampVM1Value);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorTivaTemp);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorSysTemp);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorHumidity);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorBattCapacity);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorBattStatus);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampCM2);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampCM1);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampVM2);
            this.GroupBox_Sensors.Controls.Add(this.Label_SensorLampVM1);
            this.GroupBox_Sensors.Controls.Add(this.label8);
            this.GroupBox_Sensors.Controls.Add(this.label5);
            this.GroupBox_Sensors.Controls.Add(this.label_sys_humidity);
            this.GroupBox_Sensors.Controls.Add(this.lb_BattCapTitle);
            this.GroupBox_Sensors.Controls.Add(this.lb_BattChargerStatusTitle);
            this.GroupBox_Sensors.Location = new System.Drawing.Point(8, 182);
            this.GroupBox_Sensors.Name = "GroupBox_Sensors";
            this.GroupBox_Sensors.Size = new System.Drawing.Size(352, 285);
            this.GroupBox_Sensors.TabIndex = 5;
            this.GroupBox_Sensors.TabStop = false;
            this.GroupBox_Sensors.Text = "Sensors";
            // 
            // Button_SensorRead
            // 
            this.Button_SensorRead.Location = new System.Drawing.Point(271, 253);
            this.Button_SensorRead.Name = "Button_SensorRead";
            this.Button_SensorRead.Size = new System.Drawing.Size(75, 23);
            this.Button_SensorRead.TabIndex = 0;
            this.Button_SensorRead.Text = "Read";
            this.Button_SensorRead.UseVisualStyleBackColor = true;
            this.Button_SensorRead.Click += new System.EventHandler(this.Button_SensorRead_Click);
            // 
            // Label_SensorLampCM2Value
            // 
            this.Label_SensorLampCM2Value.Location = new System.Drawing.Point(176, 218);
            this.Label_SensorLampCM2Value.Name = "Label_SensorLampCM2Value";
            this.Label_SensorLampCM2Value.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorLampCM2Value.TabIndex = 19;
            this.toolTip1.SetToolTip(this.Label_SensorLampCM2Value, "Double Click to Copy Text");
            // 
            // Label_SensorLampCM1Value
            // 
            this.Label_SensorLampCM1Value.Location = new System.Drawing.Point(177, 168);
            this.Label_SensorLampCM1Value.Name = "Label_SensorLampCM1Value";
            this.Label_SensorLampCM1Value.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorLampCM1Value.TabIndex = 18;
            this.toolTip1.SetToolTip(this.Label_SensorLampCM1Value, "Double Click to Copy Text");
            // 
            // Label_SensorLampVM2Value
            // 
            this.Label_SensorLampVM2Value.Location = new System.Drawing.Point(177, 193);
            this.Label_SensorLampVM2Value.Name = "Label_SensorLampVM2Value";
            this.Label_SensorLampVM2Value.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorLampVM2Value.TabIndex = 17;
            this.toolTip1.SetToolTip(this.Label_SensorLampVM2Value, "Double Click to Copy Text");
            // 
            // Label_SensorLampVM1Value
            // 
            this.Label_SensorLampVM1Value.Location = new System.Drawing.Point(177, 143);
            this.Label_SensorLampVM1Value.Name = "Label_SensorLampVM1Value";
            this.Label_SensorLampVM1Value.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorLampVM1Value.TabIndex = 12;
            this.toolTip1.SetToolTip(this.Label_SensorLampVM1Value, "Double Click to Copy Text");
            // 
            // Label_SensorTivaTemp
            // 
            this.Label_SensorTivaTemp.Location = new System.Drawing.Point(177, 118);
            this.Label_SensorTivaTemp.Name = "Label_SensorTivaTemp";
            this.Label_SensorTivaTemp.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorTivaTemp.TabIndex = 11;
            this.toolTip1.SetToolTip(this.Label_SensorTivaTemp, "Double Click to Copy Text");
            // 
            // Label_SensorSysTemp
            // 
            this.Label_SensorSysTemp.Location = new System.Drawing.Point(177, 93);
            this.Label_SensorSysTemp.Name = "Label_SensorSysTemp";
            this.Label_SensorSysTemp.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorSysTemp.TabIndex = 10;
            this.toolTip1.SetToolTip(this.Label_SensorSysTemp, "Double Click to Copy Text");
            // 
            // Label_SensorHumidity
            // 
            this.Label_SensorHumidity.Location = new System.Drawing.Point(177, 68);
            this.Label_SensorHumidity.Name = "Label_SensorHumidity";
            this.Label_SensorHumidity.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorHumidity.TabIndex = 9;
            this.toolTip1.SetToolTip(this.Label_SensorHumidity, "Double Click to Copy Text");
            // 
            // Label_SensorBattCapacity
            // 
            this.Label_SensorBattCapacity.Location = new System.Drawing.Point(177, 43);
            this.Label_SensorBattCapacity.Name = "Label_SensorBattCapacity";
            this.Label_SensorBattCapacity.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorBattCapacity.TabIndex = 8;
            this.toolTip1.SetToolTip(this.Label_SensorBattCapacity, "Double Click to Copy Text");
            // 
            // Label_SensorBattStatus
            // 
            this.Label_SensorBattStatus.Location = new System.Drawing.Point(177, 18);
            this.Label_SensorBattStatus.Name = "Label_SensorBattStatus";
            this.Label_SensorBattStatus.Size = new System.Drawing.Size(129, 14);
            this.Label_SensorBattStatus.TabIndex = 7;
            this.toolTip1.SetToolTip(this.Label_SensorBattStatus, "Double Click to Copy Text");
            // 
            // Label_SensorLampCM2
            // 
            this.Label_SensorLampCM2.AutoSize = true;
            this.Label_SensorLampCM2.Location = new System.Drawing.Point(6, 218);
            this.Label_SensorLampCM2.Name = "Label_SensorLampCM2";
            this.Label_SensorLampCM2.Size = new System.Drawing.Size(61, 14);
            this.Label_SensorLampCM2.TabIndex = 16;
            this.Label_SensorLampCM2.Text = "Lamp CM2";
            // 
            // Label_SensorLampCM1
            // 
            this.Label_SensorLampCM1.AutoSize = true;
            this.Label_SensorLampCM1.Location = new System.Drawing.Point(6, 168);
            this.Label_SensorLampCM1.Name = "Label_SensorLampCM1";
            this.Label_SensorLampCM1.Size = new System.Drawing.Size(61, 14);
            this.Label_SensorLampCM1.TabIndex = 15;
            this.Label_SensorLampCM1.Text = "Lamp CM1";
            // 
            // Label_SensorLampVM2
            // 
            this.Label_SensorLampVM2.AutoSize = true;
            this.Label_SensorLampVM2.Location = new System.Drawing.Point(6, 193);
            this.Label_SensorLampVM2.Name = "Label_SensorLampVM2";
            this.Label_SensorLampVM2.Size = new System.Drawing.Size(62, 14);
            this.Label_SensorLampVM2.TabIndex = 14;
            this.Label_SensorLampVM2.Text = "Lamp VM2";
            // 
            // Label_SensorLampVM1
            // 
            this.Label_SensorLampVM1.AutoSize = true;
            this.Label_SensorLampVM1.Location = new System.Drawing.Point(6, 143);
            this.Label_SensorLampVM1.Name = "Label_SensorLampVM1";
            this.Label_SensorLampVM1.Size = new System.Drawing.Size(62, 14);
            this.Label_SensorLampVM1.TabIndex = 5;
            this.Label_SensorLampVM1.Text = "Lamp VM1";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 118);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(61, 14);
            this.label8.TabIndex = 4;
            this.label8.Text = "Tiva Temp";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 93);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 14);
            this.label5.TabIndex = 3;
            this.label5.Text = "System Temp";
            // 
            // label_sys_humidity
            // 
            this.label_sys_humidity.AutoSize = true;
            this.label_sys_humidity.Location = new System.Drawing.Point(6, 68);
            this.label_sys_humidity.Name = "label_sys_humidity";
            this.label_sys_humidity.Size = new System.Drawing.Size(97, 14);
            this.label_sys_humidity.TabIndex = 2;
            this.label_sys_humidity.Text = "System Humidity";
            // 
            // lb_BattCapTitle
            // 
            this.lb_BattCapTitle.AutoSize = true;
            this.lb_BattCapTitle.Location = new System.Drawing.Point(6, 43);
            this.lb_BattCapTitle.Name = "lb_BattCapTitle";
            this.lb_BattCapTitle.Size = new System.Drawing.Size(93, 14);
            this.lb_BattCapTitle.TabIndex = 1;
            this.lb_BattCapTitle.Text = "Battery Capacity";
            // 
            // lb_BattChargerStatusTitle
            // 
            this.lb_BattChargerStatusTitle.AutoSize = true;
            this.lb_BattChargerStatusTitle.Location = new System.Drawing.Point(6, 18);
            this.lb_BattChargerStatusTitle.Name = "lb_BattChargerStatusTitle";
            this.lb_BattChargerStatusTitle.Size = new System.Drawing.Size(129, 14);
            this.lb_BattChargerStatusTitle.TabIndex = 0;
            this.lb_BattChargerStatusTitle.Text = "Battery Changer Status";
            // 
            // GroupBox_DateTime
            // 
            this.GroupBox_DateTime.Controls.Add(this.Button_DateTimeGet);
            this.GroupBox_DateTime.Controls.Add(this.Button_DateTimeSync);
            this.GroupBox_DateTime.Controls.Add(this.TextBox_DateTime);
            this.GroupBox_DateTime.Location = new System.Drawing.Point(8, 94);
            this.GroupBox_DateTime.Name = "GroupBox_DateTime";
            this.GroupBox_DateTime.Size = new System.Drawing.Size(173, 82);
            this.GroupBox_DateTime.TabIndex = 9;
            this.GroupBox_DateTime.TabStop = false;
            this.GroupBox_DateTime.Text = "Date and Time";
            // 
            // Button_DateTimeGet
            // 
            this.Button_DateTimeGet.Location = new System.Drawing.Point(88, 51);
            this.Button_DateTimeGet.Name = "Button_DateTimeGet";
            this.Button_DateTimeGet.Size = new System.Drawing.Size(75, 23);
            this.Button_DateTimeGet.TabIndex = 2;
            this.Button_DateTimeGet.Text = "Get";
            this.Button_DateTimeGet.UseVisualStyleBackColor = true;
            this.Button_DateTimeGet.Click += new System.EventHandler(this.Button_DateTimeGet_Click);
            // 
            // Button_DateTimeSync
            // 
            this.Button_DateTimeSync.Location = new System.Drawing.Point(7, 51);
            this.Button_DateTimeSync.Name = "Button_DateTimeSync";
            this.Button_DateTimeSync.Size = new System.Drawing.Size(75, 23);
            this.Button_DateTimeSync.TabIndex = 1;
            this.Button_DateTimeSync.Text = "Sync";
            this.Button_DateTimeSync.UseVisualStyleBackColor = true;
            this.Button_DateTimeSync.Click += new System.EventHandler(this.Button_DateTimeSync_Click);
            // 
            // TextBox_DateTime
            // 
            this.TextBox_DateTime.Location = new System.Drawing.Point(7, 22);
            this.TextBox_DateTime.Name = "TextBox_DateTime";
            this.TextBox_DateTime.Size = new System.Drawing.Size(156, 22);
            this.TextBox_DateTime.TabIndex = 0;
            // 
            // GroupBox_LampUsage
            // 
            this.GroupBox_LampUsage.Controls.Add(this.label4);
            this.GroupBox_LampUsage.Controls.Add(this.Button_LampUsageGet);
            this.GroupBox_LampUsage.Controls.Add(this.Button_LampUsageSet);
            this.GroupBox_LampUsage.Controls.Add(this.TextBox_LampUsage);
            this.GroupBox_LampUsage.Location = new System.Drawing.Point(187, 94);
            this.GroupBox_LampUsage.Name = "GroupBox_LampUsage";
            this.GroupBox_LampUsage.Size = new System.Drawing.Size(173, 82);
            this.GroupBox_LampUsage.TabIndex = 4;
            this.GroupBox_LampUsage.TabStop = false;
            this.GroupBox_LampUsage.Text = "Lamp Usage";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(117, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 14);
            this.label4.TabIndex = 3;
            this.label4.Text = "(hours)";
            // 
            // Button_LampUsageGet
            // 
            this.Button_LampUsageGet.Location = new System.Drawing.Point(88, 51);
            this.Button_LampUsageGet.Name = "Button_LampUsageGet";
            this.Button_LampUsageGet.Size = new System.Drawing.Size(75, 23);
            this.Button_LampUsageGet.TabIndex = 2;
            this.Button_LampUsageGet.Text = "Get";
            this.Button_LampUsageGet.UseVisualStyleBackColor = true;
            this.Button_LampUsageGet.Click += new System.EventHandler(this.Button_LampUsageGet_Click);
            // 
            // Button_LampUsageSet
            // 
            this.Button_LampUsageSet.Location = new System.Drawing.Point(7, 51);
            this.Button_LampUsageSet.Name = "Button_LampUsageSet";
            this.Button_LampUsageSet.Size = new System.Drawing.Size(75, 23);
            this.Button_LampUsageSet.TabIndex = 1;
            this.Button_LampUsageSet.Text = "Set";
            this.Button_LampUsageSet.UseVisualStyleBackColor = true;
            this.Button_LampUsageSet.Click += new System.EventHandler(this.Button_LampUsageSet_Click);
            // 
            // TextBox_LampUsage
            // 
            this.TextBox_LampUsage.Location = new System.Drawing.Point(7, 22);
            this.TextBox_LampUsage.Name = "TextBox_LampUsage";
            this.TextBox_LampUsage.Size = new System.Drawing.Size(104, 22);
            this.TextBox_LampUsage.TabIndex = 0;
            // 
            // GroupBox_DLPC150FWUpdate
            // 
            this.GroupBox_DLPC150FWUpdate.Controls.Add(this.Button_DLPC150FWUpdate);
            this.GroupBox_DLPC150FWUpdate.Controls.Add(this.ProgressBar_DLPC150FWUpdateStatus);
            this.GroupBox_DLPC150FWUpdate.Controls.Add(this.Button_DLPC150FWBrowse);
            this.GroupBox_DLPC150FWUpdate.Controls.Add(this.TextBox_DLPC150FWPath);
            this.GroupBox_DLPC150FWUpdate.Controls.Add(this.label9);
            this.GroupBox_DLPC150FWUpdate.Location = new System.Drawing.Point(366, 94);
            this.GroupBox_DLPC150FWUpdate.Name = "GroupBox_DLPC150FWUpdate";
            this.GroupBox_DLPC150FWUpdate.Size = new System.Drawing.Size(511, 82);
            this.GroupBox_DLPC150FWUpdate.TabIndex = 8;
            this.GroupBox_DLPC150FWUpdate.TabStop = false;
            this.GroupBox_DLPC150FWUpdate.Text = "DLPC150 Firmware Update";
            // 
            // Button_DLPC150FWUpdate
            // 
            this.Button_DLPC150FWUpdate.Location = new System.Drawing.Point(425, 51);
            this.Button_DLPC150FWUpdate.Name = "Button_DLPC150FWUpdate";
            this.Button_DLPC150FWUpdate.Size = new System.Drawing.Size(75, 23);
            this.Button_DLPC150FWUpdate.TabIndex = 3;
            this.Button_DLPC150FWUpdate.Text = "Update";
            this.Button_DLPC150FWUpdate.UseVisualStyleBackColor = true;
            this.Button_DLPC150FWUpdate.Click += new System.EventHandler(this.Button_DLPC150FWUpdate_Click);
            // 
            // ProgressBar_DLPC150FWUpdateStatus
            // 
            this.ProgressBar_DLPC150FWUpdateStatus.Location = new System.Drawing.Point(6, 51);
            this.ProgressBar_DLPC150FWUpdateStatus.Name = "ProgressBar_DLPC150FWUpdateStatus";
            this.ProgressBar_DLPC150FWUpdateStatus.Size = new System.Drawing.Size(413, 23);
            this.ProgressBar_DLPC150FWUpdateStatus.TabIndex = 4;
            // 
            // Button_DLPC150FWBrowse
            // 
            this.Button_DLPC150FWBrowse.Location = new System.Drawing.Point(425, 16);
            this.Button_DLPC150FWBrowse.Name = "Button_DLPC150FWBrowse";
            this.Button_DLPC150FWBrowse.Size = new System.Drawing.Size(75, 23);
            this.Button_DLPC150FWBrowse.TabIndex = 2;
            this.Button_DLPC150FWBrowse.Text = "Browse";
            this.Button_DLPC150FWBrowse.UseVisualStyleBackColor = true;
            this.Button_DLPC150FWBrowse.Click += new System.EventHandler(this.Button_DLPC150FWBrowse_Click);
            // 
            // TextBox_DLPC150FWPath
            // 
            this.TextBox_DLPC150FWPath.Location = new System.Drawing.Point(76, 19);
            this.TextBox_DLPC150FWPath.Name = "TextBox_DLPC150FWPath";
            this.TextBox_DLPC150FWPath.Size = new System.Drawing.Size(343, 22);
            this.TextBox_DLPC150FWPath.TabIndex = 1;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(7, 22);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(63, 14);
            this.label9.TabIndex = 0;
            this.label9.Text = "File Name";
            // 
            // GroupBox_TivaFWUpdate
            // 
            this.GroupBox_TivaFWUpdate.Controls.Add(this.Button_TivaFWUpdate);
            this.GroupBox_TivaFWUpdate.Controls.Add(this.ProgressBar_TivaFWUpdateStatus);
            this.GroupBox_TivaFWUpdate.Controls.Add(this.Button_TivaFWBrowse);
            this.GroupBox_TivaFWUpdate.Controls.Add(this.TextBox_TivaFWPath);
            this.GroupBox_TivaFWUpdate.Controls.Add(this.label6);
            this.GroupBox_TivaFWUpdate.Location = new System.Drawing.Point(366, 6);
            this.GroupBox_TivaFWUpdate.Name = "GroupBox_TivaFWUpdate";
            this.GroupBox_TivaFWUpdate.Size = new System.Drawing.Size(511, 82);
            this.GroupBox_TivaFWUpdate.TabIndex = 7;
            this.GroupBox_TivaFWUpdate.TabStop = false;
            this.GroupBox_TivaFWUpdate.Text = "TIVA Firmware Update";
            // 
            // Button_TivaFWUpdate
            // 
            this.Button_TivaFWUpdate.Location = new System.Drawing.Point(425, 51);
            this.Button_TivaFWUpdate.Name = "Button_TivaFWUpdate";
            this.Button_TivaFWUpdate.Size = new System.Drawing.Size(75, 23);
            this.Button_TivaFWUpdate.TabIndex = 3;
            this.Button_TivaFWUpdate.Text = "Update";
            this.Button_TivaFWUpdate.UseVisualStyleBackColor = true;
            this.Button_TivaFWUpdate.Click += new System.EventHandler(this.Button_TivaFWUpdate_Click);
            // 
            // ProgressBar_TivaFWUpdateStatus
            // 
            this.ProgressBar_TivaFWUpdateStatus.Location = new System.Drawing.Point(6, 51);
            this.ProgressBar_TivaFWUpdateStatus.Name = "ProgressBar_TivaFWUpdateStatus";
            this.ProgressBar_TivaFWUpdateStatus.Size = new System.Drawing.Size(413, 23);
            this.ProgressBar_TivaFWUpdateStatus.TabIndex = 4;
            // 
            // Button_TivaFWBrowse
            // 
            this.Button_TivaFWBrowse.Location = new System.Drawing.Point(425, 18);
            this.Button_TivaFWBrowse.Name = "Button_TivaFWBrowse";
            this.Button_TivaFWBrowse.Size = new System.Drawing.Size(75, 23);
            this.Button_TivaFWBrowse.TabIndex = 2;
            this.Button_TivaFWBrowse.Text = "Browse";
            this.Button_TivaFWBrowse.UseVisualStyleBackColor = true;
            this.Button_TivaFWBrowse.Click += new System.EventHandler(this.Button_TivaFWBrowse_Click);
            // 
            // TextBox_TivaFWPath
            // 
            this.TextBox_TivaFWPath.Location = new System.Drawing.Point(76, 19);
            this.TextBox_TivaFWPath.Name = "TextBox_TivaFWPath";
            this.TextBox_TivaFWPath.Size = new System.Drawing.Size(343, 22);
            this.TextBox_TivaFWPath.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(63, 14);
            this.label6.TabIndex = 0;
            this.label6.Text = "File Name";
            // 
            // GroupBox_SerialNumber
            // 
            this.GroupBox_SerialNumber.Controls.Add(this.Button_SerialNumberGet);
            this.GroupBox_SerialNumber.Controls.Add(this.Button_SerialNumberSet);
            this.GroupBox_SerialNumber.Controls.Add(this.TextBox_SerialNumber);
            this.GroupBox_SerialNumber.Location = new System.Drawing.Point(187, 6);
            this.GroupBox_SerialNumber.Name = "GroupBox_SerialNumber";
            this.GroupBox_SerialNumber.Size = new System.Drawing.Size(173, 82);
            this.GroupBox_SerialNumber.TabIndex = 2;
            this.GroupBox_SerialNumber.TabStop = false;
            this.GroupBox_SerialNumber.Text = "Serial Number";
            // 
            // Button_SerialNumberGet
            // 
            this.Button_SerialNumberGet.Location = new System.Drawing.Point(88, 51);
            this.Button_SerialNumberGet.Name = "Button_SerialNumberGet";
            this.Button_SerialNumberGet.Size = new System.Drawing.Size(75, 23);
            this.Button_SerialNumberGet.TabIndex = 2;
            this.Button_SerialNumberGet.Text = "Get";
            this.Button_SerialNumberGet.UseVisualStyleBackColor = true;
            this.Button_SerialNumberGet.Click += new System.EventHandler(this.Button_SerialNumberGet_Click);
            // 
            // Button_SerialNumberSet
            // 
            this.Button_SerialNumberSet.Location = new System.Drawing.Point(7, 51);
            this.Button_SerialNumberSet.Name = "Button_SerialNumberSet";
            this.Button_SerialNumberSet.Size = new System.Drawing.Size(75, 23);
            this.Button_SerialNumberSet.TabIndex = 1;
            this.Button_SerialNumberSet.Text = "Set";
            this.Button_SerialNumberSet.UseVisualStyleBackColor = true;
            this.Button_SerialNumberSet.Click += new System.EventHandler(this.Button_SerialNumberSet_Click);
            // 
            // TextBox_SerialNumber
            // 
            this.TextBox_SerialNumber.Location = new System.Drawing.Point(7, 22);
            this.TextBox_SerialNumber.Name = "TextBox_SerialNumber";
            this.TextBox_SerialNumber.Size = new System.Drawing.Size(156, 22);
            this.TextBox_SerialNumber.TabIndex = 0;
            this.TextBox_SerialNumber.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_SerialNumber.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TextBox_KeyDown);
            this.TextBox_SerialNumber.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // GroupBox_ModelName
            // 
            this.GroupBox_ModelName.Controls.Add(this.Button_ModelNameGet);
            this.GroupBox_ModelName.Controls.Add(this.Button_ModelNameSet);
            this.GroupBox_ModelName.Controls.Add(this.TextBox_ModelName);
            this.GroupBox_ModelName.Location = new System.Drawing.Point(8, 6);
            this.GroupBox_ModelName.Name = "GroupBox_ModelName";
            this.GroupBox_ModelName.Size = new System.Drawing.Size(173, 82);
            this.GroupBox_ModelName.TabIndex = 1;
            this.GroupBox_ModelName.TabStop = false;
            this.GroupBox_ModelName.Text = "Model Name";
            // 
            // Button_ModelNameGet
            // 
            this.Button_ModelNameGet.Location = new System.Drawing.Point(88, 51);
            this.Button_ModelNameGet.Name = "Button_ModelNameGet";
            this.Button_ModelNameGet.Size = new System.Drawing.Size(75, 23);
            this.Button_ModelNameGet.TabIndex = 2;
            this.Button_ModelNameGet.Text = "Get";
            this.Button_ModelNameGet.UseVisualStyleBackColor = true;
            this.Button_ModelNameGet.Click += new System.EventHandler(this.Button_ModelNameGet_Click);
            // 
            // Button_ModelNameSet
            // 
            this.Button_ModelNameSet.Location = new System.Drawing.Point(7, 51);
            this.Button_ModelNameSet.Name = "Button_ModelNameSet";
            this.Button_ModelNameSet.Size = new System.Drawing.Size(75, 23);
            this.Button_ModelNameSet.TabIndex = 1;
            this.Button_ModelNameSet.Text = "Set";
            this.Button_ModelNameSet.UseVisualStyleBackColor = true;
            this.Button_ModelNameSet.Click += new System.EventHandler(this.Button_ModelNameSet_Click);
            // 
            // TextBox_ModelName
            // 
            this.TextBox_ModelName.Location = new System.Drawing.Point(7, 22);
            this.TextBox_ModelName.Name = "TextBox_ModelName";
            this.TextBox_ModelName.Size = new System.Drawing.Size(156, 22);
            this.TextBox_ModelName.TabIndex = 0;
            this.TextBox_ModelName.TextChanged += new System.EventHandler(this.TextBox_TextChanged);
            this.TextBox_ModelName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TextBox_KeyDown);
            this.TextBox_ModelName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TextBox_KeyPress);
            // 
            // tabPage_about
            // 
            this.tabPage_about.Controls.Add(this.groupBox_About);
            this.tabPage_about.Location = new System.Drawing.Point(4, 23);
            this.tabPage_about.Name = "tabPage_about";
            this.tabPage_about.Size = new System.Drawing.Size(1256, 633);
            this.tabPage_about.TabIndex = 2;
            this.tabPage_about.Text = "About";
            this.tabPage_about.UseVisualStyleBackColor = true;
            // 
            // groupBox_About
            // 
            this.groupBox_About.Controls.Add(this.lb_GUI_Version);
            this.groupBox_About.Controls.Add(this.label12);
            this.groupBox_About.Controls.Add(this.button_About);
            this.groupBox_About.Controls.Add(this.button_AboutLicense);
            this.groupBox_About.Controls.Add(this.label_about_us);
            this.groupBox_About.Controls.Add(this.label_license_agree);
            this.groupBox_About.Location = new System.Drawing.Point(8, 12);
            this.groupBox_About.Name = "groupBox_About";
            this.groupBox_About.Size = new System.Drawing.Size(235, 116);
            this.groupBox_About.TabIndex = 16;
            this.groupBox_About.TabStop = false;
            this.groupBox_About.Text = "About";
            // 
            // lb_GUI_Version
            // 
            this.lb_GUI_Version.Location = new System.Drawing.Point(156, 28);
            this.lb_GUI_Version.Name = "lb_GUI_Version";
            this.lb_GUI_Version.Size = new System.Drawing.Size(66, 14);
            this.lb_GUI_Version.TabIndex = 11;
            this.lb_GUI_Version.Text = "v1.0.0";
            this.lb_GUI_Version.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(11, 27);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(71, 14);
            this.label12.TabIndex = 10;
            this.label12.Text = "GUI Version";
            // 
            // button_About
            // 
            this.button_About.Location = new System.Drawing.Point(154, 80);
            this.button_About.Name = "button_About";
            this.button_About.Size = new System.Drawing.Size(66, 23);
            this.button_About.TabIndex = 9;
            this.button_About.Text = "Click";
            this.button_About.UseVisualStyleBackColor = true;
            this.button_About.Click += new System.EventHandler(this.button_About_Click);
            // 
            // button_AboutLicense
            // 
            this.button_AboutLicense.Location = new System.Drawing.Point(154, 51);
            this.button_AboutLicense.Name = "button_AboutLicense";
            this.button_AboutLicense.Size = new System.Drawing.Size(66, 23);
            this.button_AboutLicense.TabIndex = 8;
            this.button_AboutLicense.Text = "Click";
            this.button_AboutLicense.UseVisualStyleBackColor = true;
            this.button_AboutLicense.Click += new System.EventHandler(this.button_AboutLicense_Click);
            // 
            // label_about_us
            // 
            this.label_about_us.AutoSize = true;
            this.label_about_us.Location = new System.Drawing.Point(11, 84);
            this.label_about_us.Name = "label_about_us";
            this.label_about_us.Size = new System.Drawing.Size(56, 14);
            this.label_about_us.TabIndex = 1;
            this.label_about_us.Text = "About Us";
            // 
            // label_license_agree
            // 
            this.label_license_agree.AutoSize = true;
            this.label_license_agree.Location = new System.Drawing.Point(11, 55);
            this.label_license_agree.Name = "label_license_agree";
            this.label_license_agree.Size = new System.Drawing.Size(110, 14);
            this.label_license_agree.TabIndex = 0;
            this.label_license_agree.Text = "License Agreement";
            // 
            // label_ErrorStatus
            // 
            this.label_ErrorStatus.AutoSize = true;
            this.label_ErrorStatus.Location = new System.Drawing.Point(550, 665);
            this.label_ErrorStatus.Name = "label_ErrorStatus";
            this.label_ErrorStatus.Size = new System.Drawing.Size(0, 14);
            this.label_ErrorStatus.TabIndex = 3;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1264, 682);
            this.Controls.Add(this.label_ErrorStatus);
            this.Controls.Add(this.tabControl_MainFunctions);
            this.Controls.Add(this.statusStrip1);
            this.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MaximizeBox = false;
            this.Name = "MainWindow";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.tabControl_MainFunctions.ResumeLayout(false);
            this.tabPage_Scan.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabScanPage.ResumeLayout(false);
            this.tabPage_ScanSetting.ResumeLayout(false);
            this.GroupBox_ContScan.ResumeLayout(false);
            this.GroupBox_ContScan.PerformLayout();
            this.GroupBox_GainControl.ResumeLayout(false);
            this.GroupBox_GainControl.PerformLayout();
            this.GroupBox_SaveScan.ResumeLayout(false);
            this.GroupBox_SaveScan.PerformLayout();
            this.GroupBox_ScanAvg.ResumeLayout(false);
            this.GroupBox_ScanAvg.PerformLayout();
            this.GroupBox_LampControl.ResumeLayout(false);
            this.GroupBox_LampControl.PerformLayout();
            this.GroupBox_RefSelect.ResumeLayout(false);
            this.GroupBox_RefSelect.PerformLayout();
            this.tabPage_ScanConfig.ResumeLayout(false);
            this.tabPage_ScanConfig.PerformLayout();
            this.GroupBox_CfgDetails.ResumeLayout(false);
            this.GroupBox_CfgDetails.PerformLayout();
            this.tabPage_SaveScans.ResumeLayout(false);
            this.panel_Saved_Scan.ResumeLayout(false);
            this.panel_Saved_Scan.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_savescan)).EndInit();
            this.tabPage_Utility.ResumeLayout(false);
            this.GroupBox_BleName.ResumeLayout(false);
            this.GroupBox_BleName.PerformLayout();
            this.groupBox_Device.ResumeLayout(false);
            this.groupBox_Device.PerformLayout();
            this.groupBox_ActivationKey.ResumeLayout(false);
            this.groupBox_ActivationKey.PerformLayout();
            this.groupBox_DevInfo.ResumeLayout(false);
            this.groupBox_DevInfo.PerformLayout();
            this.GroupBox_CalibCoeffs.ResumeLayout(false);
            this.GroupBox_CalibCoeffs.PerformLayout();
            this.GroupBox_Sensors.ResumeLayout(false);
            this.GroupBox_Sensors.PerformLayout();
            this.GroupBox_DateTime.ResumeLayout(false);
            this.GroupBox_DateTime.PerformLayout();
            this.GroupBox_LampUsage.ResumeLayout(false);
            this.GroupBox_LampUsage.PerformLayout();
            this.GroupBox_DLPC150FWUpdate.ResumeLayout(false);
            this.GroupBox_DLPC150FWUpdate.PerformLayout();
            this.GroupBox_TivaFWUpdate.ResumeLayout(false);
            this.GroupBox_TivaFWUpdate.PerformLayout();
            this.GroupBox_SerialNumber.ResumeLayout(false);
            this.GroupBox_SerialNumber.PerformLayout();
            this.GroupBox_ModelName.ResumeLayout(false);
            this.GroupBox_ModelName.PerformLayout();
            this.tabPage_about.ResumeLayout(false);
            this.groupBox_About.ResumeLayout(false);
            this.groupBox_About.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatus_DeviceStatus;
        private System.Windows.Forms.TabControl tabControl_MainFunctions;
        private System.Windows.Forms.TabPage tabPage_Utility;
        private System.Windows.Forms.GroupBox GroupBox_SerialNumber;
        private System.Windows.Forms.Button Button_SerialNumberGet;
        private System.Windows.Forms.Button Button_SerialNumberSet;
        private System.Windows.Forms.TextBox TextBox_SerialNumber;
        private System.Windows.Forms.GroupBox GroupBox_ModelName;
        private System.Windows.Forms.Button Button_ModelNameGet;
        private System.Windows.Forms.Button Button_ModelNameSet;
        private System.Windows.Forms.GroupBox GroupBox_CalibCoeffs;
        private System.Windows.Forms.Button Button_CalWriteCoeffs;
        private System.Windows.Forms.Button Button_CalReadCoeffs;
        private System.Windows.Forms.Button Button_CalRestoreDefaultCoeffs;
        private System.Windows.Forms.Button Button_CalWriteGenCoeffs;
        private System.Windows.Forms.CheckBox CheckBox_CalWriteEnable;
        private System.Windows.Forms.TextBox TextBox_ShiftVectCoeff2;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.TextBox TextBox_ShiftVectCoeff1;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.TextBox TextBox_ShiftVectCoeff0;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox TextBox_P2WCoeff2;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label Label_ScanCfgVer;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.TextBox TextBox_P2WCoeff1;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Label Label_RefCalVer;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox TextBox_P2WCoeff0;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label Label_CalCoeffVer;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.GroupBox GroupBox_Sensors;
        private System.Windows.Forms.Button Button_SensorRead;
        private System.Windows.Forms.Label Label_SensorLampVM1Value;
        private System.Windows.Forms.Label Label_SensorTivaTemp;
        private System.Windows.Forms.Label Label_SensorSysTemp;
        private System.Windows.Forms.Label Label_SensorHumidity;
        private System.Windows.Forms.Label Label_SensorBattCapacity;
        private System.Windows.Forms.Label Label_SensorBattStatus;
        private System.Windows.Forms.Label Label_SensorLampVM1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label_sys_humidity;
        private System.Windows.Forms.Label lb_BattCapTitle;
        private System.Windows.Forms.Label lb_BattChargerStatusTitle;
        private System.Windows.Forms.GroupBox GroupBox_DateTime;
        private System.Windows.Forms.Button Button_DateTimeGet;
        private System.Windows.Forms.Button Button_DateTimeSync;
        private System.Windows.Forms.TextBox TextBox_DateTime;
        private System.Windows.Forms.GroupBox GroupBox_LampUsage;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button Button_LampUsageGet;
        private System.Windows.Forms.Button Button_LampUsageSet;
        private System.Windows.Forms.TextBox TextBox_LampUsage;
        private System.Windows.Forms.GroupBox GroupBox_DLPC150FWUpdate;
        private System.Windows.Forms.Button Button_DLPC150FWUpdate;
        private System.Windows.Forms.ProgressBar ProgressBar_DLPC150FWUpdateStatus;
        private System.Windows.Forms.Button Button_DLPC150FWBrowse;
        private System.Windows.Forms.TextBox TextBox_DLPC150FWPath;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox GroupBox_TivaFWUpdate;
        private System.Windows.Forms.Button Button_TivaFWUpdate;
        private System.Windows.Forms.ProgressBar ProgressBar_TivaFWUpdateStatus;
        private System.Windows.Forms.Button Button_TivaFWBrowse;
        private System.Windows.Forms.TextBox TextBox_TivaFWPath;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.GroupBox groupBox_DevInfo;
        private System.Windows.Forms.Label label_DevInfoLampUsageValue;
        private System.Windows.Forms.Label label_DevInfoLampUsage;
        private System.Windows.Forms.Label label_DevInfoUUID;
        private System.Windows.Forms.Label label114;
        private System.Windows.Forms.Label label_DevInfoManfacSerNum;
        private System.Windows.Forms.Label label_MFC_Seri_Num;
        private System.Windows.Forms.Label label_DevInfoDevSerNum;
        private System.Windows.Forms.Label label104;
        private System.Windows.Forms.Label label_DevInfoModelName;
        private System.Windows.Forms.Label label106;
        private System.Windows.Forms.Label label_DevInfoDetectorBoardVer;
        private System.Windows.Forms.Label label108;
        private System.Windows.Forms.Label label_DevInfoMainBoardVer;
        private System.Windows.Forms.Label label110;
        private System.Windows.Forms.Label label_DevInfoDLPCVer;
        private System.Windows.Forms.Label label102;
        private System.Windows.Forms.Label label_DevInfoTivaSWVer;
        private System.Windows.Forms.Label label98;
        private System.Windows.Forms.Label label_DevInfoGUIVer;
        private System.Windows.Forms.Label label95;
        private System.Windows.Forms.GroupBox groupBox_Device;
        private System.Windows.Forms.Button button_DeviceRestoreFacRef;
        private System.Windows.Forms.Label label_RestoreFacRef;
        private System.Windows.Forms.Button button_DeviceUpdateRef;
        private System.Windows.Forms.Label label119;
        private System.Windows.Forms.Button button_DeviceResetSys;
        private System.Windows.Forms.Label label118;
        private System.Windows.Forms.GroupBox groupBox_ActivationKey;
        private System.Windows.Forms.Button button_KeySet;
        private System.Windows.Forms.TextBox TextBox_Key;
        private System.Windows.Forms.Label label117;
        private System.Windows.Forms.Label label_ActivateStatus;
        private System.Windows.Forms.Label Status;
        private System.Windows.Forms.TextBox TextBox_ModelName;
        private System.Windows.Forms.TabPage tabPage_Scan;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private LiveCharts.WinForms.CartesianChart MyChart;
        private System.Windows.Forms.TabControl tabScanPage;
        private System.Windows.Forms.TabPage tabPage_ScanSetting;
        private System.Windows.Forms.GroupBox GroupBox_SaveScan;
        private System.Windows.Forms.TextBox TextBox_SaveDirPath;
        private System.Windows.Forms.Button Button_SaveDirChange;
        private System.Windows.Forms.CheckBox CheckBox_SaveRCSV;
        private System.Windows.Forms.CheckBox CheckBox_SaveACSV;
        private System.Windows.Forms.CheckBox CheckBox_SaveICSV;
        private System.Windows.Forms.CheckBox CheckBox_SaveDAT;
        private System.Windows.Forms.CheckBox CheckBox_SaveCombCSV;
        private System.Windows.Forms.GroupBox GroupBox_ScanAvg;
        private System.Windows.Forms.TextBox textBox_ScanAvg;
        private System.Windows.Forms.Label label34;
        private System.Windows.Forms.GroupBox GroupBox_LampControl;
        private System.Windows.Forms.TextBox TextBox_LampStableTime;
        private System.Windows.Forms.RadioButton RadioButton_LampStableTime;
        private System.Windows.Forms.RadioButton RadioButton_LampOff;
        private System.Windows.Forms.RadioButton RadioButton_LampOn;
        private System.Windows.Forms.GroupBox GroupBox_RefSelect;
        private System.Windows.Forms.RadioButton RadioButton_RefFac;
        private System.Windows.Forms.RadioButton RadioButton_RefPre;
        private System.Windows.Forms.RadioButton RadioButton_RefNew;
        private System.Windows.Forms.TabPage tabPage_ScanConfig;
        private System.Windows.Forms.ListBox ListBox_TargetCfgs;
        private System.Windows.Forms.Button Button_SetActive;
        private System.Windows.Forms.Label label94;
        private System.Windows.Forms.Button Button_CfgCancel;
        private System.Windows.Forms.Button Button_CfgSave;
        private System.Windows.Forms.Button Button_CfgDelete;
        private System.Windows.Forms.Button Button_CfgEdit;
        private System.Windows.Forms.Button Button_CfgNew;
        private System.Windows.Forms.GroupBox GroupBox_CfgDetails;
        private System.Windows.Forms.Label label93;
        private System.Windows.Forms.Label label92;
        private System.Windows.Forms.Label label91;
        private System.Windows.Forms.Label label90;
        private System.Windows.Forms.Label label89;
        private System.Windows.Forms.Label label88;
        private System.Windows.Forms.TextBox TextBox_CfgDigRes1;
        private System.Windows.Forms.TextBox TextBox_CfgDigRes2;
        private System.Windows.Forms.TextBox TextBox_CfgDigRes3;
        private System.Windows.Forms.TextBox TextBox_CfgDigRes4;
        private System.Windows.Forms.TextBox TextBox_CfgDigRes5;
        private System.Windows.Forms.ComboBox ComboBox_CfgExposure1;
        private System.Windows.Forms.ComboBox ComboBox_CfgExposure2;
        private System.Windows.Forms.ComboBox ComboBox_CfgExposure3;
        private System.Windows.Forms.ComboBox ComboBox_CfgExposure4;
        private System.Windows.Forms.ComboBox ComboBox_CfgExposure5;
        private System.Windows.Forms.ComboBox ComboBox_CfgWidth1;
        private System.Windows.Forms.ComboBox ComboBox_CfgWidth2;
        private System.Windows.Forms.ComboBox ComboBox_CfgWidth3;
        private System.Windows.Forms.ComboBox ComboBox_CfgWidth4;
        private System.Windows.Forms.ComboBox ComboBox_CfgWidth5;
        private System.Windows.Forms.TextBox TextBox_CfgRangeEnd1;
        private System.Windows.Forms.TextBox TextBox_CfgRangeEnd2;
        private System.Windows.Forms.TextBox TextBox_CfgRangeEnd3;
        private System.Windows.Forms.TextBox TextBox_CfgRangeEnd4;
        private System.Windows.Forms.TextBox TextBox_CfgRangeEnd5;
        private System.Windows.Forms.TextBox TextBox_CfgRangeStart1;
        private System.Windows.Forms.TextBox TextBox_CfgRangeStart2;
        private System.Windows.Forms.TextBox TextBox_CfgRangeStart3;
        private System.Windows.Forms.TextBox TextBox_CfgRangeStart4;
        private System.Windows.Forms.ComboBox ComboBox_CfgScanType1;
        private System.Windows.Forms.ComboBox ComboBox_CfgScanType2;
        private System.Windows.Forms.ComboBox ComboBox_CfgScanType3;
        private System.Windows.Forms.ComboBox ComboBox_CfgScanType4;
        private System.Windows.Forms.TextBox TextBox_CfgRangeStart5;
        private System.Windows.Forms.ComboBox ComboBox_CfgScanType5;
        private System.Windows.Forms.Label label87;
        private System.Windows.Forms.Label label86;
        private System.Windows.Forms.Label label85;
        private System.Windows.Forms.Label label84;
        private System.Windows.Forms.Label label83;
        private System.Windows.Forms.Label label82;
        private System.Windows.Forms.TextBox TextBox_CfgAvg;
        private System.Windows.Forms.Label label81;
        private System.Windows.Forms.TextBox TextBox_CfgName;
        private System.Windows.Forms.Label label80;
        private System.Windows.Forms.RadioButton RadioButton_Reference;
        private System.Windows.Forms.RadioButton RadioButton_Intensity;
        private System.Windows.Forms.RadioButton RadioButton_Absorbance;
        private System.Windows.Forms.RadioButton RadioButton_Reflectance;
        private System.Windows.Forms.Label Label_EstimatedScanTime;
        private System.Windows.Forms.Label Label_CurrentConfig;
        private System.Windows.Forms.Label Label_ScanStatus;
        private System.Windows.Forms.Button Button_Scan;
        private System.Windows.Forms.GroupBox GroupBox_GainControl;
        private System.Windows.Forms.CheckBox CheckBox_AutoGain;
        private System.Windows.Forms.ComboBox ComboBox_PGAGain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox CheckBox_LampOn;
        private System.Windows.Forms.Label Label_CfgNumPatterns;
        private System.Windows.Forms.TabPage tabPage_about;
        private System.Windows.Forms.GroupBox groupBox_About;
        private System.Windows.Forms.Button button_About;
        private System.Windows.Forms.Button button_AboutLicense;
        private System.Windows.Forms.Label label_about_us;
        private System.Windows.Forms.Label label_license_agree;
        private System.Windows.Forms.Label label_ActiveConfig;
        private ComboBox comboBox_cfgNumSec;
        private Label label_ErrorStatus;
        private TabPage tabPage_SaveScans;
        private Panel panel_Saved_Scan;
        private Label label33;
        private Label label30;
        private Label Label_SavedAvg;
        private Label label77;
        private Label label76;
        private Label label75;
        private Label label74;
        private Label label73;
        private Label Label_SavedExposure1;
        private Label Label_SavedExposure2;
        private Label Label_SavedExposure3;
        private Label Label_SavedExposure4;
        private Label Label_SavedExposure5;
        private Label label67;
        private Label Label_SavedDigRes1;
        private Label Label_SavedDigRes2;
        private Label Label_SavedDigRes3;
        private Label Label_SavedDigRes4;
        private Label Label_SavedDigRes5;
        private Label label61;
        private Label Label_SavedWidth1;
        private Label Label_SavedWidth2;
        private Label Label_SavedWidth3;
        private Label Label_SavedWidth4;
        private Label Label_SavedWidth5;
        private Label label55;
        private Label Label_SavedRangeEnd1;
        private Label Label_SavedRangeEnd2;
        private Label Label_SavedRangeEnd3;
        private Label Label_SavedRangeEnd4;
        private Label Label_SavedRangeEnd5;
        private Label label49;
        private Label Label_SavedRangeStart1;
        private Label Label_SavedRangeStart2;
        private Label Label_SavedRangeStart3;
        private Label Label_SavedRangeStart4;
        private Label Label_SavedRangeStart5;
        private Label label43;
        private Label Label_SavedScanType1;
        private Label Label_SavedScanType2;
        private Label Label_SavedScanType3;
        private Label Label_SavedScanType4;
        private Label Label_SavedScanType5;
        private Label label37;
        private Label label36;
        private Label label35;
        private Button Button_DisplayDirChange;
        private Label label79;
        private CheckBox Check_Overlay;
        private GroupBox GroupBox_ContScan;
        private Label label3;
        private TextBox Text_ContDelay;
        private Label label2;
        private TextBox Text_ContScan;
        private Label Label_ContScan;
        private Button Button_MoveCfgT2L;
        private Button Button_MoveCfgL2T;
        private Button Button_CopyCfgT2L;
        private Button Button_CopyCfgL2T;
        private ListBox ListBox_LocalCfgs;
        private TextBox TextBox_FileNamePrefix1;
        private CheckBox CheckBox_FileNamePrefix;
        private CheckBox CheckBox_SaveRJDX;
        private CheckBox CheckBox_SaveAJDX;
        private CheckBox CheckBox_SaveIJDX;
        private CheckBox CheckBox_SaveOneCSV;
        private Button Button_ClearAllErrors;
        private Label label_ref;
        private DataGridView dataGridView_savescan;
        private TextBox textBox_filter;
        private Label label13;
        private Button button_clear;
        private CheckBox checkBox_zoom;
        private CheckBox checkBox_tooltip;
        private Label label_pattern5;
        private Label label_pattern4;
        private Label label_pattern3;
        private Label label_pattern2;
        private Label label_pattern1;
        private Label label14;
        private Label Label_SensorLampCM2;
        private Label Label_SensorLampCM1;
        private Label Label_SensorLampVM2;
        private Label Label_SensorLampCM2Value;
        private Label Label_SensorLampCM1Value;
        private Label Label_SensorLampVM2Value;
        private GroupBox GroupBox_BleName;
        private Button Button_Get_BLE_Display_Name;
        private Button Button_Set_BLE_Display_Name;
        private Button Button_Clear_BLE_Display_Name;
        private TextBox TextBox_BLE_Display_Name;
        private Button Button_LockButton;
        private Button Button_UnlockButton;
        private Label Label_ButtonStatus;
        private Button button_restore_fac_ref_warning;
        private Label label_maxPattern5;
        private Label label_maxPattern4;
        private Label label_maxPattern3;
        private Label label_maxPattern2;
        private Label label_maxPattern1;
        private Label label_totalPatterns;
        private Label label7;
        private CheckBox checkBox_StopOnError;
        private Label label11;
        private Label label10;
        private TextBox TextBox_FileNamePrefix3;
        private TextBox TextBox_FileNamePrefix2;
        private ToolTip toolTip1;
        private TextBox TextBox_SavedFileDirPath;
        private Label Label_BleNameValue;
        private Label Label_Blename;
        private Label lb_GUI_Version;
        private Label label12;
        private Label label16;
        private Label label15;
    }
}

