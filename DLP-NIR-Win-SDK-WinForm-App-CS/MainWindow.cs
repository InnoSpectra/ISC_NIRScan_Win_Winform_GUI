﻿using DLP_NIR_Win_SDK_CS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.IO;
using System.Xml;
using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using LiveCharts.Configurations;
using System.Threading;
using System.Timers;
using LiveCharts.Geared;
using System.Xml.Serialization;
using System.Windows.Threading;
using System.Text.RegularExpressions;

namespace DLP_NIR_Win_SDK_WinForm_App_CS
{
    public partial class MainWindow : Form
    {
        // For Configuration
        private const Int32 MIN_WAVELENGTH = 900;
        private const Int32 MAX_WAVELENGTH = 1700;
        private const Int32 MAX_CFG_SECTION = 5;

        private List<ScanConfig.SlewScanConfig> LocalConfig = new List<ScanConfig.SlewScanConfig>();
        private List<ComboBox> ComboBox_CfgScanType = new List<ComboBox>();
        private List<ComboBox> ComboBox_CfgWidth = new List<ComboBox>();
        private List<ComboBox> ComboBox_CfgExposure = new List<ComboBox>();
        private List<TextBox> TextBox_CfgRangeStart = new List<TextBox>();
        private List<TextBox> TextBox_CfgRangeEnd = new List<TextBox>();
        private List<TextBox> TextBox_CfgDigRes = new List<TextBox>();
        private List<Label> Label_Pattern = new List<Label>();

        private Int32 TargetCfg_SelIndex = -1;       // Rocord device selected config
        private Int32 TargetCfg_Last_SelIndex = -1;  // Rocord last device selected config
        private Boolean SelCfg_IsTarget = false;     // Record target config or local config
        private Int32 LocalCfg_SelIndex = -1;        // Record local selected config
        private Int32 LocalCfg_Last_SelIndex = -1;   // Rocord last local selected config
        Boolean NewConfig = false;                   // Record new config or existed config
        Boolean EditConfig = false;
        private Int32 DevCurCfg_Index = -1;          // Record current config which set to device
        int EditSelectIndex = -1;                    // Record edit select index   

        private BackgroundWorker bwDLPCUpdate;
        private BackgroundWorker bwTivaUpdate;

        private String Display_Dir = String.Empty;
        private String Scan_Dir = String.Empty;
        public static String ConfigDir = String.Empty;
        private Int32 ScanFile_Formats = 0;

        // Saved Scans
        List<Label> Label_SavedScanType = new List<Label>();
        List<Label> Label_SavedRangeStart = new List<Label>();
        List<Label> Label_SavedRangeEnd = new List<Label>();
        List<Label> Label_SavedWidth = new List<Label>();
        List<Label> Label_SavedDigRes = new List<Label>();
        List<Label> Label_SavedExposure = new List<Label>();

        public bool IsActivated { get { if (Device.GetActivationResult() == 1) return true; else return false; } }

        private BackgroundWorker bwRefScanProgress;
        private String Tiva_FWDir = String.Empty;
        private String DLPC_FWDir = String.Empty;

        // For Scan
        private DateTime TimeScanStart = new DateTime();
        private DateTime TimeScanEnd = new DateTime();
        private static UInt32 LampStableTime = 625;
        private Scan.SCAN_REF_TYPE ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_NEW;
        private BackgroundWorker bwScan;

        private Boolean DevCurCfg_IsTarget = false; // Record current config is device or local
        private String pre_ref_time = "";
        public static String buildin_ref_time = "";
        private Boolean isCancelConfig = false;
        private Boolean isSelectConfig = false;
        private List<String> ListFilesTime = new List<String>();

        public static bool UserCancelScan = false;
        private int TargetScanCounts = 0;
        private int ScannedCounts = 0;
        private bool SaveOneCSVFile = false;
        private String OneScanFileName = String.Empty;

        public enum FW_LEVEL
        {
            LEVEL_0,
            LEVEL_1,
            LEVEL_2,
            LEVEL_3,
            LEVEL_4,
        };
        public enum ScanConfigMode
        {
            INITIAL,
            NEW,
            EDIT,
            DELETE,
            SAVE,
            CANCEL,
        };

        public MainWindow()
        {
            InitializeComponent();
            label11.Text = "";
            label12.Text = "";
            
            label_ref.Visible = false;
            Label_ContScan.Text = "";
            this.FormClosing += Form1_FormClosing;
            CheckForIllegalCrossThreadCalls = false;//solve across thread is invalid
            MainWindow_Loaded();
            initBackgroundWorker();
            initChart();
            String version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            this.Text = string.Format("ISC NIRScan GUI SDK WinForm v{0}", version.Substring(0, version.LastIndexOf('.')));
            UI_no_connection();
            Label_CurrentConfig.ForeColor = System.Drawing.Color.OrangeRed;
            //load save scan 
            LoadSavedScanList(false);
            CheckScanDirPath();
            // Enable the CPP DLL debug output for development
            DBG.Enable_CPP_Console();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            String HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;
            if ((GetFW_LEVEL() >= FW_LEVEL.LEVEL_2 && Device.ChkBleExist() == 1) || HWRev == String.Empty)
                Device.SetBluetooth(true);
            SaveSettings();
            Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);
        }
        private void initBackgroundWorker()
        {
            bwDLPCUpdate = new BackgroundWorker();
            bwTivaUpdate = new BackgroundWorker();
            bwDLPCUpdate.WorkerReportsProgress = true;
            bwTivaUpdate.WorkerReportsProgress = true;
            bwDLPCUpdate.WorkerSupportsCancellation = true;
            bwTivaUpdate.WorkerSupportsCancellation = true;
            bwDLPCUpdate.DoWork += new DoWorkEventHandler(bwDLPCUpdate_DoWork);
            bwTivaUpdate.DoWork += new DoWorkEventHandler(bwTivaUpdate_DoWork);
            bwDLPCUpdate.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwDLPCUpdate_DoWorkCompleted);
            bwTivaUpdate.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwTivaUpdate_DoSacnCompleted);
            bwDLPCUpdate.ProgressChanged += new ProgressChangedEventHandler(bwDLPCUpdate_ProgressChanged);
            bwTivaUpdate.ProgressChanged += new ProgressChangedEventHandler(bwTivaUpdate_ProgressChanged);

            bwScan = new BackgroundWorker
            {
                WorkerReportsProgress = false,
                WorkerSupportsCancellation = true
            };
            bwScan.DoWork += new DoWorkEventHandler(bwScan_DoScan);
            bwScan.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwScan_DoSacnCompleted);
        }
        private void MainWindow_Loaded()
        {
            InitGUIItem();
            Device.Init();
            InitSavedScanCfgItems();

            SDK.OnDeviceConnectionLost += new Action<bool>(Device_Disconncted_Handler);
            SDK.OnDeviceConnected += new Action<string>(Device_Connected_Handler);
            SDK.OnDeviceFound += new Action(Device_Found_Handler);
            SDK.OnDeviceError += new Action<string>(Device_Error_Handler);
            SDK.OnButtonScan += new Action(StartButtonScan);
            SDK.OnErrorStatusFound += new Action(RefreshErrorStatus);
            SDK.OnBeginConnectingDevice += new Action(Connecting_Device);
            SDK.OnBeginScan += new Action(BeginScan);
            SDK.OnScanCompleted += new Action(ScanCompleted);

            if (!Device.IsConnected())
            {
                SDK.AutoSearch = true;
            }
           
            LoadConfigDir();
            LoadScanPageSetting();
            TextBox_SaveDirPath.Text = Scan_Dir;
            TextBox_DisplayDirPath.Text = Display_Dir;
        }
        private void InitGUIItem()
        {
            //init GUI item
            toolStripStatus_DeviceStatus.Image = Properties.Resources.Led_Gray;
            RadioButton_RefNew.Checked = true;
            ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_NEW;
            Button_Scan.Text = "Reference Scan";
            ComboBox_PGAGain.SelectedItem = "64";
            ComboBox_PGAGain.Enabled = false;
            CheckBox_AutoGain.Checked = true;

            //SaveScan DataGridView
            DataGridViewTextBoxColumn FileName_col = new DataGridViewTextBoxColumn
            {
                HeaderText = "FileName",
                Name = "FileName"
            };
            dataGridView_savescan.Columns.Add(FileName_col);
            dataGridView_savescan.Columns[0].ReadOnly = true;
            FileName_col.Width = 180;

            DataGridViewTextBoxColumn Time_col = new DataGridViewTextBoxColumn
            {
                HeaderText = "Time",
                Name = "Time"
            };
            dataGridView_savescan.Columns.Add(Time_col);
            dataGridView_savescan.Columns[1].ReadOnly = true;
            Time_col.Width = 140;
        }
        #region Error 
        private void RefreshErrorStatus()
        {
            String ErrMsg = String.Empty;

            if (Device.ReadErrorStatusAndCode() != 0)
                return;

            if ((Device.ErrStatus & 0x00000001) == 0x00000001)  // Scan Error
            {
                if (Device.ErrCode[0] == 0x00000001)
                    ErrMsg += "Scan Error: DLPC150 Boot Error Detected.    ";
                else if (Device.ErrCode[0] == 0x00000002)
                    ErrMsg += "Scan Error: DLPC150 Init Error Detected.    ";
                else if (Device.ErrCode[0] == 0x00000004)
                    ErrMsg += "Scan Error: DLPC150 Lamp Driver Rrror Detected.    ";
                else if (Device.ErrCode[0] == 0x00000008)
                    ErrMsg += "Scan Error: DLPC150 Crop Image Failed.    ";
                else if (Device.ErrCode[0] == 0x00000010)
                    ErrMsg += "Scan Error: ADC Data Error.    ";
                else if (Device.ErrCode[0] == 0x00000020)
                    ErrMsg += "Scan Error: Scan Config Invalid.    ";
                else if (Device.ErrCode[0] == 0x00000040)
                    ErrMsg += "Scan Error: Scan Pattern Streaming Error.    ";
                else if (Device.ErrCode[0] == 0x00000080)
                    ErrMsg += "Scan Error: DLPC150 Read Error.    ";
            }

            if ((Device.ErrStatus & 0x00000002) == 0x00000002)  // ADC Error
            {
                if (Device.ErrCode[1] == 0x00000001)
                    ErrMsg += "ADC Error: Timeout Error.    ";
                else if (Device.ErrCode[1] == 0x00000002)
                    ErrMsg += "ADC Error: PowerDown Error.    ";
                else if (Device.ErrCode[1] == 0x00000003)
                    ErrMsg += "ADC Error: PowerUp Error.    ";
                else if (Device.ErrCode[1] == 0x00000004)
                    ErrMsg += "ADC Error: Standby Error.    ";
                else if (Device.ErrCode[1] == 0x00000005)
                    ErrMsg += "ADC Error: WakeUp Error.    ";
                else if (Device.ErrCode[1] == 0x00000006)
                    ErrMsg += "ADC Error: Read Register Error.    ";
                else if (Device.ErrCode[1] == 0x00000007)
                    ErrMsg += "ADC Error: Write Register Error.    ";
                else if (Device.ErrCode[1] == 0x00000008)
                    ErrMsg += "ADC Error: Configure Error.    ";
                else if (Device.ErrCode[1] == 0x00000009)
                    ErrMsg += "ADC Error: Set Buffer Error.    ";
                else if (Device.ErrCode[1] == 0x0000000A)
                    ErrMsg += "ADC Error: Command Error.    ";
                else if (Device.ErrCode[1] == 0x0000000B)
                    ErrMsg += "ADC Error: Set PGA Error.    ";
            }

            if ((Device.ErrStatus & 0x00000004) == 0x00000004)  // SD Card Error
            {
                ErrMsg += "SD Card Error.    ";
            }

            if ((Device.ErrStatus & 0x00000008) == 0x00000008)  // EEPROM Error
            {
                ErrMsg += "EEPROM Error.    ";
            }

            if ((Device.ErrStatus & 0x00000010) == 0x00000010)  // BLE Error
            {
                ErrMsg += "Bluetooth Error.    ";
            }

            if ((Device.ErrStatus & 0x00000020) == 0x00000020)  // Spectrum Library Error
            {
                ErrMsg += "Spectrum Library Error.    ";
            }

            if ((Device.ErrStatus & 0x00000040) == 0x00000040)  // Hardware Error
            {
                if (Device.ErrCode[6] == 0x00000001)
                    ErrMsg += "HW Error: DLPC150 Error.    ";
                else if (Device.ErrCode[6] == 0x00000002)
                    ErrMsg += "HW Error: Read UUID Error.    ";
                else if (Device.ErrCode[6] == 0x00000003)
                    ErrMsg += "HW Error: Flash Initial Error.    ";
            }

            if ((Device.ErrStatus & 0x00000080) == 0x00000080)  // TMP Sensor Error
            {
                if (GetFW_LEVEL() == FW_LEVEL.LEVEL_1)
                {
                    // Reset error status because TMP sensor phased out, but older Tiva FW still exist this error status.
                    Device.ResetErrorStatus(0x00000080);
                    RefreshErrorStatus();
                }
                else
                {
                    if (Device.ErrCode[7] == 0x00000001)
                        ErrMsg += "TMP Error: Invalid Manufacturing ID.    ";
                    else if (Device.ErrCode[7] == 0x00000002)
                        ErrMsg += "TMP Error: Invalid Device ID.    ";
                    else if (Device.ErrCode[7] == 0x00000003)
                        ErrMsg += "TMP Error: Reset Error.    ";
                    else if (Device.ErrCode[7] == 0x00000004)
                        ErrMsg += "TMP Error: Read Register Error.    ";
                    else if (Device.ErrCode[7] == 0x00000005)
                        ErrMsg += "TMP Error: Write Register Error.    ";
                    else if (Device.ErrCode[7] == 0x00000006)
                        ErrMsg += "TMP Error: Timeout Error.    ";
                    else if (Device.ErrCode[7] == 0x00000007)
                        ErrMsg += "TMP Error: I2C Error.    ";
                }
            }

            if ((Device.ErrStatus & 0x00000100) == 0x00000100)  // HDC Sensor Error
            {
                if (Device.ErrCode[8] == 0x00000001)
                    ErrMsg += "HDC Error: Invalid Manufacturing ID.    ";
                else if (Device.ErrCode[8] == 0x00000002)
                    ErrMsg += "HDC Error: Invalid Device ID.    ";
                else if (Device.ErrCode[8] == 0x00000003)
                    ErrMsg += "HDC Error: Reset Error.    ";
                else if (Device.ErrCode[8] == 0x00000004)
                    ErrMsg += "HDC Error: Read Register Error.    ";
                else if (Device.ErrCode[8] == 0x00000005)
                    ErrMsg += "HDC Error: Write Register Error.    ";
                else if (Device.ErrCode[8] == 0x00000006)
                    ErrMsg += "HDC Error: Timeout Error.    ";
                else if (Device.ErrCode[8] == 0x00000007)
                    ErrMsg += "HDC Error: I2C Error.    ";
            }

            if ((Device.ErrStatus & 0x00000200) == 0x00000200)  // Battery Error
            {
                if (Device.ErrCode[9] == 0x00000001)
                    ErrMsg += "Battery Error: Battery Low.    ";
            }

            if ((Device.ErrStatus & 0x00000400) == 0x00000400)  // Insufficient Memory Error
            {
                ErrMsg += "Not Enough Memory.    ";
            }

            if ((Device.ErrStatus & 0x00000800) == 0x00000800)  // UART Error
            {
                ErrMsg += "UART Error.    ";
            }

            if ((Device.ErrStatus & 0x00001000) == 0x00001000)  // System Error
            {
                if (Device.ErrCode[12] == 0x00000001)
                    ErrMsg += "System Error: Unstable Lamp ADC.    ";
                else if (Device.ErrCode[12] == 0x00000002)
                    ErrMsg += "System Error: Unstable Peak Intensity.    ";
                else if (Device.ErrCode[12] == 0x00000003)
                    ErrMsg += "System Error: ADS1255 Error.    ";
                else if (Device.ErrCode[12] == 0x00000004)
                    ErrMsg += "System Error: Auto PGA Error.    ";
            }

            label_ErrorStatus.Text = ErrMsg;
            label_ErrorStatus.ForeColor = System.Drawing.Color.Red;
        }

        private void Device_Error_Handler(string error)
        {
            if(GetFW_LEVEL() >= FW_LEVEL.LEVEL_2)
            {
                if(error .Contains("Scan Configuration List read failed") && ScanConfig.TargetConfig.Count == 0)
                {
                    Message.ShowInfo("Scan config is empty. Will set default config to device.");
                    SetDefaultConfig();
                }
                else
                {
                    Message.ShowWarning(error);  // Device Information, Calibration Coefficients, Configuration Lists
                }                
            }
                
        }

        public  void ShowWarning(String Text)
        {
            String text = Text;
            MessageBox.Show(text, "Warning");
        }
        #endregion

        #region connect device
        private void Device_Disconncted_Handler(bool error)
        {
            ListBox_LocalCfgs.Items.Clear();
            ListBox_TargetCfgs.Items.Clear();
            toolStripStatus_DeviceStatus.Image = Properties.Resources.Led_R;
            toolStripStatus_DeviceStatus.Text = "Device disconnect!";
            ClearScanPlotsUI();
            if (error)
            {
                DBG.WriteLine("Device disconnected abnormally !");
                SDK.AutoSearch = true;
            }
            else
                DBG.WriteLine("Device disconnected successfully !");
            UI_no_connection();
        }
        private void Device_Found_Handler()
        {
            SDK.AutoSearch = false;
            Enumerate_Devices();
        }
        private void Enumerate_Devices()
        {
            Device.Enumerate();
            Device.Open(null);
        }
        private void Connecting_Device()
        {
        }
        private void Device_Connected_Handler(String SerialNumber)
        {
            label11.Text = "";
            label12.Text = "";
            if (SerialNumber == null)
                DBG.WriteLine("Device connecting failed !");
            else
            {
                DBG.WriteLine("Device <{0}> connected successfullly !", SerialNumber);
            }
            int count = 0;
            while (!Device.IsConnected())//wait device connect 
            {
                count++;
                if(count>30)//avoid loop
                {
                    Message.ShowError("Can not connect to device. Please reopen the device or software.");
                    return;
                }
                Thread.Sleep(200);
            }
            String HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;
            if (GetFW_LEVEL() == FW_LEVEL.LEVEL_1)
            {
                Message.ShowWarning("The version is too old.\nPlease update your TIVA FW.");
            }           
            else if (GetFW_LEVEL() == FW_LEVEL.LEVEL_0)
            {
                if(HWRev == "E")
                {
                    Message.ShowWarning("This GUI does not support device with extended wavelengths. Please use the standard wavelength device.");
                }
                else
                {
                    Message.ShowWarning("The device is not ISC product. Functions may be abnormal!");
                }             
            }

            
            if ((GetFW_LEVEL() >= FW_LEVEL.LEVEL_2 && Device.ChkBleExist() == 1) || HWRev == String.Empty)
                Device.SetBluetooth(false);
            if ((int)GetFW_LEVEL() > 0 && HWRev != String.Empty)
                CheckFactoryRefData();

            // Sync device date time
            DateTime Current = DateTime.Now;
            Device.DeviceDateTime DevDateTime = new Device.DeviceDateTime
            {
                Year = Current.Year,
                Month = Current.Month,
                Day = Current.Day,
                DayOfWeek = (Int32)Current.DayOfWeek,
                Hour = Current.Hour,
                Minute = Current.Minute,
                Second = Current.Second
            };
            Device.SetDateTime(DevDateTime);
       
            RefreshErrorStatus();
            GetDeviceInfo();
            // Scan Config
            LoadLocalCfgList();
            PopulateCfgDetailItems();
            RefreshTargetCfgList();  // Only refresh UI because target config list has been loaded after device opened
            
            GetActivationKeyStatus();
            EnableCfgItem(false);

            // Scan Plot Area
            Int32 ActiveIndex = ScanConfig.GetTargetActiveScanIndex();
            ListBox_TargetCfgs.SelectedIndex = ActiveIndex;
            if (ActiveIndex >= 0)
            {
                SetScanConfig(ScanConfig.TargetConfig[ActiveIndex],true, ActiveIndex);
            }

            if (Scan.IsLocalRefExist)
                RadioButton_RefPre.Enabled = true;
            else
                RadioButton_RefPre.Enabled = false;

            // Scan Setting
            if(ActiveIndex<0)
            {
                return;
            }

            RadioButton_RefNew.Checked = true;
            //TextBox_LampStableTime.Text = LampStableTime.ToString();
            TextBox_LampStableTime.Text = "625";
            CheckBox_AutoGain.Checked = true;
            CheckBox_AutoGain.Enabled = true;
            ComboBox_PGAGain.Enabled = false;
              
            //Device status
            toolStripStatus_DeviceStatus.Image = Properties.Resources.Led_G;
            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_2 && !IsActivated) 
            {
                toolStripStatus_DeviceStatus.Text = "Device " + Device.DevInfo.ModelName + " (" + Device.DevInfo.SerialNumber + ") connected but advanced functions locked!";
            }
            else
            {
                toolStripStatus_DeviceStatus.Text = "Device " + Device.DevInfo.ModelName + " (" + Device.DevInfo.SerialNumber + ") connected!";
            }
            OpenColseScanConfigButton(nameof(ScanConfigMode.INITIAL));
            //get build-in ref time
            GetBuildInRefTime();
            //get ref time
            if (Scan.IsLocalRefExist)
            {
                Scan.GetRefTime(Scan.SCAN_REF_TYPE.SCAN_REF_PREV);
                Byte[] time = Scan.ReferenceScanDateTime;
                if (time[0] != 0)
                {
                    pre_ref_time = "Previous reference last set on : 20" + time[0].ToString() + "/" + time[1].ToString() + "/" + time[2].ToString()
                    + " @ " + time[3].ToString() + ":" + time[4].ToString() + ":" + time[5].ToString();
                }
            }
            UI_on_connection();//已經連線，會開啟GUI使用
            EnableCfgItem(false);
            if(!CheckBox_CalWriteEnable.Checked)
            {
                //utility item
                Button_CalWriteCoeffs.Enabled = false;
                Button_CalWriteGenCoeffs.Enabled = false;
                Button_CalRestoreDefaultCoeffs.Enabled = false;
            }
            else
            {
                //utility item
                Button_CalWriteCoeffs.Enabled = true;
                Button_CalWriteGenCoeffs.Enabled = true;
                if(GetFW_LEVEL() >=FW_LEVEL.LEVEL_2 && IsActivated)
                    Button_CalRestoreDefaultCoeffs.Enabled = true;
                else
                    Button_CalRestoreDefaultCoeffs.Enabled = false;
            }       
            if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2)
            {
                groupBox_ActivationKey.Enabled = false;
                button_DeviceBackUpFacRef.Enabled = false;
                button_DeviceRestoreFacRef.Enabled = false;
            }
            else
            {
                groupBox_ActivationKey.Enabled = true;
                //check Factory reference can backup and restore or not 
                if(label11.Text.Contains("factory reference data"))
                {
                    button_DeviceBackUpFacRef.Enabled = false;
                    button_DeviceBackUpFacRef.Text = "N/A";
                    label120.Enabled = false;
                }
                else
                {
                    button_DeviceBackUpFacRef.Enabled = true;
                    button_DeviceBackUpFacRef.Text = "BackUp";
                    label120.Enabled = true;
                }
                if(!CheckFactoryRefData())
                {
                    button_DeviceRestoreFacRef.Enabled = false;
                    label121.Enabled = false;
                    button_DeviceRestoreFacRef.Text = "N/A";
                    label12.Text = "(Not have factory reference in local!)";
                    label12.ForeColor = System.Drawing.Color.Red;
                }
                else
                {
                    button_DeviceRestoreFacRef.Enabled = true;
                    label121.Enabled = true;
                    button_DeviceRestoreFacRef.Text = "Restore";
                    label12.Text = "";
                }               
            }

            if (GetFW_LEVEL() < FW_LEVEL.LEVEL_1)
            {
                GroupBox_ModelName.Enabled = false;
            }
            else
            {
                GroupBox_ModelName.Enabled = true;
            }
            if (GetFW_LEVEL() == FW_LEVEL.LEVEL_1)
            {
                label116.Enabled = false;
            }
            else
            {
                label116.Enabled = true;
            }

            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_4)
            {
                Label_ButtonStatus.Visible = true;
                Button_LockButton.Visible = true;
                Button_UnlockButton.Visible = true;
                GroupBox_BleName.Visible = true;

                if (IsActivated)
                {
                    Label_ButtonStatus.Enabled = true;
                    Button_LockButton.Enabled = true;
                    Button_UnlockButton.Enabled = true;
                    GroupBox_BleName.Enabled = true;
                }
                else
                {
                    Label_ButtonStatus.Enabled = false;
                    Button_LockButton.Enabled = false;
                    Button_UnlockButton.Enabled = false;
                    GroupBox_BleName.Enabled = false;
                }

                Int32 status = Device.GetButtonLockStatus();
                if (status == 1)
                    Label_ButtonStatus.Text = "Button Status: Locked!";
                else if (status == 0)
                    Label_ButtonStatus.Text = "Button Status: Unlocked!";
                else
                    Label_ButtonStatus.Text = "Button Status: Read Failed!";
            }
            else
            {
                Label_ButtonStatus.Visible = false;
                Button_LockButton.Visible = false;
                Button_UnlockButton.Visible = false;
                GroupBox_BleName.Visible = false;
            }

            //init scan config list UI 
            ListBox_LocalCfgs.BackColor = System.Drawing.Color.White;
            ListBox_TargetCfgs.BackColor = System.Drawing.Color.AliceBlue;
            Button_SetActive.Enabled = true;
        }
        private void GetBuildInRefTime()
        {
            Scan.GetRefTime(Scan.SCAN_REF_TYPE.SCAN_REF_BUILT_IN);
            Byte[] buildintime = Scan.ReferenceScanDateTime;
            String refname = Scan.ReferenceScanConfigData.head.config_name;
            if (buildintime[0] != 0)
            {
                if (refname == "SystemTest")
                {
                    buildin_ref_time = "Factory Reference : 20" + buildintime[0].ToString() + "/" + buildintime[1].ToString() + "/" + buildintime[2].ToString()
                            + " @ " + buildintime[3].ToString() + ":" + buildintime[4].ToString() + ":" + buildintime[5].ToString();
                }
                else
                {
                    buildin_ref_time = "User Reference : 20" + buildintime[0].ToString() + "/" + buildintime[1].ToString() + "/" + buildintime[2].ToString()
                            + " @ " + buildintime[3].ToString() + ":" + buildintime[4].ToString() + ":" + buildintime[5].ToString();
                }
            }
        }
        private Boolean CheckFactoryRefData()
        {
            Boolean isExist = false;
            String FacRefFile = Device.DevInfo.SerialNumber + "_FacRef.dat";
            String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            String FilePath = Path.Combine(path, "InnoSpectra\\Reference Data", FacRefFile);

            if (!File.Exists(FilePath))
            {
                DeviceConnectBackUpRef();
                isExist = false;
            }
            else
            {
                isExist = true;
            }
            return isExist;
        }
        private void StartButtonScan()
        { 
            BeginInvoke((Action)(() => //Invoke at UI thread
            { //run in UI thread
                RadioButton_RefFac.Checked = true;
                Button_Scan_Click(null, null);
            }), null);
        }
        #endregion

        #region init
        private void InitSavedScanCfgItems()
        {
            Label_SavedScanType.Clear();
            Label_SavedScanType.Add(Label_SavedScanType1);
            Label_SavedScanType.Add(Label_SavedScanType2);
            Label_SavedScanType.Add(Label_SavedScanType3);
            Label_SavedScanType.Add(Label_SavedScanType4);
            Label_SavedScanType.Add(Label_SavedScanType5);
            Label_SavedRangeStart.Clear();
            Label_SavedRangeStart.Add(Label_SavedRangeStart1);
            Label_SavedRangeStart.Add(Label_SavedRangeStart2);
            Label_SavedRangeStart.Add(Label_SavedRangeStart3);
            Label_SavedRangeStart.Add(Label_SavedRangeStart4);
            Label_SavedRangeStart.Add(Label_SavedRangeStart5);
            Label_SavedRangeEnd.Clear();
            Label_SavedRangeEnd.Add(Label_SavedRangeEnd1);
            Label_SavedRangeEnd.Add(Label_SavedRangeEnd2);
            Label_SavedRangeEnd.Add(Label_SavedRangeEnd3);
            Label_SavedRangeEnd.Add(Label_SavedRangeEnd4);
            Label_SavedRangeEnd.Add(Label_SavedRangeEnd5);
            Label_SavedWidth.Clear();
            Label_SavedWidth.Add(Label_SavedWidth1);
            Label_SavedWidth.Add(Label_SavedWidth2);
            Label_SavedWidth.Add(Label_SavedWidth3);
            Label_SavedWidth.Add(Label_SavedWidth4);
            Label_SavedWidth.Add(Label_SavedWidth5);
            Label_SavedDigRes.Clear();
            Label_SavedDigRes.Add(Label_SavedDigRes1);
            Label_SavedDigRes.Add(Label_SavedDigRes2);
            Label_SavedDigRes.Add(Label_SavedDigRes3);
            Label_SavedDigRes.Add(Label_SavedDigRes4);
            Label_SavedDigRes.Add(Label_SavedDigRes5);
            Label_SavedExposure.Clear();
            Label_SavedExposure.Add(Label_SavedExposure1);
            Label_SavedExposure.Add(Label_SavedExposure2);
            Label_SavedExposure.Add(Label_SavedExposure3);
            Label_SavedExposure.Add(Label_SavedExposure4);
            Label_SavedExposure.Add(Label_SavedExposure5);

            ClearSavedScanCfgItems();
        }

        private void PopulateCfgDetailItems()
        {
            ComboBox_CfgScanType.Clear();
            ComboBox_CfgScanType1.Items.Clear();
            ComboBox_CfgScanType2.Items.Clear();
            ComboBox_CfgScanType3.Items.Clear();
            ComboBox_CfgScanType4.Items.Clear();
            ComboBox_CfgScanType5.Items.Clear();
            ComboBox_CfgScanType.Add(ComboBox_CfgScanType1);
            ComboBox_CfgScanType.Add(ComboBox_CfgScanType2);
            ComboBox_CfgScanType.Add(ComboBox_CfgScanType3);
            ComboBox_CfgScanType.Add(ComboBox_CfgScanType4);
            ComboBox_CfgScanType.Add(ComboBox_CfgScanType5);
            ComboBox_CfgWidth.Clear();
            ComboBox_CfgWidth1.Items.Clear();
            ComboBox_CfgWidth2.Items.Clear();
            ComboBox_CfgWidth3.Items.Clear();
            ComboBox_CfgWidth4.Items.Clear();
            ComboBox_CfgWidth5.Items.Clear();
            ComboBox_CfgWidth.Add(ComboBox_CfgWidth1);
            ComboBox_CfgWidth.Add(ComboBox_CfgWidth2);
            ComboBox_CfgWidth.Add(ComboBox_CfgWidth3);
            ComboBox_CfgWidth.Add(ComboBox_CfgWidth4);
            ComboBox_CfgWidth.Add(ComboBox_CfgWidth5);
            ComboBox_CfgExposure.Clear();
            ComboBox_CfgExposure1.Items.Clear();
            ComboBox_CfgExposure2.Items.Clear();
            ComboBox_CfgExposure3.Items.Clear();
            ComboBox_CfgExposure4.Items.Clear();
            ComboBox_CfgExposure5.Items.Clear();
            ComboBox_CfgExposure.Add(ComboBox_CfgExposure1);
            ComboBox_CfgExposure.Add(ComboBox_CfgExposure2);
            ComboBox_CfgExposure.Add(ComboBox_CfgExposure3);
            ComboBox_CfgExposure.Add(ComboBox_CfgExposure4);
            ComboBox_CfgExposure.Add(ComboBox_CfgExposure5);
            TextBox_CfgRangeStart.Clear();
            TextBox_CfgRangeStart.Add(TextBox_CfgRangeStart1);
            TextBox_CfgRangeStart.Add(TextBox_CfgRangeStart2);
            TextBox_CfgRangeStart.Add(TextBox_CfgRangeStart3);
            TextBox_CfgRangeStart.Add(TextBox_CfgRangeStart4);
            TextBox_CfgRangeStart.Add(TextBox_CfgRangeStart5);
            TextBox_CfgRangeEnd.Clear();
            TextBox_CfgRangeEnd.Add(TextBox_CfgRangeEnd1);
            TextBox_CfgRangeEnd.Add(TextBox_CfgRangeEnd2);
            TextBox_CfgRangeEnd.Add(TextBox_CfgRangeEnd3);
            TextBox_CfgRangeEnd.Add(TextBox_CfgRangeEnd4);
            TextBox_CfgRangeEnd.Add(TextBox_CfgRangeEnd5);
            TextBox_CfgDigRes.Clear();
            TextBox_CfgDigRes.Add(TextBox_CfgDigRes1);
            TextBox_CfgDigRes.Add(TextBox_CfgDigRes2);
            TextBox_CfgDigRes.Add(TextBox_CfgDigRes3);
            TextBox_CfgDigRes.Add(TextBox_CfgDigRes4);
            TextBox_CfgDigRes.Add(TextBox_CfgDigRes5);
            Label_Pattern.Clear();
            Label_Pattern.Add(label_pattern1);
            Label_Pattern.Add(label_pattern2);
            Label_Pattern.Add(label_pattern3);
            Label_Pattern.Add(label_pattern4);
            Label_Pattern.Add(label_pattern5);

            for (Int32 i = 0; i < MAX_CFG_SECTION; i++)
            {
                // Initialize combobox items
                for (Int32 j = 0; j < 2; j++)
                {
                    String Type = Helper.ScanTypeIndexToMode(j).Substring(0, 3);
                    ComboBox_CfgScanType[i].Items.Add(Type);
                }
                for (Int32 j = 0; j < Helper.CfgWidthItemsCount(); j++)
                {
                    Double WidthNM = Helper.CfgWidthIndexToNM(j);
                    ComboBox_CfgWidth[i].Items.Add(Math.Round(WidthNM, 2));
                }
                for (Int32 j = 0; j < Helper.CfgExpItemsCount(); j++)
                {
                    Double ExpTime = Helper.CfgExpIndexToTime(j);
                    ComboBox_CfgExposure[i].Items.Add(ExpTime);
                }
            }
        }

        private void SetDetailColorWhite()
        {
            TextBox_CfgName.BackColor = System.Drawing.Color.White;
            TextBox_CfgAvg.BackColor = System.Drawing.Color.White;
            for(int i=0;i< TextBox_CfgRangeStart.Count;i++)
            {
                TextBox_CfgRangeStart[i].BackColor = System.Drawing.Color.White;
                TextBox_CfgRangeEnd[i].BackColor = System.Drawing.Color.White;
                TextBox_CfgDigRes[i].BackColor = System.Drawing.Color.White;
            }
        }

        private void EnableCfgItem(bool enable)
        {
            TextBox_CfgName.Enabled = enable;
            TextBox_CfgAvg.Enabled = enable;
            comboBox_cfgNumSec.Enabled = enable;
            for (int i = 0; i < 5; i++)
            {
                ComboBox_CfgScanType[i].Enabled = enable;
                TextBox_CfgRangeStart[i].Enabled = enable;
                TextBox_CfgRangeEnd[i].Enabled = enable;
                ComboBox_CfgWidth[i].Enabled = enable;
                TextBox_CfgDigRes[i].Enabled = enable;
                ComboBox_CfgExposure[i].Enabled = enable;
            }
        }

        private void InitCfgDetailsContent()
        {
            TextBox_CfgName.Clear();
            TextBox_CfgAvg.Clear();
            comboBox_cfgNumSec.SelectedIndex = 0;

            for (Int32 i = 0; i < MAX_CFG_SECTION; i++)
            {
                ComboBox_CfgScanType[i].SelectedIndex = 0;
                TextBox_CfgRangeStart[i].Clear();
                TextBox_CfgRangeEnd[i].Clear();
                ComboBox_CfgWidth[i].SelectedIndex = 5;
                ComboBox_CfgExposure[i].SelectedIndex = 0;
                TextBox_CfgDigRes[i].Clear();
                Label_Pattern[i].Text = String.Empty;
            }
        }
        private void LoadConfigDir()
        {
            // Config Directory
            String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            ConfigDir = Path.Combine(path, "InnoSpectra\\Config Data");

            if (Directory.Exists(ConfigDir) == false)
            {
                Directory.CreateDirectory(ConfigDir);
                DBG.WriteLine("The directory {0} was created.", ConfigDir);
            }
        }
        private void LoadScanPageSetting()
        {
            String FilePath = Path.Combine(ConfigDir, "ScanPageSettings.xml");
            if (!File.Exists(FilePath))
            {
                String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Scan_Dir = Path.Combine(path, "InnoSpectra\\Scan Results");
                Display_Dir = Scan_Dir;
                ScanFile_Formats = 0x81;

                if (Directory.Exists(Scan_Dir) == false)
                {
                    Directory.CreateDirectory(Scan_Dir);
                    DBG.WriteLine("The directory {0} was created.", Scan_Dir);
                }
            }
            else
            {
                XmlDocument XmlDoc = new XmlDocument();
                XmlDoc.Load(FilePath);

                XmlNode ScanDir = XmlDoc.SelectSingleNode("/Settings/ScanDir");
                if (ScanDir.InnerText == String.Empty)
                    Scan_Dir = Path.Combine(Directory.GetCurrentDirectory(), "Scan Results");
                else
                    Scan_Dir = ScanDir.InnerText;

                XmlNode DisplayDir = XmlDoc.SelectSingleNode("/Settings/DisplayDir");
                if (DisplayDir.InnerText == String.Empty)
                    Display_Dir = Scan_Dir;
                else
                    Display_Dir = DisplayDir.InnerText;

                XmlNode FileFormats = XmlDoc.SelectSingleNode("/Settings/FileFormats");
                if (FileFormats.InnerText == String.Empty)
                    ScanFile_Formats = 0x81;
                else
                    ScanFile_Formats = Int32.Parse(FileFormats.InnerText);
            }
            Int32 buf_ScanFile_Formats = ScanFile_Formats;
            CheckBox_SaveCombCSV.Checked = ((buf_ScanFile_Formats & 0x01) >> 0 == 1) ? true : false;
            CheckBox_SaveICSV.Checked = ((buf_ScanFile_Formats & 0x02) >> 1 == 1) ? true : false;
            CheckBox_SaveACSV.Checked = ((buf_ScanFile_Formats & 0x04) >> 2 == 1) ? true : false;
            CheckBox_SaveRCSV.Checked = ((buf_ScanFile_Formats & 0x08) >> 3 == 1) ? true : false;
            CheckBox_SaveIJDX.Checked = ((buf_ScanFile_Formats & 0x10) >> 4 == 1) ? true : false;
            CheckBox_SaveAJDX.Checked = ((buf_ScanFile_Formats & 0x20) >> 5 == 1) ? true : false;
            CheckBox_SaveRJDX.Checked = ((buf_ScanFile_Formats & 0x40) >> 6 == 1) ? true : false;
            CheckBox_SaveDAT.Checked = ((buf_ScanFile_Formats & 0x80) >> 7 == 1) ? true : false;
        }
        #endregion

        #region Scan
        private void BeginScan()
        {
            /*try { ProgressWindowCompleted(); } catch { }
            ProgressWindowStart("Start scan", false);*/
            /* String text = "Start Scan!";
             pbws = new ProgressBar(text) { Owner = this };
             pbws.ShowDialog();*/
        }
        private void ScanCompleted()
        {
            //  ProgressWindowCompleted();
            //pbws.Close();
        }
        private void RadioButton_RefNew_CheckedChanged(object sender, EventArgs e)
        {
            if(RadioButton_RefNew.Checked == true)
            {
                Button_Scan.Text = "Reference Scan";
                ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_NEW;
                GroupBox_ContScan.Enabled = false;
                CheckBox_SaveOneCSV.Enabled = false;
                label_ref.Visible = false;
            }
        }


        private void RadioButton_RefPre_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButton_RefPre.Checked == false)
                return;
            else if (Scan.IsLocalRefExist)
            {              
                Button_Scan.Text = "Scan";
                ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_PREV;
                GroupBox_ContScan.Enabled = true;
                Text_ContScan_TextChanged(null, null);

                if(!String.IsNullOrEmpty(pre_ref_time))
                {
                    label_ref.Visible = true;
                    label_ref.Text = pre_ref_time;
                }
                else
                {
                    label_ref.Visible = false;
                }
            }
            else
            {
                RadioButton_RefPre.Checked = false;
                RadioButton_RefPre.Enabled = false;
                RadioButton_RefNew.Checked = true;
                label_ref.Text = "";
                label_ref.Visible = false;
            }
        }

        private void RadioButton_RefFac_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButton_RefFac.Checked == true)
            {
                Button_Scan.Text = "Scan";
                ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_BUILT_IN;
                GroupBox_ContScan.Enabled = true;
                Text_ContScan_TextChanged(null, null);
                if (!String.IsNullOrEmpty(buildin_ref_time))
                {
                    label_ref.Visible = true;
                    label_ref.Text = buildin_ref_time;
                }
                else
                {
                    label_ref.Visible = false;
                }
            }
        }

        #endregion
        #region Scan Config
        private void Button_SetActive_Click(object sender, EventArgs e)
        {
            if (TargetCfg_SelIndex < 0)
            {
                String text = "No item selected.";
                MessageBox.Show(text, "Warning");
                return;
            }
            ScanConfig.SetTargetActiveScanIndex(TargetCfg_SelIndex);
            if(DevCurCfg_IsTarget)
            {
                SetScanConfig(ScanConfig.TargetConfig[DevCurCfg_Index], true, DevCurCfg_Index);
            }
            else
            {
                SetScanConfig(LocalConfig[DevCurCfg_Index], false, DevCurCfg_Index);
            }
            int bufindex = TargetCfg_SelIndex;
            RefreshTargetCfgList();
            ListBox_TargetCfgs.SelectedIndex = bufindex;
        }
        private void SetScanConfig(ScanConfig.SlewScanConfig Config,Boolean IsTarget , Int32 index)
        {
            ClearScanPlotsUI();
            if (ScanConfig.SetScanConfig(Config) == SDK.FAIL)
            {
                String text = "Device config (" + Config.head.config_name + ") is not correct, please check it again!";
                MessageBox.Show(text, "Error");
            }
            else
            {
                DevCurCfg_Index = index;
                DevCurCfg_IsTarget = IsTarget;

                if (IsTarget)
                    Label_CurrentConfig.Text = "Device Config: " + Config.head.config_name;
                else
                    Label_CurrentConfig.Text = "Local Config: " + Config.head.config_name;

                textBox_ScanAvg.Text = Config.head.num_repeats.ToString();
                Double ScanTime = Scan.GetEstimatedScanTime();
                if (ScanTime > 0)
                    Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";

                Button_Scan.Enabled = true;

            }
        }
        private void ListBox_DeviceScanConfig_SelectedIndexChanged(object sender, EventArgs e)
        {
            isSelectConfig = true;
            if (NewConfig == true || EditConfig == true)
            {
                EditConfig = false;
                NewConfig = false;
                Button_CfgCancel_Click(this, e);
            }
            TargetCfg_Last_SelIndex = TargetCfg_SelIndex;
            TargetCfg_SelIndex = ListBox_TargetCfgs.SelectedIndex;
            if (TargetCfg_SelIndex < 0 || ScanConfig.TargetConfig.Count == 0)
            {
                if(ListBox_TargetCfgs.SelectedIndex ==-1 && ListBox_LocalCfgs.SelectedIndex == -1)//new config situation
                {
                    return;
                }
                else
                {
                    ListBox_LocalCfgs.BackColor = System.Drawing.Color.AliceBlue;
                    ListBox_TargetCfgs.BackColor = System.Drawing.Color.White;
                    Button_SetActive.Enabled = false;
                    return;
                }               
            }
            if(ListBox_LocalCfgs.Items.Count == 0)
            {
                ListBox_LocalCfgs.BackColor = System.Drawing.Color.White;
                ListBox_TargetCfgs.BackColor = System.Drawing.Color.AliceBlue;
            }
            SelCfg_IsTarget = true;
            FillCfgDetailsContent();
            OpenColseScanConfigButton(nameof(ScanConfigMode.INITIAL));
            // Clear target listbox index after local config data refreshed.
            if (ListBox_LocalCfgs.SelectedIndex != -1)
                ListBox_LocalCfgs.SelectedIndex = -1;
            SetDetailColorWhite();
            isSelectConfig = false;
            IsCfgLegal(true);
        }
        private void ListBox_DeviceScanConfig_MouseDoubleClick(object sender, EventArgs e)
        {
            SetScanConfig(ScanConfig.TargetConfig[TargetCfg_SelIndex],true, TargetCfg_SelIndex);
            TargetCfg_SelIndex = ListBox_TargetCfgs.SelectedIndex;    
        }

        private void FillCfgDetailsContent()
        {
            Int32 i, NumSection = 0;
            ScanConfig.SlewScanConfig CurConfig = new ScanConfig.SlewScanConfig
            {
                section = new ScanConfig.SlewScanSection[5]
            };

            InitCfgDetailsContent();
            if (SelCfg_IsTarget == true)
                CurConfig = ScanConfig.TargetConfig[TargetCfg_SelIndex];
            else
                CurConfig = LocalConfig[LocalCfg_SelIndex];


            NumSection = CurConfig.head.num_sections;

            TextBox_CfgName.Text = CurConfig.head.config_name;
            TextBox_CfgAvg.Text = CurConfig.head.num_repeats.ToString();
            comboBox_cfgNumSec.SelectedItem = NumSection.ToString();

            for (i = 0; i < NumSection; i++)
            {
                ComboBox_CfgScanType[i].SelectedIndex = CurConfig.section[i].section_scan_type;
                TextBox_CfgRangeStart[i].Text = CurConfig.section[i].wavelength_start_nm.ToString();
                TextBox_CfgRangeEnd[i].Text = CurConfig.section[i].wavelength_end_nm.ToString();
                ComboBox_CfgWidth[i].SelectedIndex = (Helper.CfgWidthPixelToIndex(CurConfig.section[i].width_px) > -1) ? (Helper.CfgWidthPixelToIndex(CurConfig.section[i].width_px)) : (5);
                TextBox_CfgDigRes[i].Text = CurConfig.section[i].num_patterns.ToString();
                ComboBox_CfgExposure[i].SelectedIndex = CurConfig.section[i].exposure_time;
            }
            EnableCfgItem(false);
        }
        private void RefreshTargetCfgList()
        {
            TargetCfg_SelIndex = -1;
            Int32 ActiveIndex = ScanConfig.GetTargetActiveScanIndex();

            ListBox_TargetCfgs.Items.Clear();
            if (ScanConfig.TargetConfig.Count > 0)
            {
                for (Int32 i = 0; i < ScanConfig.TargetConfig.Count; i++)
                {
                    ListBox_TargetCfgs.Items.Add(ScanConfig.TargetConfig[i].head.config_name);
                }
                label_ActiveConfig.Text = ScanConfig.TargetConfig[ActiveIndex].head.config_name;
            }
        }

        private void Button_CfgNew_Click(object sender, EventArgs e)
        {
            // Clear listbox index before someone set focus.
            if (ListBox_LocalCfgs.SelectedIndex != -1)
                ListBox_LocalCfgs.SelectedIndex = -1;
            if (ListBox_TargetCfgs.SelectedIndex != -1)
                ListBox_TargetCfgs.SelectedIndex = -1;

            InitCfgDetailsContent();
            EnableCfgItem(true);
            NewConfig = true;
            comboBox_cfgNumSec.SelectedItem = "1";
            OpenColseScanConfigButton(nameof(ScanConfigMode.NEW));
            TextBox_CfgName.Focus();
        }

        private void Button_CfgCancel_Click(object sender, EventArgs e)
        {
            isCancelConfig = true;
            OpenColseScanConfigButton(nameof(ScanConfigMode.CANCEL));
            if (NewConfig || EditConfig)
            {
                int bufindex = 0;
                if (SelCfg_IsTarget == true)
                {
                    bufindex = TargetCfg_SelIndex;
                }
                else
                {
                    bufindex = LocalCfg_SelIndex;
                }
                InitCfgDetailsContent();
                EnableCfgItem(false);

                if (EditConfig)//Edit config
                {
                    if (SelCfg_IsTarget == true)
                    {
                        ListBox_TargetCfgs.ClearSelected();
                        ListBox_TargetCfgs.SelectedIndex = bufindex;
                    }
                    else
                    {
                        ListBox_LocalCfgs.ClearSelected();
                        ListBox_LocalCfgs.SelectedIndex = bufindex;
                    }
                        
                }
                else//New config
                {
                    if (SelCfg_IsTarget == true)
                    {
                        ListBox_TargetCfgs.ClearSelected();
                        ListBox_TargetCfgs.SelectedIndex = TargetCfg_Last_SelIndex;
                    }
                    else
                    {
                        ListBox_LocalCfgs.ClearSelected();
                        ListBox_LocalCfgs.SelectedIndex = LocalCfg_Last_SelIndex;
                    }
                }

                Button_CfgEdit.Enabled = true;
                Button_CfgDelete.Enabled = true;
                Button_CfgNew.Enabled = true;
                NewConfig = false;
                EditConfig = false;
            }
            isCancelConfig = false;
        }

        private void Button_CfgSave_Click(object sender, EventArgs e)
        {
            EnableCfgItem(false);
            OpenColseScanConfigButton(nameof(ScanConfigMode.SAVE));
            if (IsCfgLegal(true) == SDK.FAIL)
            {
                Message.ShowError("Error configuration data can't be saved!");
                EnableCfgItem(true);
                Button_CfgCancel.Enabled = true;
                return;
            }
            if(NewConfig == true)
            {
                if(SelCfg_IsTarget && ListBox_TargetCfgs.BackColor == System.Drawing.Color.AliceBlue)
                {
                    if(checkConfigName(false,false))//device and not edit
                    {
                        Message.ShowError("The name has exist in device list!");
                        EnableCfgItem(true);
                        Button_CfgCancel.Enabled = true;
                        return;
                    }
                    if (ScanConfig.TargetConfig.Count >= 20)//Confirm the current number of device configuration before saving
                    {
                        Message.ShowWarning("Number of scan configs in device cannot exceed 20.");
                        EnableCfgItem(true);
                        Button_CfgCancel.Enabled = true;
                        return;
                    }
                    SaveCfgToList(true, true);//target and new
                    NewConfig = false;
                    if (DevCurCfg_IsTarget)
                    {
                        SetScanConfig(ScanConfig.TargetConfig[DevCurCfg_Index], true, DevCurCfg_Index);
                    }
                    else
                    {
                        SetScanConfig(LocalConfig[DevCurCfg_Index], false, DevCurCfg_Index);
                    }
                }
                else
                {
                    if (checkConfigName(true,false))//local and not edit
                    {
                        Message.ShowError("The name has exist in local list!");
                        EnableCfgItem(true);
                        Button_CfgCancel.Enabled = true;
                        return;
                    }
                    SaveCfgToList(false, true);//Local and new
                    NewConfig = false;
                }
            }
            else if(EditConfig == true)
            {
                if (SelCfg_IsTarget == true)
                {
                    if (checkConfigName(false, true))//device and edit
                    {
                        Message.ShowError("The name has exist in device list!");
                        EnableCfgItem(true);
                        Button_CfgCancel.Enabled = true;
                        return;
                    }
                    SaveCfgToList(true,false);//target and edit
                    if(DevCurCfg_IsTarget && DevCurCfg_Index == TargetCfg_SelIndex)//update device config
                    {
                        SetScanConfig(ScanConfig.TargetConfig[DevCurCfg_Index], true, DevCurCfg_Index);
                        ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_BUILT_IN;
                        RadioButton_RefFac.PerformClick();
                    }
                    EditConfig = false;
                }
                else
                {
                    if (checkConfigName(true, true))//local and edit
                    {
                        Message.ShowError("The name has exist in local list!");
                        EnableCfgItem(true);
                        Button_CfgCancel.Enabled = true;
                        return;
                    }
                    SaveCfgToList(false, false);//Local and edit
                    if (!DevCurCfg_IsTarget && DevCurCfg_Index == LocalCfg_SelIndex)//update device config
                    {
                        SetScanConfig(LocalConfig[DevCurCfg_Index], false, DevCurCfg_Index);
                        ReferenceSelect = Scan.SCAN_REF_TYPE.SCAN_REF_BUILT_IN;
                        RadioButton_RefFac.PerformClick();
                    }
                    EditConfig = false;
                }
            }
            Button_CfgEdit.Enabled = true;
            Button_CfgDelete.Enabled = true;
            Button_CfgNew.Enabled = true;
        }
        private Boolean checkConfigName(Boolean isLocal,Boolean isEdit)
        {
            Boolean isExist = false;
            if(isLocal)
            {
                for(int i=0;i< ListBox_LocalCfgs.Items.Count;i++)
                {
                    if(ListBox_LocalCfgs.Items[i].ToString() == TextBox_CfgName.Text)
                    {
                        if (isEdit && i == LocalCfg_SelIndex)
                        {
                        }
                        else
                        {
                            TextBox_CfgName.BackColor = System.Drawing.Color.LightPink;
                            isExist = true;
                            return isExist;
                        }                     
                    }
                }
            }
            else
            {
                for (int i = 0; i < ListBox_TargetCfgs.Items.Count; i++)
                {
                    if (ListBox_TargetCfgs.Items[i].ToString() == TextBox_CfgName.Text)
                    {
                        if(isEdit && i == TargetCfg_SelIndex)
                        {
                        }
                        else
                        {
                            TextBox_CfgName.BackColor = System.Drawing.Color.LightPink;
                            isExist = true;
                            return isExist;
                        }                    
                    }
                }
            }
            return isExist;
        }
        private Int32 SaveCfgToList(Boolean IsTarget, Boolean IsNew)
        {
            ScanConfig.SlewScanConfig CurConfig = new ScanConfig.SlewScanConfig
            {
                section = new ScanConfig.SlewScanSection[5]
            };

            CurConfig.head.config_name = Helper.CheckRegex(TextBox_CfgName.Text);
            CurConfig.head.scan_type = 2;
            CurConfig.head.num_sections = Byte.Parse(comboBox_cfgNumSec.SelectedItem.ToString());
            CurConfig.head.num_repeats = UInt16.Parse(TextBox_CfgAvg.Text);

            for (Int32 i = 0; i < CurConfig.head.num_sections; i++)
            {
                CurConfig.section[i].wavelength_start_nm = UInt16.Parse(TextBox_CfgRangeStart[i].Text);
                CurConfig.section[i].wavelength_end_nm = UInt16.Parse(TextBox_CfgRangeEnd[i].Text);
                CurConfig.section[i].num_patterns = UInt16.Parse(TextBox_CfgDigRes[i].Text);
                CurConfig.section[i].section_scan_type = (Byte)(ComboBox_CfgScanType[i].SelectedIndex);
                CurConfig.section[i].width_px = (Byte)Helper.CfgWidthIndexToPixel(ComboBox_CfgWidth[i].SelectedIndex);
                CurConfig.section[i].exposure_time = (UInt16)ComboBox_CfgExposure[i].SelectedIndex;
            }

            if ( IsNew == true)
            {
                if(IsTarget)
                {
                    ScanConfig.TargetConfig.Add(CurConfig);
                    RefreshTargetCfgList();
                }
                else
                {
                    LocalConfig.Add(CurConfig);
                    RefreshLocalCfgList();
                }
            }
            else
            {
                if(IsTarget)
                {
                    ScanConfig.TargetConfig.RemoveAt(TargetCfg_SelIndex);
                    ScanConfig.TargetConfig.Insert(TargetCfg_SelIndex, CurConfig);
                    RefreshTargetCfgList();
                }
                else
                {
                    LocalConfig.RemoveAt(LocalCfg_SelIndex);
                    LocalConfig.Insert(LocalCfg_SelIndex, CurConfig);
                    RefreshLocalCfgList();
                }
            }
            return SaveCfgToLocalOrDevice(IsTarget);
        }

        private void Button_CfgDelete_Click(object sender, EventArgs e)
        {
            if (SelCfg_IsTarget == true)
            {
                if (TargetCfg_SelIndex < 0)
                {
                    Message.ShowError("No item selected.");
                    return;
                }
                else if (TargetCfg_SelIndex == 0)
                {
                    Message.ShowError("The built-in default configuration cannot be deleted.");
                    return;
                }
            }
            else
            {
                if (LocalCfg_SelIndex < 0)
                {
                    Message.ShowError("No item selected.");
                    return;
                }
            }
            OpenColseScanConfigButton(nameof(ScanConfigMode.DELETE));
            Int32 ActiveIndex = ScanConfig.GetTargetActiveScanIndex();
            if (ScanConfig.TargetConfig.Count > 1 && SelCfg_IsTarget)
            {                
                ScanConfig.TargetConfig.RemoveAt(TargetCfg_SelIndex);
                if (TargetCfg_SelIndex == ActiveIndex)
                {
                    ActiveIndex = 0;
                    ScanConfig.SetTargetActiveScanIndex(ActiveIndex);
                }
                else if (TargetCfg_SelIndex < ActiveIndex)
                {
                    ActiveIndex--;
                    ScanConfig.SetTargetActiveScanIndex(ActiveIndex);
                }
                Boolean deletecurrentconfig = false;
                if (TargetCfg_SelIndex == DevCurCfg_Index)//刪到current config,因此將current config 換成機器的active config
                {
                    SetScanConfig(ScanConfig.TargetConfig[ActiveIndex],true, ActiveIndex);
                    deletecurrentconfig = true;
                }
                RefreshTargetCfgList();
                SaveCfgToLocalOrDevice(true);
                if (deletecurrentconfig)//刪到current config,因此將current config 換成機器的active config
                {
                    ListBox_TargetCfgs.SelectedIndex = ActiveIndex;
                }
                DBG.WriteLine("Delete this Device configuration");
            }
            else
            {
                LocalConfig.RemoveAt(LocalCfg_SelIndex);
                Boolean deletecurrentconfig = false;
                if (LocalCfg_SelIndex == DevCurCfg_Index && DevCurCfg_IsTarget == false)//刪到current config,因此將current config 換成機器的active config
                {
                    SetScanConfig(ScanConfig.TargetConfig[ActiveIndex], true, ActiveIndex);
                    deletecurrentconfig = true;
                }
                RefreshLocalCfgList();
                SaveCfgToLocalOrDevice(false);
                if (deletecurrentconfig)//刪到current config,因此將current config 換成機器的active config
                {
                    ListBox_TargetCfgs.SelectedIndex = ActiveIndex;
                }
                if(ListBox_LocalCfgs.Items.Count==0 && !deletecurrentconfig)
                {
                    Button_CfgNew.Enabled = true;
                    Button_CfgEdit.Enabled = false;
                    Button_CfgDelete.Enabled = false;
                    Button_CfgSave.Enabled = false;
                    Button_CfgCancel.Enabled = false;
                    ClearDetailValue();
                }
            }
        }

        private void Button_CfgEdit_Click(object sender, EventArgs e)
        {
            if (SelCfg_IsTarget == true)
            {
                if (TargetCfg_SelIndex < 0)
                {
                    Message.ShowWarning("No item selected!");
                    return;
                }
                EditSelectIndex = ListBox_TargetCfgs.SelectedIndex;
            }
            else
            {
                if (LocalCfg_SelIndex < 0)
                {
                    Message.ShowWarning("No item selected!");
                    return;
                }
                EditSelectIndex = ListBox_LocalCfgs.SelectedIndex;
            }
            EnableCfgItem(true);
            NewConfig = false;
            EditConfig = true;
            OpenColseScanConfigButton(nameof(ScanConfigMode.EDIT));
            TextBox_CfgName.Focus();
        }

        private void CfgDetails_Validated(object sender, EventArgs e)
        {
            IsCfgLegal(false);
        }

        private void CfgSection_SelectionChanged(object sender, EventArgs e)
        {
            int section = int.Parse(comboBox_cfgNumSec.SelectedItem.ToString());

            for (int i = 0; i < 5; i++)
            {
                if (i < section)
                {
                    ComboBox_CfgScanType[i].Visible = true;
                    TextBox_CfgRangeStart[i].Visible = true;
                    TextBox_CfgRangeEnd[i].Visible = true;
                    ComboBox_CfgWidth[i].Visible = true;
                    TextBox_CfgDigRes[i].Visible = true;
                    ComboBox_CfgExposure[i].Visible = true;
                    Label_Pattern[i].Visible = true;
                }
                else
                {
                    ComboBox_CfgScanType[i].Visible = false;
                    TextBox_CfgRangeStart[i].Visible = false;
                    TextBox_CfgRangeEnd[i].Visible = false;
                    ComboBox_CfgWidth[i].Visible = false;
                    TextBox_CfgDigRes[i].Visible = false;
                    ComboBox_CfgExposure[i].Visible = false;
                    Label_Pattern[i].Visible = false;
                }
            }
        }

        private Int32 IsCfgLegal(Boolean IsColored)
        {
            Int32 ret = SDK.PASS;
            Int32 TotalPatterns = 0;
            ScanConfig.SlewScanConfig CurConfig = new ScanConfig.SlewScanConfig
            {
                section = new ScanConfig.SlewScanSection[5]
            };
            CurConfig.head.scan_type = 2;

            // Config Name
            if (TextBox_CfgName.Text == String.Empty)
            {
                if (IsColored) TextBox_CfgName.BackColor = System.Drawing.Color.LightPink;
                ret = SDK.FAIL;
            }
            else
            {
                if (IsColored) TextBox_CfgName.BackColor = System.Drawing.Color.White;
                CurConfig.head.config_name = Helper.CheckRegex(TextBox_CfgName.Text);
            }

            // Num Scans to Average
            if (UInt16.TryParse(TextBox_CfgAvg.Text, out CurConfig.head.num_repeats) == false || CurConfig.head.num_repeats == 0)
            {
                if (IsColored) TextBox_CfgAvg.BackColor = System.Drawing.Color.LightPink;
                ret = SDK.FAIL;
            }
            else
            {
                if (IsColored) TextBox_CfgAvg.BackColor = System.Drawing.Color.White;
            }

            // Sections
            CurConfig.head.num_sections = byte.Parse(comboBox_cfgNumSec.SelectedItem.ToString());
            for (Byte i = 0; i < CurConfig.head.num_sections; i++)
            {
                CurConfig.section[i].section_scan_type = (Byte)(ComboBox_CfgScanType[i].SelectedIndex);
                CurConfig.section[i].width_px = (Byte)Helper.CfgWidthIndexToPixel(ComboBox_CfgWidth[i].SelectedIndex);
                CurConfig.section[i].exposure_time = (UInt16)ComboBox_CfgExposure[i].SelectedIndex;

                // Start nm
                if (UInt16.TryParse(TextBox_CfgRangeStart[i].Text, out CurConfig.section[i].wavelength_start_nm) == false ||
                    CurConfig.section[i].wavelength_start_nm < MIN_WAVELENGTH)
                {
                    if (IsColored) TextBox_CfgRangeStart[i].BackColor = System.Drawing.Color.LightPink;
                    ret = SDK.FAIL;
                }
                else
                {
                    if (IsColored) TextBox_CfgRangeStart[i].BackColor = System.Drawing.Color.White;
                }

                // End nm
                if (UInt16.TryParse(TextBox_CfgRangeEnd[i].Text, out CurConfig.section[i].wavelength_end_nm) == false ||
                    CurConfig.section[i].wavelength_end_nm > MAX_WAVELENGTH || CurConfig.section[i].wavelength_end_nm < MIN_WAVELENGTH)
                {
                    if (IsColored) TextBox_CfgRangeEnd[i].BackColor = System.Drawing.Color.LightPink;
                    ret = SDK.FAIL;
                }
                else
                {
                    if (IsColored) TextBox_CfgRangeEnd[i].BackColor = System.Drawing.Color.White;
                }
                if (CurConfig.section[i].wavelength_start_nm >= CurConfig.section[i].wavelength_end_nm)
                {
                    if (IsColored) TextBox_CfgRangeStart[i].BackColor = System.Drawing.Color.LightPink;
                    if (IsColored) TextBox_CfgRangeEnd[i].BackColor = System.Drawing.Color.LightPink;
                    ret = SDK.FAIL;
                }

                Int32 MaxPattern = 0;
                Int32 HadPattern = 0;
                // Check Max Patterns(user input start wav and end wav will check)
                if (UInt16.TryParse(TextBox_CfgRangeStart[i].Text, out CurConfig.section[i].wavelength_start_nm) == true &&
                    CurConfig.section[i].wavelength_start_nm >= MIN_WAVELENGTH && 
                    UInt16.TryParse(TextBox_CfgRangeEnd[i].Text, out CurConfig.section[i].wavelength_end_nm) == true &&
                    CurConfig.section[i].wavelength_end_nm <= MAX_WAVELENGTH && CurConfig.section[i].wavelength_end_nm >= MIN_WAVELENGTH)
                {

                    MaxPattern = ScanConfig.GetMaxResolutions(CurConfig, i);
                    if ((UInt16.TryParse(TextBox_CfgDigRes[i].Text, out CurConfig.section[i].num_patterns) == false) ||
                        (CurConfig.section[i].section_scan_type == 0 && CurConfig.section[i].num_patterns < 2) ||  // Column Mode
                        (CurConfig.section[i].section_scan_type == 1 && CurConfig.section[i].num_patterns < 3) ||  // Hadamard Mode
                        (CurConfig.section[i].num_patterns > MaxPattern) ||
                        (MaxPattern <= 0))
                    {
                        if (IsColored) TextBox_CfgDigRes[i].BackColor = System.Drawing.Color.LightPink;
                        if (MaxPattern < 0) MaxPattern = 0;
                        ret = SDK.FAIL;
                    }
                    else
                    {
                        if (IsColored) TextBox_CfgDigRes[i].BackColor = System.Drawing.Color.White;
                        HadPattern = ScanConfig.GetHadamardUsedPatterns(CurConfig, i);
                        if(HadPattern!=-1)
                        {
                            if(HadPattern > MaxPattern)
                            {
                                if (IsColored) TextBox_CfgDigRes[i].BackColor = System.Drawing.Color.LightPink;
                                ret = SDK.FAIL;
                            }
                        }
                        else
                        {
                            if(CurConfig.section[i].num_patterns> MaxPattern)
                            {
                                if (IsColored) TextBox_CfgDigRes[i].BackColor = System.Drawing.Color.LightPink;
                                ret = SDK.FAIL;
                            }
                        }
                    }
                    
                }
                if (HadPattern!=-1)
                {
                    Label_Pattern[i].Text = HadPattern.ToString() + "/" + MaxPattern.ToString();
                    TotalPatterns += HadPattern;
                }
                else
                {
                    Label_Pattern[i].Text = CurConfig.section[i].num_patterns.ToString() + "/" + MaxPattern.ToString();
                    TotalPatterns += CurConfig.section[i].num_patterns;
                }
               
            }

            // Check total patterns
            if (TotalPatterns > 624 && !isCancelConfig && !isSelectConfig || (TotalPatterns > 624 && NewConfig == true && !isCancelConfig))
            {
                String text = "Total number of patterns " + TotalPatterns.ToString() + " exceeds 624!";
                Message.ShowWarning(text);
                ret = SDK.FAIL;
            }

            return ret;
        }
        #endregion
        #region Save scan
        private void LoadSavedScanList(Boolean filter)
        {
            List<String> ListFiles = EnumerateFiles("*.dat");
            dataGridView_savescan.Rows.Clear();
   
            if (ListFiles.Count>0)
            {
                if(filter)
                {
                    int count = 0;
                    for (int i = 0; i < ListFiles.Count; i++)
                    {
                        if(ListFiles[i].Contains(textBox_filter.Text))
                        {
                            count++;
                        }                       
                    }
                    dataGridView_savescan.RowCount = count;
                    int index = 0;
                    for (int i = 0; i < ListFiles.Count; i++)
                    {
                        if (ListFiles[i].Contains(textBox_filter.Text))
                        {
                            dataGridView_savescan.Rows[index].Cells["FileName"].Value = ListFiles[i];
                            dataGridView_savescan.Rows[index].Cells["Time"].Value = ListFilesTime[i];
                            index++;
                        }
                    }
                    if(count!=0 && tabScanPage.SelectedIndex == 2)
                    {
                        dataGridView_savescan.Rows[0].Selected = true;
                        dataGridView_savescan_MouseClick(null, null);
                    }                  
                }
                else
                {
                    dataGridView_savescan.RowCount = ListFiles.Count;
                    for (int i = 0; i < ListFiles.Count; i++)
                    {
                        dataGridView_savescan.Rows[i].Cells["FileName"].Value = ListFiles[i];
                        dataGridView_savescan.Rows[i].Cells["Time"].Value = ListFilesTime[i];
                    }
                    if(tabScanPage.SelectedIndex == 2)
                    {
                        dataGridView_savescan.Rows[0].Selected = true;
                        dataGridView_savescan_MouseClick(null, null);
                    }                 
                }               
            }            
            if (!LoadFileSuccess)
            {
                LoadFileSuccess = true;
                String FilePath = Path.Combine(ConfigDir, "ScanPageSettings.xml");
                File.Delete(FilePath);
                LoadScanPageSetting();
                SaveSettings();
                String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Scan_Dir = Path.Combine(path, "InnoSpectra\\Scan Results");

                Display_Dir = Scan_Dir;
                TextBox_DisplayDirPath.Text = Scan_Dir;
                TextBox_SaveDirPath.Text = Scan_Dir;
                LoadSavedScanList(false);
                //ClearSavedScanCfgItems();              
            }
        }
        private void CheckScanDirPath()
        {
            if (!Directory.Exists(TextBox_SaveDirPath.Text))
            {
                Message.ShowWarning("The scan directory has not exist. Will set to default path.");
                String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Scan_Dir = Path.Combine(path, "InnoSpectra\\Scan Results");
                TextBox_SaveDirPath.Text = Scan_Dir;
            }
        }
        Boolean LoadFileSuccess = true;
        private List<String> EnumerateFiles(String SearchPattern)
        {
            List<String> ListFiles = new List<String>();
            ListFilesTime = new List<String>();
            try
            {
                String DirPath = TextBox_DisplayDirPath.Text;

                foreach (String Files in Directory.EnumerateFiles(DirPath, SearchPattern))
                {
                    String FileName = Files.Substring(Files.LastIndexOf("\\") + 1);
                    ListFiles.Add(FileName);
                    String FileTime = File.GetLastWriteTime(Files).ToString();
                    ListFilesTime.Add(FileTime);
                }
            }
            catch (UnauthorizedAccessException UAEx) { DBG.WriteLine(UAEx.Message); }
            catch (PathTooLongException PathEx) { DBG.WriteLine(PathEx.Message); }
            catch (Exception e)
            {
                Message.ShowWarning("The display directory has not exist. Will set to  default path.");
                LoadFileSuccess = false;
                DBG.WriteLine(e.Message);
            }
            return ListFiles;
        }
        //--------------------------------------------------------------------------------
        private void Button_DisplayDirChange_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = TextBox_DisplayDirPath.Text
            };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Display_Dir = dlg.SelectedPath;
                TextBox_DisplayDirPath.Text = dlg.SelectedPath;

                LoadSavedScanList(false);
                ClearSavedScanCfgItems();
            }
        }
        private void ClearSavedScanCfgItems()
        {
            for (Int32 i = 0; i < MAX_CFG_SECTION; i++)
            {
                Label_SavedScanType[i].Text = String.Empty;
                Label_SavedRangeStart[i].Text = String.Empty;
                Label_SavedRangeEnd[i].Text = String.Empty;
                Label_SavedWidth[i].Text = String.Empty;
                Label_SavedDigRes[i].Text = String.Empty;
                Label_SavedExposure[i].Text = String.Empty;
            }
            Label_SavedAvg.Text = String.Empty;
        }
        #endregion
        #region Chart
        private void initChart()
        {
            // Initial Chart
            RadioButton_Intensity.Checked = true; // Avoid the y-axis long number at initail stage
                                                  /* MyAxisY.Title = "Intensity";
                                                   MyAxisY.FontSize = 14;
                                                   MyAxisX.FontSize = 14;
                                                   MyAxisY.LabelFormatter = value => Math.Round(value, 3).ToString(); //set this to limit the label digits of axis
                                                   MyChart.ChartLegend = null;
                                                   MyChart.DataTooltip = null;
                                                   MyChart.Zoom = ZoomingOptions.None;*/
            MyChart.DataTooltip = null;
            MyChart.Zoom = ZoomingOptions.None;
        }
        private void SpectrumPlot()
        {
            if (MyChart.Series.Count > 0)
            {               
                if (!Check_Overlay.Checked)
                    MyChart.Series.Clear();
                else
                {
                    if (TargetScanCounts > 1 && ScannedCounts == 1)  //做continuous scan,掃描第一次要移除之前的serious,Y軸的圖才會有精確度 
                    {
                        MyChart.Series.Clear();
                    }
                    else if (MyChart.Series[0].Values.Count == 0)
                    {
                        MyChart.Series.RemoveAt(0);
                    }
                }
            }

            double[] valY = new double[Scan.ScanDataLen];
            double[] valX = new double[Scan.ScanDataLen];
            int dataCount = 0;
            string label = "";
            valX = Scan.WaveLength.ToArray();

            if ((bool)RadioButton_Intensity.Checked == true)
            {
                List<double> doubleList = Scan.Intensity.ConvertAll(x => (double)x);
                valY = doubleList.ToArray();
                label = "Intensity";
            }
            else if ((bool)RadioButton_Absorbance.Checked == true)
            {
                valY = Scan.Absorbance.ToArray();
                label = "Absorbance";
            }
            else if ((bool)RadioButton_Reflectance.Checked == true)
            {
                valY = Scan.Reflectance.ToArray();
                label = "Reflectance";
            }
            else if ((bool)RadioButton_Reference.Checked == true)
            {
                List<double> doubleList = Scan.ReferenceIntensity.ConvertAll(x => (double)x);
                valY = doubleList.ToArray();
                label = "Reference";
            }
            MyChart.AxisX.Clear();
            MyChart.AxisY.Clear();
            if(Scan.ScanConfigData.section!=null)
            {
                int min = MAX_WAVELENGTH, max = MIN_WAVELENGTH;
                GetMaxMinWav(ref min, ref max);
                MyChart.AxisX.Add(new Axis { Title = "Wavelength (nm)", MinValue = min, MaxValue = max });
            }
            else
            {
                MyChart.AxisX.Add(new Axis { Title = "Wavelength (nm)", MinValue = MIN_WAVELENGTH, MaxValue = MAX_WAVELENGTH });
            }          
            MyChart.AxisY.Add(new Axis { Title = label });

            for (int i = 0; i < Scan.ScanDataLen; i++)
            {
                if(valY.Length ==0)//fix 進去一開始點build-in再點overlay程式crash
                {
                    break;
                }
                if (Double.IsNaN(valY[i]) || Double.IsInfinity(valY[i]))
                    valY[i] = 0;
            }
            if (valY.Length == 0)//fix 進去一開始點build-in再點overlay程式crash
            { 
            }
            else if (Check_Overlay.Checked)
            {  
                for (int i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
                {
                    var ChartValues = new GearedValues<ObservablePoint>();
                    for (int j = 0; j < Scan.ScanConfigData.section[i].num_patterns; j++)
                        ChartValues.Add(new ObservablePoint(valX[j + dataCount], valY[j + dataCount]));

                    dataCount += Scan.ScanConfigData.section[i].num_patterns;
                    MyChart.Series.Add(new GLineSeries
                    {
                        Values = ChartValues,
                        Title = Scan.ScanConfigData.head.num_sections > 1
                        ? string.Format("#{0}->[{1}]", MyChart.Series.Count / Scan.ScanConfigData.head.num_sections + 1, i)
                        : string.Format("#{0}", MyChart.Series.Count + 1),
                        StrokeThickness = 1,
                        Fill = System.Windows.Media.Brushes.Transparent,
                        LineSmoothness = 0,
                        PointGeometry = null,
                        PointGeometrySize = 0,
                    });
                }
            }
            else
            {
                for (int i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
                {
                    var chartValues = new ChartValues<ObservablePoint>();
                    for (int j = 0; j < Scan.ScanConfigData.section[i].num_patterns; j++)
                        chartValues.Add(new ObservablePoint(valX[j + dataCount], valY[j + dataCount]));

                    dataCount += Scan.ScanConfigData.section[i].num_patterns;
                    MyChart.Series.Add(new LineSeries
                    {
                        Values = chartValues,
                        Title = Scan.ScanConfigData.head.num_sections == 1 ? string.Format("{0}", label) : string.Format("{0}\nsection[{1}]", label, i),
                        StrokeThickness = 1,
                        Fill = null,
                        LineSmoothness = 0,
                        PointGeometry = null,
                        PointGeometrySize = 0,
                    });
                }
            }
            

            // For initial the chart to avoid the crazy axis numbers
            if (Scan.ScanConfigData.head.num_sections == 0)
            {
                MyChart.Series.Add(new LineSeries
                {
                    Values = new ChartValues<ObservablePoint>(),
                    Title = "Intensity",
                    PointGeometry = null,
                    StrokeThickness = 1
                });
            }
        }
        private void GetMaxMinWav(ref int min, ref int max)
        {
            for (int i=0;i< Scan.ScanConfigData.head.num_sections;i++)
            {
                if(Scan.ScanConfigData.section[i].wavelength_start_nm < min)
                {
                    min = Scan.ScanConfigData.section[i].wavelength_start_nm;
                }
                if(Scan.ScanConfigData.section[i].wavelength_end_nm > max)
                {
                    max = Scan.ScanConfigData.section[i].wavelength_end_nm;
                }
            }
        }
        private void RadioButton_Reflectance_CheckedChanged(object sender, EventArgs e)
        {
            if (Check_Overlay.Checked)//清除overlay的資料
            {
                MyChart.Series.Clear();
            }
            SpectrumPlot();
        }

        private void RadioButton_Absorbance_CheckedChanged(object sender, EventArgs e)
        {
            if (Check_Overlay.Checked)//清除overlay的資料
            {
                MyChart.Series.Clear();
            }
            SpectrumPlot();
        }

        private void RadioButton_Intensity_CheckedChanged(object sender, EventArgs e)
        {
            if (Check_Overlay.Checked)//清除overlay的資料
            {
                MyChart.Series.Clear();
            }
            SpectrumPlot();
        }

        private void RadioButton_Reference_CheckedChanged(object sender, EventArgs e)
        {
            if (Check_Overlay.Checked)//清除overlay的資料
            {
                MyChart.Series.Clear();
            }
            SpectrumPlot();
        }
        #endregion
        #region scan item set
        private void CheckBox_AutoGain_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_AutoGain.Checked == true)
            {
                ComboBox_PGAGain.Enabled = false;
            }
            else
            {
                ComboBox_PGAGain.Enabled = true;
            }
        }
        private void CheckBox_LampOn_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_LampOn.Checked == true)
                RadioButton_LampOn_CheckedChanged(sender, e);
            else
                RadioButton_LampStableTime_CheckedChanged(sender, e);
        }
        private void RadioButton_LampOn_CheckedChanged(object sender, EventArgs e)
        {
            String HWRev = String.Empty;
            if (Device.IsConnected())
                HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

            //GetActivationKeyStatus();
            if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2 || label_ActivateStatus.Text.Equals("Activated!") == false)
            {
                CheckBox_AutoGain.Checked = false;
                CheckBox_AutoGain.Enabled = false;
                CheckBox_AutoGain_CheckedChanged(sender, e);
            }
            RadioButton_Absorbance.Enabled = true;
            RadioButton_Reflectance.Enabled = true;
            TextBox_LampStableTime.Enabled = false;
            Scan.SetLamp(Scan.LAMP_CONTROL.ON_SCAN);

            Double ScanTime = Scan.GetEstimatedScanTime();
            if (ScanTime > 0)
                Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";
        }

        private void RadioButton_LampOff_CheckedChanged(object sender, EventArgs e)
        {
            String HWRev = String.Empty;
            if (Device.IsConnected())
                HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

            if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2 || label_ActivateStatus.Text.Equals("Activated!") == false)
            {
                CheckBox_AutoGain.Checked = false;
                CheckBox_AutoGain.Enabled = false;
                CheckBox_AutoGain_CheckedChanged(sender, e);
            }
            if (RadioButton_LampOff.Checked)
            {
                RadioButton_Intensity.Checked = true;
            }
            TextBox_LampStableTime.Enabled = false;
            Scan.SetLamp(Scan.LAMP_CONTROL.OFF_SCAN);

            Double ScanTime = Scan.GetEstimatedScanTime();
            if (ScanTime > 0)
                Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";
        }

        private void RadioButton_LampStableTime_CheckedChanged(object sender, EventArgs e)
        {
            String HWRev = String.Empty;
            if (Device.IsConnected())
                HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

            if (label_ActivateStatus.Text.Equals("Activated!") == true)
                TextBox_LampStableTime.Enabled = true;
            CheckBox_AutoGain.Enabled = true;
            CheckBox_AutoGain_CheckedChanged(sender, e);
            RadioButton_Absorbance.Enabled = true;
            RadioButton_Reflectance.Enabled = true;
            Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);

            Double ScanTime = Scan.GetEstimatedScanTime();
            if (ScanTime > 0)
                Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";
        }

        private void ScanAvg_TextChanged(object sender, EventArgs e)
        {
             try
             {
                 ushort value = ushort.Parse(textBox_ScanAvg.Text.ToString());
                 Scan.SetScanNumRepeats(value);
                 Double ScanTime = Scan.GetEstimatedScanTime();
                 if (ScanTime > 0)
                     Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";
                 }
             catch (Exception e1)
             {
                 Console.WriteLine("Exception caught: {0}", e1);
             }
        }

        private void TextBox_LampStableTime_TextChanged(object sender, EventArgs e)
        {
            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_2)
            {
                if (UInt32.TryParse(TextBox_LampStableTime.Text, out LampStableTime) == false)
                {
                    String text = "Lamp Stable Time must be numeric!";
                    MessageBox.Show(text, "Warning");
                    TextBox_LampStableTime.Text = "625";
                    return;
                }

                Scan.SetLampDelay(LampStableTime);
            }
                
            Double ScanTime = Scan.GetEstimatedScanTime();
            if (ScanTime > 0)
                Label_EstimatedScanTime.Text = "Est. Device Scan Time: " + String.Format("{0:0.000}", ScanTime) + " secs.";
        }

        public enum GUI_State
        {
            DEVICE_ON,
            DEVICE_ON_SCANTAB_SELECT,
            DEVICE_OFF,
            DEVICE_OFF_SCANTAB_SELECT,
            SCAN,
            SCAN_FINISHED,
            FW_UPDATE,
            FW_UPDATE_FINISHED,
            REFERENCE_DATA_UPDATE,
            REFERENCE_DATA_UPDATE_FINISHED,
            KEY_ACTIVATE,
            KEY_NOT_ACTIVATE,
        };
        private void GUI_Handler(int state)
        {
            switch (state)
            {
                case (int)MainWindow.GUI_State.KEY_ACTIVATE:
                {
                    CheckBox_LampOn.Visible =false;
                    RadioButton_LampOn.Visible =true;

                    RadioButton_LampOff.Visible = true;
                    RadioButton_LampStableTime.Visible = true;
                    TextBox_LampStableTime.Visible = true;

                    RadioButton_LampOff.Enabled = true;
                    RadioButton_LampStableTime.Enabled = true;
                    TextBox_LampStableTime.Enabled = true;

                    RadioButton_LampStableTime.Checked = true;
                    if(CheckBox_CalWriteEnable.Checked)
                    {
                        Button_CalRestoreDefaultCoeffs.Enabled = true;
                    }

                    Label_ButtonStatus.Enabled = true;
                    Button_LockButton.Enabled = true;
                    Button_UnlockButton.Enabled = true;
                    GroupBox_BleName.Enabled = true;

                    toolStripStatus_DeviceStatus.Text = "Device " + Device.DevInfo.ModelName + " (" + Device.DevInfo.SerialNumber + ") connected!";
                    break;
                }
                case (int)MainWindow.GUI_State.KEY_NOT_ACTIVATE:
                {
                    String HWRev = String.Empty;
                    if (Device.IsConnected())
                        HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

                    CheckBox_LampOn.Visible = true;
                    RadioButton_LampOn.Visible = false;

                    RadioButton_LampOff.Visible = false;
                    RadioButton_LampStableTime.Visible = false;
                    TextBox_LampStableTime.Visible = false;

                    RadioButton_LampStableTime.Checked = false;

                    CheckBox_LampOn.Checked = false;
                    RadioButton_LampOn.Checked = false;
                    RadioButton_LampOff.Checked = false;
                    Button_CalRestoreDefaultCoeffs.Enabled = false;

                    Label_ButtonStatus.Enabled = false;
                    Button_LockButton.Enabled = false;
                    Button_UnlockButton.Enabled = false;
                    GroupBox_BleName.Enabled = false;

                    toolStripStatus_DeviceStatus.Text = "Device " + Device.DevInfo.ModelName + " (" + Device.DevInfo.SerialNumber + ") connected but advanced functions locked!";
                    break;
                }
                default:
                    break;
            }
        }
        #endregion
        #region save scan to file
        private void SaveToFiles()
        {
            String FileName = String.Empty;
            if (CheckBox_FileNamePrefix.Checked == true)
            {
                String Prefix = Helper.CheckRegex_Chinese(TextBox_FileNamePrefix.Text);
                if (Prefix.Length > 50)
                {
                    Prefix = Prefix.Substring(0, 50);
                    TextBox_FileNamePrefix.Text = Prefix;
                    Message.ShowWarning("File name prefix is too long, only catch the first 50 characters.");
                }
                FileName = Path.Combine(Scan_Dir, Prefix + Scan.ScanConfigData.head.config_name + "_" + TimeScanStart.ToString("yyyyMMdd_HHmmss"));
            }
            else
            {
                FileName = Path.Combine(Scan_Dir, Scan.ScanConfigData.head.config_name + "_" + TimeScanStart.ToString("yyyyMMdd_HHmmss"));
            }

            //check path
            String dirpath = Path.GetDirectoryName(FileName);
            String file = Path.GetFileName(FileName);
            if (!Directory.Exists(dirpath))
            {
                DialogResult result = Message.ShowQuestion("The directory has not exist. Do you want to create?\n    Yes,\t\t create directory.\n    No,\t\t save to default directory.\n    Cancel,\t\t not create and save.", null, MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Directory.CreateDirectory(dirpath);
                        TextBox_SaveDirPath.Text = dirpath;
                    }
                    catch (Exception e)
                    {
                        Message.ShowError("Create directroy failed!");
                        DBG.WriteLine(e.Message);
                        return;
                    };
                }
                else if (result == DialogResult.No)
                {
                    try
                    {
                        String path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        String defpath = Path.Combine(path, "InnoSpectra\\Scan Results");
                        if (!Directory.Exists(defpath))
                        {
                            Directory.CreateDirectory(defpath);
                        }
                        TextBox_SaveDirPath.Text = defpath;
                        FileName = defpath + "\\" + file;
                    }
                    catch (Exception e)
                    {
                        Message.ShowError("Create directroy failed!");
                        DBG.WriteLine(e.Message);
                        return;
                    };
                }
                else
                {
                    return;
                }
            }

            SaveToCSV(FileName + ".csv");
            SaveToJCAMP(FileName + ".jdx");

            if (CheckBox_SaveDAT.Checked == true)
                Scan.SaveScanResultToBinFile(FileName + ".dat");  // For populating saved scan
            LoadSavedScanList(false);
        }

        private void SaveToCSV(String FileName)
        {
            if (CheckBox_SaveCombCSV.Checked == true)
            {
                FileStream fs = new FileStream(FileName, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                SaveHeader_CSV(sw);

                sw.WriteLine("Wavelength (nm),Absorbance (AU),Reference Signal (unitless),Sample Signal (unitless)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Absorbance[i] + "," + Scan.ReferenceIntensity[i] + "," + Scan.Intensity[i]);
                }
                if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_4)
                {
                    byte[] HW_Ver = Encoding.ASCII.GetBytes(Device.DevInfo.HardwareRev);
                    int MB_Ver = HW_Ver[0];

                    int dataNum;
                    if (Scan.ScanConfigData.head.num_repeats < 30)
                        dataNum = Scan.ScanConfigData.head.num_repeats;
                    else
                        dataNum = 30;

                    sw.WriteLine("\n\nLampInScanAverage:");

                    for (int i = 0; i < dataNum; i++)
                    {
                        if (MB_Ver + 1 > 'F')
                        {
                            sw.WriteLine(Scan.SensorData[4 + i * 4] + "," + Scan.SensorData[4 + i * 4 + 1] + "," +
                            Scan.SensorData[4 + i * 4 + 2] + "," + Scan.SensorData[4 + i * 4 + 3]);
                        }
                        else
                        {
                            sw.WriteLine(Scan.SensorData[4 + i]);
                        }
                    }
                }

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream 
            }

            if (CheckBox_SaveICSV.Checked == true)
            {
                String FileName_i = FileName.Insert(FileName.LastIndexOf(".csv"), "_i");
                FileStream fs = new FileStream(FileName_i, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                SaveHeader_CSV(sw);

                sw.WriteLine("Wavelength (nm),Sample Signal (unitless)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Intensity[i]);
                }

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream
            }

            if (CheckBox_SaveACSV.Checked == true)
            {
                String FileName_a = FileName.Insert(FileName.LastIndexOf(".csv"), "_a");
                FileStream fs = new FileStream(FileName_a, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                SaveHeader_CSV(sw);

                sw.WriteLine("Wavelength (nm),Absorbance (AU)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Absorbance[i]);
                }

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream
            }

            if (CheckBox_SaveRCSV.Checked == true)
            {
                String FileName_r = FileName.Insert(FileName.LastIndexOf(".csv"), "_r");
                FileStream fs = new FileStream(FileName_r, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                SaveHeader_CSV(sw);

                sw.WriteLine("Wavelength (nm),Reflectance (AU)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Reflectance[i]);
                }

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream
            }

            if (SaveOneCSVFile)
            {
                if (OneScanFileName == String.Empty)
                    OneScanFileName = FileName;

                String FileName_one = OneScanFileName.Insert(OneScanFileName.LastIndexOf("_", OneScanFileName.Length - 20), "_one");

                using (FileStream fs = new FileStream(FileName_one, FileMode.Append, FileAccess.Write))
                {
                    using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8))
                    {
                        if (fs.Length == 0)
                        {
                            SaveHeader_CSV(sw);

                            sw.Write("Wavelength (nm),");
                            for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                                sw.Write(Scan.WaveLength[i] + ",");
                            sw.Write("\n");

                            sw.Write("Reference Signal (unitless),");
                            for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                                sw.Write(Scan.ReferenceIntensity[i] + ",");
                            sw.Write("\n");
                        }

                        sw.Write("Sample Signal (unitless),");
                        for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                            sw.Write(Scan.Intensity[i] + ",");
                        sw.Write("\n");
                    }
                }

                if (TargetScanCounts == ScannedCounts)
                {
                    SaveOneCSVFile = false;
                    OneScanFileName = String.Empty;
                }
            }
        }
        private void SaveToJCAMP(String FileName)
        {
            if (CheckBox_SaveIJDX.Checked == true)
            {
                String FileName_i = FileName.Insert(FileName.LastIndexOf("_", FileName.Length - 20), "_i");
                FileStream fs = new FileStream(FileName_i, FileMode.Create);
                SaveHeader(fs, out StreamWriter sw, true);

                sw.WriteLine("##XUNITS=Wavelength(nm)");
                sw.WriteLine("##YUNITS=Intensity");
                sw.WriteLine("##FIRSTX=" + Scan.WaveLength[0]);
                sw.WriteLine("##LASTX=" + Scan.WaveLength[Scan.ScanDataLen - 1]);
                sw.WriteLine("##PEAK TABLE=X+(Y..Y)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Intensity[i]);
                }
                sw.WriteLine("##END=");

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream 
            }

            if (CheckBox_SaveAJDX.Checked == true)
            {
                String FileName_a = FileName.Insert(FileName.LastIndexOf("_", FileName.Length - 20), "_a");
                FileStream fs = new FileStream(FileName_a, FileMode.Create);
                SaveHeader(fs, out StreamWriter sw, true);

                sw.WriteLine("##XUNITS=Wavelength(nm)");
                sw.WriteLine("##YUNITS=Absorbance(AU)");
                sw.WriteLine("##FIRSTX=" + Scan.WaveLength[0]);
                sw.WriteLine("##LASTX=" + Scan.WaveLength[Scan.ScanDataLen - 1]);
                sw.WriteLine("##PEAK TABLE=X+(Y..Y)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Absorbance[i]);
                }
                sw.WriteLine("##END=");

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream 
            }

            if (CheckBox_SaveRJDX.Checked == true)
            {
                String FileName_r = FileName.Insert(FileName.LastIndexOf("_", FileName.Length - 20), "_r");
                FileStream fs = new FileStream(FileName_r, FileMode.Create);
                SaveHeader(fs, out StreamWriter sw, true);

                sw.WriteLine("##XUNITS=Wavelength(nm)");
                sw.WriteLine("##YUNITS=Reflectance(AU)");
                sw.WriteLine("##FIRSTX=" + Scan.WaveLength[0]);
                sw.WriteLine("##LASTX=" + Scan.WaveLength[Scan.ScanDataLen - 1]);
                sw.WriteLine("##PEAK TABLE=X+(Y..Y)");
                for (Int32 i = 0; i < Scan.ScanDataLen; i++)
                {
                    sw.WriteLine(Scan.WaveLength[i] + "," + Scan.Reflectance[i]);
                }
                sw.WriteLine("##END=");

                sw.Flush();  // Clear buffer
                sw.Close();  // Close file stream 
            }
        }
        private void SaveHeader(FileStream fs, out StreamWriter sw, Boolean ifJCAMP)
        {
            sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            String TmpStrScan = String.Empty, TmpStrRef = String.Empty, PreStr = String.Empty;
            UInt16 TotalScanPtns = 0, TotalRefPtns = 0;

            String ModelName = Device.DevInfo.ModelName;
            String TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                           + Device.DevInfo.TivaRev[1].ToString() + "."
                           + Device.DevInfo.TivaRev[2].ToString() + "."
                           + Device.DevInfo.TivaRev[3].ToString();
            String DLPCRev = Device.DevInfo.DLPCRev[0].ToString() + "."
                           + Device.DevInfo.DLPCRev[1].ToString() + "."
                           + Device.DevInfo.DLPCRev[2].ToString();
            String UUID = BitConverter.ToString(Device.DevInfo.DeviceUUID).Replace("-", ":");
            String HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

            if (ifJCAMP == true)
            {
                PreStr = "##";

                sw.WriteLine("##TITLE=" + Scan.ScanConfigData.head.config_name);
                sw.WriteLine("##JCAMP-DX=4.24");
                sw.WriteLine("##DATA TYPE=INFRARED SPECTRUM");
            }
            else
            {
                PreStr = String.Empty;
            }

            TmpStrScan = Scan.ScanConfigData.head.config_name;
            TmpStrRef = (RadioButton_RefFac.Checked == true) ? "Built-In Reference" : "User Reference";
            sw.WriteLine(PreStr + "Method:," + TmpStrScan + "," + TmpStrRef + ",,,Model Name," + ModelName);

            TmpStrScan = Scan.ScanDateTime[2] + "/" + Scan.ScanDateTime[1] + "/" + Scan.ScanDateTime[0] + " @ " +
                         Scan.ScanDateTime[3] + ":" + Scan.ScanDateTime[4] + ":" + Scan.ScanDateTime[5];
            TmpStrRef = Scan.ReferenceScanDateTime[2] + "/" + Scan.ReferenceScanDateTime[1] + "/" + Scan.ReferenceScanDateTime[0] + " @ " +
                        Scan.ReferenceScanDateTime[3] + ":" + Scan.ReferenceScanDateTime[4] + ":" + Scan.ReferenceScanDateTime[5];
            String version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            sw.WriteLine(PreStr + "Host Date-Time:," + TmpStrScan + "," + TmpStrRef + ",,,GUI Version," + version.Substring(0, version.LastIndexOf('.')));

            sw.WriteLine(PreStr + "Header Version:," + Scan.ScanDataVersion + "," + Scan.ReferenceScanDataVersion + ",,,Tiva Version," + TivaRev);

            sw.WriteLine(PreStr + "System Temp (C):," + Scan.SensorData[0] + "," + Scan.ReferenceSensorData[0] + ",,,DLPC Version," + DLPCRev);

            sw.WriteLine(PreStr + "Detector Temp (C)," + Scan.SensorData[1] + "," + Scan.ReferenceSensorData[1] + ",,,UUID," + UUID);

            sw.WriteLine(PreStr + "Humidity (%):," + Scan.SensorData[2] + "," + Scan.ReferenceSensorData[2] + ",,,Main Board Version," + HWRev);

            sw.WriteLine(PreStr + "Lamp PD:," + Scan.SensorData[3] + "," + Scan.ReferenceSensorData[3]);

            sw.WriteLine(PreStr + "Shift Vector Coefficients:," + Device.Calib_Coeffs.ShiftVectorCoeffs[0] + "," +
                                                                  Device.Calib_Coeffs.ShiftVectorCoeffs[1] + "," +
                                                                  Device.Calib_Coeffs.ShiftVectorCoeffs[2]);

            sw.WriteLine(PreStr + "Pixel to Wavelength Coefficients:," + Device.Calib_Coeffs.PixelToWavelengthCoeffs[0] + "," +
                                                                         Device.Calib_Coeffs.PixelToWavelengthCoeffs[1] + "," +
                                                                         Device.Calib_Coeffs.PixelToWavelengthCoeffs[2]);

            sw.WriteLine(PreStr + "Serial Number:," + Scan.ScanConfigData.head.ScanConfig_serial_number + "," +
                                                      Scan.ReferenceScanConfigData.head.ScanConfig_serial_number);

            sw.WriteLine(PreStr + "Scan Config Name:," + Scan.ScanConfigData.head.config_name + "," +
                                                         Scan.ReferenceScanConfigData.head.config_name);

            if (Scan.ScanConfigData.head.num_sections == 1)
            {
                TmpStrScan = Helper.ScanTypeIndexToMode(Scan.ScanConfigData.section[0].section_scan_type);
                TmpStrRef = Helper.ScanTypeIndexToMode(Scan.ReferenceScanConfigData.section[0].section_scan_type);
            }
            else
            {
                TmpStrScan = Helper.ScanTypeIndexToMode(Scan.ScanConfigData.head.scan_type);
                TmpStrRef = Helper.ScanTypeIndexToMode(Scan.ReferenceScanConfigData.head.scan_type);
            }
            sw.WriteLine(PreStr + "Scan Config Type:," + TmpStrScan + "," + TmpStrRef);

            for (Int32 i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
            {
                if (Scan.ScanConfigData.head.num_sections > 1)
                {
                    TmpStrScan = Helper.ScanTypeIndexToMode(Scan.ScanConfigData.section[i].section_scan_type);
                    TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Helper.ScanTypeIndexToMode(Scan.ReferenceScanConfigData.section[i].section_scan_type) : String.Empty;
                    sw.WriteLine(PreStr + "Section " + (i + 1) + "," + TmpStrScan + "," + TmpStrRef);
                }

                TmpStrScan = Scan.ScanConfigData.section[i].wavelength_start_nm.ToString();
                TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Scan.ReferenceScanConfigData.section[i].wavelength_start_nm.ToString() : String.Empty;
                sw.WriteLine(PreStr + "Start wavelength (nm):," + TmpStrScan + "," + TmpStrRef);

                TmpStrScan = Scan.ScanConfigData.section[i].wavelength_end_nm.ToString();
                TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Scan.ReferenceScanConfigData.section[i].wavelength_end_nm.ToString() : String.Empty;
                sw.WriteLine(PreStr + "End wavelength (nm):," + TmpStrScan + "," + TmpStrRef);

                TmpStrScan = Math.Round(Helper.CfgWidthPixelToNM(Scan.ScanConfigData.section[i].width_px), 2).ToString();
                TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Math.Round(Helper.CfgWidthPixelToNM(Scan.ReferenceScanConfigData.section[i].width_px), 2).ToString() : String.Empty;
                sw.WriteLine(PreStr + "Pattern Pixel Width (nm):," + TmpStrScan + "," + TmpStrRef);

                TmpStrScan = Helper.CfgExpIndexToTime(Scan.ScanConfigData.section[i].exposure_time).ToString();
                TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Helper.CfgExpIndexToTime(Scan.ReferenceScanConfigData.section[i].exposure_time).ToString() : String.Empty;
                sw.WriteLine(PreStr + "Exposure (ms):," + TmpStrScan + "," + TmpStrRef);

                TmpStrScan = Scan.ScanConfigData.section[i].num_patterns.ToString();
                TmpStrRef = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Scan.ReferenceScanConfigData.section[i].num_patterns.ToString() : String.Empty;
                sw.WriteLine(PreStr + "Digital Resolution:," + TmpStrScan + "," + TmpStrRef);
            }

            for (Int32 i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
            {
                TotalScanPtns += Scan.ScanConfigData.section[i].num_patterns;
            }
            for (Int32 i = 0; i < Scan.ReferenceScanConfigData.head.num_sections; i++)
            {
                TotalRefPtns += Scan.ReferenceScanConfigData.section[i].num_patterns;
            }
            if (ifJCAMP == true)
            {
                sw.WriteLine("##NPOINTS=" + TotalScanPtns);
            }
            else
            {
                sw.WriteLine("Total Digital Resolution:," + TotalScanPtns + "," + TotalRefPtns);
            }

            sw.WriteLine(PreStr + "Num Repeats:," + Scan.ScanConfigData.head.num_repeats + "," + Scan.ReferenceScanConfigData.head.num_repeats);

            sw.WriteLine(PreStr + "PGA Gain:," + Scan.PGA + "," + Scan.ReferencePGA);

            TimeSpan ts = new TimeSpan(TimeScanEnd.Ticks - TimeScanStart.Ticks);
            sw.WriteLine(PreStr + "Total Measurement Time in sec:," + ts.TotalSeconds);
        }
        private void SaveHeader_CSV(StreamWriter sw)
        {
            String ModelName = Device.DevInfo.ModelName;
            String TivaRev = String.Empty;
            String DLPCRev = String.Empty;
            String SpecLibRev = String.Empty;
            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_3 || GetFW_LEVEL() == FW_LEVEL.LEVEL_1)
            {
                TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                        + Device.DevInfo.TivaRev[1].ToString() + "."
                        + Device.DevInfo.TivaRev[2].ToString();
            }
            else
            {
                TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                        + Device.DevInfo.TivaRev[1].ToString() + "."
                        + Device.DevInfo.TivaRev[2].ToString() + "."
                        + Device.DevInfo.TivaRev[3].ToString();
            }
            if(Device.DevInfo.DLPCRev.Length == 3)
            {
                DLPCRev = Device.DevInfo.DLPCRev[0].ToString() + "."
                          + Device.DevInfo.DLPCRev[1].ToString() + "."
                          + Device.DevInfo.DLPCRev[2].ToString();
            }
            else
            {
                DLPCRev = Device.DevInfo.DLPCRev[0].ToString() + "."
                        + Device.DevInfo.DLPCRev[1].ToString() + "."
                        + Device.DevInfo.DLPCRev[2].ToString() + "."
                        + Device.DevInfo.DLPCRev[3].ToString();
            }
            if (Device.DevInfo.SpecLibRev[3].ToString() == "0")
            {
                SpecLibRev = Device.DevInfo.SpecLibRev[0].ToString() + "."
                           + Device.DevInfo.SpecLibRev[1].ToString() + "."
                           + Device.DevInfo.SpecLibRev[2].ToString();
            }
            else
            {
                SpecLibRev = Device.DevInfo.SpecLibRev[0].ToString() + "."
                           + Device.DevInfo.SpecLibRev[1].ToString() + "."
                           + Device.DevInfo.SpecLibRev[2].ToString() + "."
                           + Device.DevInfo.SpecLibRev[3].ToString();
            }
            String UUID = BitConverter.ToString(Device.DevInfo.DeviceUUID).Replace("-", ":");
            String HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;
            String Detector_Board_HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(2, 1) : String.Empty;
            String Manufacturing_SerNum = Device.DevInfo.Manufacturing_SerialNumber;
            //--------------------------------------------------------
            String Data_Date_Time = String.Empty, Ref_Config_Name = String.Empty, Ref_Data_Date_Time = String.Empty;
            String Section_Config_Type = String.Empty, Ref_Section_Config_Type = String.Empty;
            String Pattern_Width = String.Empty, Ref_Pattern_Width = String.Empty;
            String Exposure = String.Empty, Ref_Exposure = String.Empty;
            //----------------------------------------------

            String[,] CSV = new String[29, 15];
            for (int i = 0; i < 29; i++)
                for (int j = 0; j < 15; j++)
                    CSV[i, j] = ",";

            // Section information field names
            CSV[0, 0] = "***Scan Config Information***,";
            CSV[0, 7] = "***Reference Scan Information***";
            CSV[17, 0] = "***General Information***,";
            CSV[17, 7] = "***Calibration Coefficients***";
            CSV[28, 0] = "***Scan Data***";
            // Config field names & values
            for (int i = 0; i < 2; i++)
            {
                CSV[1, i * 7] = "Scan Config Name:,";
                CSV[2, i * 7] = "Scan Config Type:,";
                CSV[2, i * 7 + 2] = "Num Section:,";
                CSV[3, i * 7] = "Section Config Type:,";
                CSV[4, i * 7] = "Start Wavelength (nm):,";
                CSV[5, i * 7] = "End Wavelength (nm):,";
                CSV[6, i * 7] = "Pattern Width (nm):,";
                CSV[7, i * 7] = "Exposure (ms):,";
                CSV[8, i * 7] = "Digital Resolution:,";
                CSV[9, i * 7] = "Num Repeats:,";
                CSV[10, i * 7] = "PGA Gain:,";
                CSV[11, i * 7] = "System Temp (C):,";
                CSV[12, i * 7] = "Humidity (%):,";
                if (HWRev == "F")
                {
                    CSV[13, i * 7] = "Lamp ADC:,";
                }
                else
                   CSV[13, i * 7] = "Lamp PT:,";
                CSV[14, i * 7] = "Data Date-Time:,";
            }

            for (int i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
            {
                if (i == 0)
                {
                    // Scan config values
                    CSV[1, 1] = Scan.ScanConfigData.head.config_name + ",";
                    CSV[2, 1] = "Slew,";
                    CSV[2, 3] = Scan.ScanConfigData.head.num_sections + ",";
                    CSV[9, 1] = Scan.ScanConfigData.head.num_repeats + ",";
                    CSV[10, 1] = Scan.PGA + ",";
                    CSV[11, 1] = Scan.SensorData[0] + ",";
                    CSV[12, 1] = Scan.SensorData[2] + ",";
                    CSV[13, 1] = Scan.SensorData[3] + ",";
                    if(HWRev == "F")
                    {
                        CSV[13,2] = Scan.SensorData[1]*100 + ",";
                    }

                    Data_Date_Time = Scan.ScanDateTime[2] + "/" + Scan.ScanDateTime[1] + "/" + Scan.ScanDateTime[0] + " @ " +
                                 Scan.ScanDateTime[3] + ":" + Scan.ScanDateTime[4] + ":" + Scan.ScanDateTime[5];
                    CSV[14, 1] = Data_Date_Time + ",";
                    //Reference config values
                    Ref_Config_Name = (RadioButton_RefFac.Checked == true) ? "Built-In Reference" : "User Reference";
                    if (RadioButton_RefFac.Checked == true)
                    {
                        if (Scan.ReferenceScanConfigData.head.config_name == "SystemTest")
                        {
                            Ref_Config_Name = "Built-in Factory Reference,";
                        }
                        else
                        {
                            Ref_Config_Name = "Built-in User Reference,";
                        }
                    }
                    else
                    {
                        Ref_Config_Name = "Local New Reference,";
                    }
                    CSV[1, 8] = Ref_Config_Name + ",";
                    CSV[2, 8] = "Slew,";
                    CSV[2, 10] = Scan.ReferenceScanConfigData.head.num_sections + ",";
                    CSV[9, 8] = Scan.ReferenceScanConfigData.head.num_repeats + ",";
                    CSV[10, 8] = Scan.ReferencePGA + ",";
                    CSV[11, 8] = Scan.ReferenceSensorData[0] + ",";
                    CSV[12, 8] = Scan.ReferenceSensorData[2] + ",";
                    CSV[13, 8] = Scan.ReferenceSensorData[3] + ",";

                    Ref_Data_Date_Time = Scan.ReferenceScanDateTime[2] + "/" + Scan.ReferenceScanDateTime[1] + "/" + Scan.ReferenceScanDateTime[0] + " @ " +
                           Scan.ReferenceScanDateTime[3] + ":" + Scan.ReferenceScanDateTime[4] + ":" + Scan.ReferenceScanDateTime[5];
                    CSV[14, 8] = Ref_Data_Date_Time + ",";
                }
                // Scan config section values
                Section_Config_Type = Helper.ScanTypeIndexToMode(Scan.ScanConfigData.section[i].section_scan_type);
                CSV[3, i + 1] = Section_Config_Type + ",";
                CSV[4, i + 1] = Scan.ScanConfigData.section[i].wavelength_start_nm.ToString() + ",";
                CSV[5, i + 1] = Scan.ScanConfigData.section[i].wavelength_end_nm.ToString() + ",";

                Pattern_Width = Math.Round(Helper.CfgWidthPixelToNM(Scan.ScanConfigData.section[i].width_px), 2).ToString();
                CSV[6, i + 1] = Pattern_Width + ",";

                Exposure = Helper.CfgExpIndexToTime(Scan.ScanConfigData.section[i].exposure_time).ToString();
                CSV[7, i + 1] = Exposure + ",";
                CSV[8, i + 1] = Scan.ScanConfigData.section[i].num_patterns.ToString() + ",";

                // Reference config section values
                if (i < Scan.ReferenceScanConfigData.head.num_sections)
                {
                    Ref_Section_Config_Type = Helper.ScanTypeIndexToMode(Scan.ReferenceScanConfigData.section[i].section_scan_type);
                    CSV[3, i + 8] = Ref_Section_Config_Type + ",";

                    CSV[4, i + 8] = Scan.ReferenceScanConfigData.section[i].wavelength_start_nm.ToString() + ",";
                    CSV[5, i + 8] = Scan.ReferenceScanConfigData.section[i].wavelength_end_nm.ToString() + ",";

                    Ref_Pattern_Width = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Math.Round(Helper.CfgWidthPixelToNM(Scan.ReferenceScanConfigData.section[i].width_px), 2).ToString() : String.Empty;
                    CSV[6, i + 8] = Ref_Pattern_Width + ",";

                    Ref_Exposure = (Scan.ReferenceScanConfigData.head.num_sections > 1 || i == 0) ? Helper.CfgExpIndexToTime(Scan.ReferenceScanConfigData.section[i].exposure_time).ToString() : String.Empty;
                    CSV[7, i + 8] = Ref_Exposure + ",";
                    CSV[8, i + 8] = Scan.ReferenceScanConfigData.section[i].num_patterns.ToString() + ",";
                }
            }

            // Measure Time field name & value
            CSV[15, 0] = "Total Measurement Time in sec:,";
            TimeSpan ts = new TimeSpan(TimeScanEnd.Ticks - TimeScanStart.Ticks);
            CSV[15, 1] = ts.TotalSeconds.ToString() + ",";

            // Coefficients filed names & valus
            CSV[18, 7] = "Shift Vector Coefficients:,";
            CSV[18, 8] = Device.Calib_Coeffs.ShiftVectorCoeffs[0].ToString() + ",";
            CSV[18, 9] = Device.Calib_Coeffs.ShiftVectorCoeffs[1].ToString() + ",";
            CSV[18, 10] = Device.Calib_Coeffs.ShiftVectorCoeffs[2].ToString() + ",";
            CSV[19, 7] = "Pixel to Wavelength Coefficients:,";
            CSV[19, 8] = Device.Calib_Coeffs.PixelToWavelengthCoeffs[0].ToString() + ",";
            CSV[19, 9] = Device.Calib_Coeffs.PixelToWavelengthCoeffs[1].ToString() + ",";
            CSV[19, 10] = Device.Calib_Coeffs.PixelToWavelengthCoeffs[2].ToString() + ",";

            // General information field names & values
            CSV[18, 0] = "Model Name:,";
            CSV[18, 1] = ModelName + ",";
            CSV[19, 0] = "Serial Number:,";
            CSV[19, 1] = Device.DevInfo.SerialNumber + ",";
            CSV[19, 2] = "(" + Manufacturing_SerNum + "),";
            CSV[20, 0] = "GUI Version:,";
            String GUIRev = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            GUIRev = GUIRev.Substring(0, GUIRev.LastIndexOf('.'));
            CSV[20, 1] = GUIRev;
            CSV[21, 0] = "TIVA Version:,";
            CSV[21, 1] = TivaRev + ",";
            CSV[22, 0] = "DLPC Version:,";
            CSV[22, 1] = DLPCRev + ",";
            CSV[23, 0] = "DLPSPECLIB Version:,";
            CSV[23, 1] = SpecLibRev + ",";
            CSV[24, 0] = "UUID:,";
            CSV[24, 1] = UUID + ",";
            CSV[25, 0] = "Main Board Version:,";
            CSV[26, 0] = "Detector Board Version:,";
            CSV[25, 1] = HWRev + ",";
            CSV[26, 1] = Detector_Board_HWRev + ",";

            string buf = "";
            for (int i = 0; i < 29; i++)
            {
                for (int j = 0; j < 15; j++)
                {
                    buf += CSV[i, j];
                    if (j == 14)
                    {
                        sw.WriteLine(buf);
                    }
                }
                buf = "";
            }
        }
        #endregion
        #region startScan
        private void Button_Scan_Click(object sender, EventArgs e)
        {
            Button_CfgCancel_Click(this, e);
            if (Device.IsConnected())
            {
                try
                {
                    TargetScanCounts = int.Parse(Text_ContScan.Text);
                }catch(Exception e1)
                {
                    Message.ShowError("Invalid Cont. Scan value.");
                    Text_ContScan.Text = "1";
                    UserCancelScan = true;
                    Button_Scan.Text = "Scan";
                    DBG.WriteLine(e1.Message);
                    return;
                };

                UserCancelScan = false;  // Clear this flag before scanning

                if (CheckBox_SaveOneCSV.Checked)
                    SaveOneCSVFile = true;

                if (RadioButton_RefNew.Checked || TargetScanCounts == 0) TargetScanCounts = 1;
                Text_ContScan.Text = (TargetScanCounts - ScannedCounts).ToString();

                Scan.SetScanNumRepeats(ushort.Parse(textBox_ScanAvg.Text.ToString()));

                if (CheckBox_AutoGain.Checked == false)
                {                    
                    if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2 && CheckBox_LampOn.Checked == true)
                        Scan.SetPGAGain(GetPGA());
                    else
                        Scan.SetFixedPGAGain(true, GetPGA());
                }
                else
                {
                    if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2)
                        Scan.SetFixedPGAGain(false, GetPGA());
                    else
                        Scan.SetFixedPGAGain(true, 0); // This is set to auto PGA
                }

                if (bwScan.IsBusy != true)
                    bwScan.RunWorkerAsync();
                else
                {
                    String text = "Scanning in progress...\n\nPlease wait!";
                    MessageBox.Show(text, "Wait");
                }
                Label_ContScan.Text = string.Empty;
                if (!Check_Overlay.Checked)
                {
                    Chart_Refresh();
                }
                ProgressWindowStart("Scan", "Start Scan,Please Wait");
            }
            else
            {
                String text = "Please connect a device before performing scan!";
                MessageBox.Show(text, "Warning");
            }
        }
        private byte GetPGA()
        {
            byte pga = 64;
            if(ComboBox_PGAGain.Text.ToString().Equals("64"))
            {
                pga = 64;
            }
            else if(ComboBox_PGAGain.Text.ToString().Equals("32"))
            {
                pga = 32;
            }
            else if (ComboBox_PGAGain.Text.ToString().Equals("16"))
            {
                pga = 16;
            }
            else if (ComboBox_PGAGain.Text.ToString().Equals("8"))
            {
                pga = 8;
            }
            else if (ComboBox_PGAGain.Text.ToString().Equals("4"))
            {
                pga = 4;
            }
            else if (ComboBox_PGAGain.Text.ToString().Equals("2"))
            {
                pga = 2;
            }
            else if (ComboBox_PGAGain.Text.ToString().Equals("1"))
            {
                pga = 1;
            }
            return pga;
        }

        private void bwScan_DoScan(object sender, DoWorkEventArgs e)
        {
            DBG.WriteLine("Performing scan... Remained scans: {0}", TargetScanCounts - ScannedCounts);
            TimeScanStart = DateTime.Now;
            List<object> arguments = new List<object>();

            if (Scan.PerformScan(ReferenceSelect) == 0)
            {
                DBG.WriteLine("Scan completed!");
                TimeScanEnd = DateTime.Now;
                TimeSpan ts = new TimeSpan(TimeScanEnd.Ticks - TimeScanStart.Ticks);

                arguments.Add("pass");
                arguments.Add(ts);
                e.Result = arguments;
            }
            else
            {
                arguments.Add("failed");
                arguments.Add(TimeSpan.Zero);
                e.Result = arguments;
            }
        }

        private void bwScan_DoSacnCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<object> arguments = e.Result as List<object>;
            string result = (string)arguments[0];
            TimeSpan ts = (TimeSpan)arguments[1];
            Byte pga = Scan.PGA;  // If PGA is auto, it can only read the current value after scanning.

            ScannedCounts++;
            Text_ContScan.Text = (TargetScanCounts - ScannedCounts).ToString();
            Label_ContScan.Text = "(" + ScannedCounts.ToString() + "/" + TargetScanCounts.ToString() + ")";

            if (result != "failed")
            {
                SpectrumPlot();

                ComboBox_PGAGain.SelectedItem = pga.ToString();
                Label_ScanStatus.Text = "Total Scan Time: " + String.Format("{0:0.000}", ts.TotalSeconds) + " secs.";

                if (ReferenceSelect != Scan.SCAN_REF_TYPE.SCAN_REF_NEW)  // Save scan results except new reference selection
                    SaveToFiles();
                else if (Scan.IsLocalRefExist)
                {
                    Scan.GetScanResult();
                    Byte[] time = Scan.ReferenceScanDateTime;
                    if (time[0] != 0)
                    {
                        pre_ref_time = "Previous reference last set on : 20" + time[0].ToString() + "/" + time[1].ToString() + "/" + time[2].ToString()
                        + " @ " + time[3].ToString() + ":" + time[4].ToString() + ":" + time[5].ToString();
                    }
                    RadioButton_RefPre.Enabled = true;
                    RadioButton_RefPre.Checked = true;
                    RadioButton_RefNew.Checked = false;
                   
                }

                if ((TargetScanCounts - ScannedCounts) > 0 && !UserCancelScan)
                {
                    // The Thread.Sleep() will cause GUI stopping execution. Using timer to avoid this situation.
                    if (int.TryParse(Text_ContDelay.Text, out int DelaySec) == false)
                        DelaySec = 0;

                    DateTime ScanCurrent = DateTime.Now;
                    while (DateTime.Now < ScanCurrent.AddSeconds(DelaySec))
                    {
                        Application.DoEvents();
                    }

                    if (CheckBox_AutoGain.Checked == false)
                    {
                        if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2 && CheckBox_LampOn.Checked == true)
                            Scan.SetPGAGain(pga);
                        else
                            Scan.SetFixedPGAGain(true, pga);
                    }
                    else
                    {
                        if (GetFW_LEVEL() < FW_LEVEL.LEVEL_2)
                            Scan.SetFixedPGAGain(false, pga);
                        else
                            Scan.SetFixedPGAGain(true, 0); // This is set to auto PGA
                    }

                    bwScan.RunWorkerAsync();
                }
                else
                {
                    UserCancelScan = true;
                    ScannedCounts = 0;
                    Button_Scan.Text = "Scan";
                    ProgressWindowCompleted();
                }
            }
            else
            {
                String text = "Scan Failed!";
                UserCancelScan = true;
                Button_Scan.Text = "Scan";
                MessageBox.Show(text, "Error");
                ProgressWindowCompleted();
            }
        }
        #endregion
        //------------------------------------------------------------------------------------
        #region Utility
        #region Model Name
        private void Button_ModelNameGet_Click(object sender, EventArgs e)
        {
            StringBuilder pOutBuf = new StringBuilder(128);

            if (Device.ReadModelName(pOutBuf) == 0)
                TextBox_ModelName.Text = pOutBuf.ToString();
            else
                TextBox_ModelName.Text = "Read Failed!";

            pOutBuf.Clear();
        }

        private void Button_ModelNameSet_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to write it?","Model Name",MessageBoxButtons.YesNo);
            if(result == DialogResult.Yes)
            {
                if (Device.SetModelName(Helper.CheckRegex(TextBox_ModelName.Text)) == 0)
                {
                    if (Device.Information() != 0)
                        DBG.WriteLine("Device Information read failed!");
                    GetDeviceInfo();
                    if (!String.IsNullOrEmpty(Device.DevInfo.ModelName))
                        TextBox_ModelName.Text = Device.DevInfo.ModelName;
                    else
                        TextBox_ModelName.Text = "Read Failed!";

                    CheckLampFuncUseful();
                }
                else
                    TextBox_ModelName.Text = "Write Failed!";
            }
        }
        #endregion
        #region Serial Number
        //Serial Number
        private void Button_SerialNumberSet_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to write it?", "Serial Number", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (Device.SetSerialNumber(Helper.CheckRegex(TextBox_SerialNumber.Text)) == 0)
                {
                    if (Device.Information() != 0)
                        DBG.WriteLine("Device Information read failed!");
                    GetDeviceInfo();
                    if (!String.IsNullOrEmpty(Device.DevInfo.SerialNumber))
                        TextBox_SerialNumber.Text = Device.DevInfo.SerialNumber;
                    else
                        TextBox_SerialNumber.Text = "Read Failed!";
                }
                else
                    TextBox_SerialNumber.Text = "Write Failed!";
            }  
        }

        private void Button_SerialNumberGet_Click(object sender, EventArgs e)
        {
            StringBuilder pOutBuf = new StringBuilder(128);

            if (Device.GetSerialNumber(pOutBuf) == 0)
                TextBox_SerialNumber.Text = pOutBuf.ToString();
            else
                TextBox_SerialNumber.Text = "Read Failed!";

            pOutBuf.Clear();
        }
        #endregion
        #region Date and Time
        //Date and Time
        private void Button_DateTimeSync_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to sync. it?", "Date and Time", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Device.DeviceDateTime DevDateTime = new Device.DeviceDateTime();
                DateTime Current = DateTime.Now;

                DevDateTime.Year = Current.Year;
                DevDateTime.Month = Current.Month;
                DevDateTime.Day = Current.Day;
                DevDateTime.DayOfWeek = (Int32)Current.DayOfWeek;
                DevDateTime.Hour = Current.Hour;
                DevDateTime.Minute = Current.Minute;
                DevDateTime.Second = Current.Second;

                if (Device.SetDateTime(DevDateTime) == 0)
                    TextBox_DateTime.Text = Current.ToString("yyyy/M/d  H:m:s");
                else
                    TextBox_DateTime.Text = "Sync Failed!";
            }
        }

        private void Button_DateTimeGet_Click(object sender, EventArgs e)
        {
            if (Device.GetDateTime() == 0)
            {
                TextBox_DateTime.Text = Device.DevDateTime.Year + "/"
                                      + Device.DevDateTime.Month + "/"
                                      + Device.DevDateTime.Day + "  "
                                      + Device.DevDateTime.Hour + ":"
                                      + Device.DevDateTime.Minute + ":"
                                      + Device.DevDateTime.Second;
            }
            else
                TextBox_DateTime.Text = "Get Failed!";
        }
        #endregion
        #region Lamp Usage
        //Lamp Usage
        private String GetLampUsage()
        {
            String lampusage = "";
            UInt64 buf = Device.LampUsage / 1000;

            if (buf / 86400 != 0)
            {
                lampusage += buf / 86400 + "day ";
                buf -= 86400 * (buf / 86400);
            }
            if (buf / 3600 != 0)
            {
                lampusage += buf / 3600 + "hr ";
                buf -= 3600 * (buf / 3600);
            }
            if (buf / 60 != 0)
            {
                lampusage += buf / 60 + "min ";
                buf -= 60 * (buf / 60);
            }
            lampusage += buf + "sec ";
            return lampusage;
        }

        private void Button_LampUsageSet_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to write it?", "Lamp Usage", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (Double.TryParse(TextBox_LampUsage.Text, out Double LampUsage) == false)
                {
                    TextBox_LampUsage.Text = "Not Numeric!";
                    return;
                }

                if (Device.WriteLampUsage((UInt64)(LampUsage * 3600000)) == 0)  // hour to milliseconds
                    Button_LampUsageGet_Click(sender, e);
                else
                    TextBox_LampUsage.Text = "Write Failed!";
                GetDeviceInfo();
            }
        }

        private void Button_LampUsageGet_Click(object sender, EventArgs e)
        {
            if (Device.ReadLampUsage() == 0)
                TextBox_LampUsage.Text = ((Double)Device.LampUsage / 3600000).ToString();  // milliseconds to hour
            else
                TextBox_LampUsage.Text = "Read Failed!";
        }
        #endregion
        #region Sensors
        //Sensors
        private void Button_SensorRead_Click(object sender, EventArgs e)
        {
            if (Device.ReadSensorsData() == 0)
            {
                Label_SensorBattStatus.Text = Device.DevSensors.BattStatus;
                Label_SensorBattCapacity.Text = (Device.DevSensors.BattCapicity != -1) ? (Device.DevSensors.BattCapicity.ToString() + " %") : ("Read Failed!");
                Label_SensorHumidity.Text = (Device.DevSensors.Humidity != -1) ? (Device.DevSensors.Humidity.ToString() + " %") : ("Read Failed!");
                Label_SensorSysTemp.Text = (Device.DevSensors.HDCTemp != -1) ? (Device.DevSensors.HDCTemp.ToString() + " C") : ("Read Failed!");
                Label_SensorTivaTemp.Text = (Device.DevSensors.TivaTemp != -1) ? (Device.DevSensors.TivaTemp.ToString() + " C") : ("Read Failed!");
                Label_SensorLampVM1Value.Text = (Device.DevSensors.PhotoDetector != -1) ? (Device.DevSensors.PhotoDetector.ToString()) : ("Read Failed!");
            }
            else
            {
                Label_SensorBattStatus.Text = "Read Failed!";
                Label_SensorBattCapacity.Text = "Read Failed!";
                Label_SensorHumidity.Text = "Read Failed!";
                Label_SensorSysTemp.Text = "Read Failed!";
                Label_SensorTivaTemp.Text = "Read Failed!";
                Label_SensorLampVM1Value.Text = "Read Failed!";
            }

            String Model = (!String.IsNullOrEmpty(Device.DevInfo.ModelName)) ? Device.DevInfo.ModelName : String.Empty;
            Model = (Model != String.Empty) ? Model.Substring(Model.LastIndexOf('-') + 1, 1) : String.Empty;

            // Convert main board version to ASCII
            Byte[] HWRev = Encoding.ASCII.GetBytes(Device.DevInfo.HardwareRev);
            Int32 MB_Ver = HWRev[0];

            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_4 && Model != "F" && Model != "f")
            {
                Scan.SetLamp(Scan.LAMP_CONTROL.ON_SCAN);

                if (MB_Ver >= 'F')
                {
                    if (Device.ReadLampParam() == SDK.PASS)
                    {
                        Double voltage, current;

                        voltage = (Double)Device.LampADC[0] / 4096 * 3.3 * 2;
                        current = (Double)Device.LampADC[2] / 4096 * 3.3 / 50 / 0.1 * 1000;
                        Label_SensorLampVM1Value.Text = String.Format("{0:0.00} V", voltage);
                        Label_SensorLampCM1Value.Text = String.Format("{0:0.00} mA", current);

                        if (Model == "R")
                        {
                            voltage = (Double)Device.LampADC[1] / 4096 * 3.3 * 2;
                            current = (Double)Device.LampADC[3] / 4096 * 3.3 / 50 / 0.1 * 1000;
                            Label_SensorLampVM2Value.Text = String.Format("{0:0.00} V", voltage);
                            Label_SensorLampCM2Value.Text = String.Format("{0:0.00} mA", current);
                        }
                    }
                    else
                    {
                        Label_SensorLampVM1Value.Text = "Read Failed!";
                        Label_SensorLampCM1Value.Text = "Read Failed!";
                        if (Model == "R")
                        {
                            Label_SensorLampVM2Value.Text = "Read Failed!";
                            Label_SensorLampCM2Value.Text = "Read Failed!";
                        }
                    }
                }
                else
                {
                    if (Device.ReadLampParam() == SDK.PASS)
                    {
                        Label_SensorLampVM1Value.Text = Device.LampADC[0].ToString();
                        Label_SensorLampCM1Value.Text = String.Empty;
                        Label_SensorLampVM2Value.Text = String.Empty;
                        Label_SensorLampCM2Value.Text = String.Empty;
                    }
                    else
                    {
                        Label_SensorLampVM1Value.Text = "Read Failed!";
                        Label_SensorLampCM1Value.Text = String.Empty;
                        Label_SensorLampVM2Value.Text = String.Empty;
                        Label_SensorLampCM2Value.Text = String.Empty;
                    }
                }

                Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);
            }
        }
        #endregion
        #region Tiva FW update
        //Tiva FW update
        private void Button_TivaFWBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                InitialDirectory = (Tiva_FWDir == String.Empty) ? (Directory.GetCurrentDirectory()) : (Tiva_FWDir),
                FileName = "",                  // Default file name
                DefaultExt = ".bin",            // Default file extension
                Filter = "Binary File|*.bin"    // Filter files by extension
            };

            // Show open file dialog box
            dlg.ShowDialog();
            // Process open file dialog box results
            if (dlg.FileName != "")
            {
                TextBox_TivaFWPath.Text = dlg.FileName;
                Tiva_FWDir = dlg.FileName.Substring(0, dlg.FileName.LastIndexOf("\\"));
                ControlSingleControl(Button_TivaFWUpdate, true);
            }
        }
        private void Button_TivaFWUpdate_Click(object sender, EventArgs e)
        {
            SDK.AutoSearch = false;
            SDK.IsEnableNotify = false;
            SDK.IsConnectionChecking = false;

            String filePath = (String)TextBox_TivaFWPath.Text;

            if (Device.IsConnected() && filePath != "")
            {
                int Ret = SDK.PASS;
                int retry = 0;
                TimerCallback callback = new TimerCallback(TimerTask);
                TivaUpdateTime = 1;
                timer = new System.Threading.Timer(callback, null, 1000, 1000);

                ProgressBar_TivaFWUpdateStatus.Value = 10;
                if (Device.ReadDeviceStatus() == 0 && (Device.DeviceStatus & 0x00000001) == 1 && (Device.DeviceStatus & 0x00000002) == 0)
                    Device.Set_Tiva_To_Bootloader();

                while (!Device.IsDFUConnected())
                {
                    if (++retry > 50)
                    {
                        Ret = SDK.FAIL;
                        break;
                    }
                    Thread.Sleep(100);
                }

                if (Ret == SDK.PASS)
                {
                    List<object> arguments = new List<object> { filePath };
                    bwTivaUpdate.RunWorkerAsync(arguments);
                }
                else
                {
                    SDK.AutoSearch = true;
                    SDK.IsEnableNotify = true;
                    MessageBox.Show("Can not find \"Tiva DFU\"!", "Error");
                    SDK.IsConnectionChecking = true;
                    ProgressBar_TivaFWUpdateStatus.Value = 0;
                    timer.Dispose();
                }

            }
            else if (Device.IsDFUConnected())
            {
                TimerCallback callback = new TimerCallback(TimerTask);
                TivaUpdateTime = 1;
                timer = new System.Threading.Timer(callback, null, 1000, 1000);
                List<object> arguments = new List<object> { filePath };
                bwTivaUpdate.RunWorkerAsync(arguments);
            }
            else
            {
                SDK.AutoSearch = true;
                SDK.IsEnableNotify = true;
                MessageBox.Show("Device dose not exist or image file path error!", "Error");
                SDK.IsConnectionChecking = true;
                ProgressBar_TivaFWUpdateStatus.Value = 0;
            }
        }

        private int pValue = 30;
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            if (pValue < 99)
            {
                pValue += 1;
                bwTivaUpdate.ReportProgress(pValue);
            }
        }

        private void bwTivaUpdate_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> arguments = e.Argument as List<object>;
            String filePath = (String)arguments[0];
            bwTivaUpdate.ReportProgress(30);

            pValue = 30;
            System.Timers.Timer pTimer = new System.Timers.Timer(100);
            pTimer.Elapsed += OnTimedEvent;
            pTimer.AutoReset = true;
            pTimer.Enabled = true;

            int ret = Device.Tiva_FW_Update(filePath);
            if (ret == 0)
                e.Result = 0;
            else
                e.Result = ret;

            pTimer.Enabled = false;
        }

        private void bwTivaUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int percentage = e.ProgressPercentage;
            ProgressBar_TivaFWUpdateStatus.Value = percentage;
        }

        private void bwTivaUpdate_DoSacnCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            int ret = (int)e.Result;
            ProgressBar_TivaFWUpdateStatus.Value = 100;

            if (ret == 0)
            {
                timer.Dispose();
                String text = "Tiva FW updated successfully!";
                MessageBox.Show(text, "Success");
                Device.ResetTiva(true);  // Only reconnect the device
            }
            else
            {
                timer.Dispose();
                switch (ret)
                {
                    case -1:
                        String text = "The driver, lmdfu.dll, for the USB Device Firmware Upgrade device cannot be found!";
                        MessageBox.Show(text, "Error");
                        break;
                    case -2:
                        text = "The driver for the USB Device Firmware Upgrade device was found but appears to be a version which this program does not support!";
                        MessageBox.Show(text, "Error");
                        break;
                    case -3:
                        text = "An error was reported while attempting to load the device driver for the USB Device Firmware Upgrade device!";
                        MessageBox.Show(text, "Error");
                        break;
                    case -4:
                        text = "Unable to open binary file.Copy binary file to a folder with Admin / read / write permission and try again.";
                        MessageBox.Show(text, "Error");
                        break;
                    case -5:
                        text = "Memory alloc for file read failed!";
                        MessageBox.Show(text, "Error");
                        break;
                    case -6:
                        text = "This image does not appear to be valid for the target device.";
                        MessageBox.Show(text, "Error");
                        break;
                    case -7:
                        text = "This image is not valid NIRNANO FW Image!";
                        MessageBox.Show(text, "Error");
                        break;
                    case -8:
                        text = "Error reported during file download!";
                        MessageBox.Show(text, "Error");
                        break;
                    default:
                        text = "Unknown error occured!";
                        MessageBox.Show(text, "Error");
                        break;
                }
            }
            SDK.AutoSearch = true;
            SDK.IsEnableNotify = true;
            SDK.IsConnectionChecking = true;
            ProgressBar_TivaFWUpdateStatus.Value = 0;
        }
        #endregion
        #region DLPC150 FW Update
        //DLPC150 FW Update
        private void Button_DLPC150FWBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                InitialDirectory = (DLPC_FWDir == String.Empty) ? (Directory.GetCurrentDirectory()) : (DLPC_FWDir),
                FileName = "",              // Default file name
                DefaultExt = ".img",        // Default file extension
                Filter = "Image File|*.img" // Filter files by extension
            };

            dlg.ShowDialog();
            if (dlg.FileName != "")
            {
                TextBox_DLPC150FWPath.Text = dlg.FileName;
                DLPC_FWDir = dlg.FileName.Substring(0, dlg.FileName.LastIndexOf("\\"));
            }
        }
        private void Button_DLPC150FWUpdate_Click(object sender, EventArgs e)
        {
            if (Device.IsConnected() && TextBox_DLPC150FWPath.Text != "")
            {
                bwDLPCUpdate.RunWorkerAsync(TextBox_DLPC150FWPath.Text);
            }
            else
            {
                String text = "Device dose not exist or image file path error!";
                MessageBox.Show(text, "Error");
            }
        }

        private void bwDLPCUpdate_DoWork(object sender, DoWorkEventArgs e)
        {
            int expectedChecksum = 0, chksum = 0, ret = 0;
            String fileName = (String)e.Argument;
            byte[] imgByteBuff = File.ReadAllBytes(fileName);
            e.Result = false;

            int dataLen = imgByteBuff.Length;

            if (!Device.DLPC_CheckSignature(imgByteBuff))
            {
                DBG.WriteLine("Invalid DLPC150 image file!");
                return;
            }

            ret = Device.DLPC_SetImageSize(dataLen);
            if (ret < 0)
            {
                DBG.WriteLine("Set DLPC150 image size failed! (error: {0})", ret);
                return;
            }

            for (int i = 0; i < dataLen; i++)
            {
                expectedChecksum += imgByteBuff[i];
            }

            Thread.Sleep(1000);

            int bytesToSend = dataLen, bytesSent = 0;
            while (bytesToSend > 0)
            {
                byte[] byteArrayToSent = new byte[bytesToSend];
                Buffer.BlockCopy(imgByteBuff, dataLen - bytesToSend, byteArrayToSent, 0, bytesToSend);

                bytesSent = Device.DLPC_FW_Update_WriteData(byteArrayToSent, bytesToSend);

                if (bytesSent < 0)
                {
                    DBG.WriteLine("DLPC150 update: Data send Failed!");
                    break;
                }

                bytesToSend -= bytesSent;

                // Report the FW update status
                float updateProgress;
                updateProgress = ((float)(dataLen - bytesToSend) / dataLen) * 100;
                bwDLPCUpdate.ReportProgress((int)updateProgress);
            }

            chksum = Device.DLPC_Get_Checksum();

            if (chksum < 0)
                DBG.WriteLine("Error Reading DLPC150 Flash Checksum! (error: {0})", chksum);
            else if (chksum != expectedChecksum)
                DBG.WriteLine("Checksum mismatched: (Expected: {0}, DLPC Flash: {1})", expectedChecksum, chksum);
            else
            {
                DBG.WriteLine("DLPC150 updated successfully!");
                e.Result = true;
            }
        }

        private void bwDLPCUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int percentage = e.ProgressPercentage;
            ProgressBar_DLPC150FWUpdateStatus.Value = e.ProgressPercentage;
        }

        private void bwDLPCUpdate_DoWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if ((bool)e.Result)
            {
                String text = "DLPC150 FW updated successfully!";
                MessageBox.Show(text, "Success");
            }
            else
            {
                String text = "DLPC150 FW update failed!";
                MessageBox.Show(text, "Error");
            }

            ProgressBar_DLPC150FWUpdateStatus.Value = 0;
            bwTivaReset = new BackgroundWorker
            {
                WorkerReportsProgress = false,
                WorkerSupportsCancellation = true
            };
            bwTivaReset.DoWork += new DoWorkEventHandler(bwTivaReset_DoWork);
            bwTivaReset.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwTivaReset_DoWorkCompleted);

            SDK.IsConnectionChecking = false;
            bwTivaReset.RunWorkerAsync();
        }
        #endregion
        #region Calibration Coefficients
        //Calibration Coefficients
        private void Button_CalWriteGenCoeffs_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to write it?", "Generic Coefficients", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (Device.SetGenericCalibStruct() == SDK.PASS)
                    Button_CalReadCoeffs_Click(sender, e);
                else
                {
                    Message.ShowError("Write Failed!");
                }
            }
           
        }

        private void Button_CalReadCoeffs_Click(object sender, EventArgs e)
        {
            if (Device.GetCalibStruct() == SDK.PASS)
            {
                Label_CalCoeffVer.Text = Device.DevInfo.CalRev.ToString();
                Label_RefCalVer.Text = Device.DevInfo.RefCalRev.ToString();
                Label_ScanCfgVer.Text = Device.DevInfo.CfgRev.ToString();
                TextBox_P2WCoeff0.Text = Device.Calib_Coeffs.PixelToWavelengthCoeffs[0].ToString();
                TextBox_P2WCoeff1.Text = Device.Calib_Coeffs.PixelToWavelengthCoeffs[1].ToString();
                TextBox_P2WCoeff2.Text = Device.Calib_Coeffs.PixelToWavelengthCoeffs[2].ToString();
                TextBox_ShiftVectCoeff0.Text = Device.Calib_Coeffs.ShiftVectorCoeffs[0].ToString();
                TextBox_ShiftVectCoeff1.Text = Device.Calib_Coeffs.ShiftVectorCoeffs[1].ToString();
                TextBox_ShiftVectCoeff2.Text = Device.Calib_Coeffs.ShiftVectorCoeffs[2].ToString();
            }
            else
            {
                Label_CalCoeffVer.Text = "0";
                Label_RefCalVer.Text = "0";
                Label_ScanCfgVer.Text = "0";
                TextBox_P2WCoeff0.Text = "Read Failed!";
                TextBox_P2WCoeff1.Text = "Read Failed!";
                TextBox_P2WCoeff2.Text = "Read Failed!";
                TextBox_ShiftVectCoeff0.Text = "Read Failed!";
                TextBox_ShiftVectCoeff1.Text = "Read Failed!";
                TextBox_ShiftVectCoeff2.Text = "Read Failed!";
            }
        }

        private void Button_CalRestoreDefaultCoeffs_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to restore it?", "Coefficient", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                int ret = Device.RestoreDefaultCalibStruct();
                if (ret == 0)
                    Button_CalReadCoeffs_Click(sender, e);
                else if (ret == -4)
                {
                    Message.ShowError("Device does not have backup data!");
                }
                else
                {
                    Message.ShowError("Restore Failed!");
                }
            }  
        }

        private void Button_CalWriteCoeffs_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to write it?", "Coefficient", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Device.CalibCoeffs Calib_Coeffs = new Device.CalibCoeffs
                {
                    PixelToWavelengthCoeffs = new Double[3],
                    ShiftVectorCoeffs = new Double[3]
                };

                if ((Double.TryParse(TextBox_P2WCoeff0.Text, out Calib_Coeffs.PixelToWavelengthCoeffs[0]) == false) ||
                    (Double.TryParse(TextBox_P2WCoeff1.Text, out Calib_Coeffs.PixelToWavelengthCoeffs[1]) == false) ||
                    (Double.TryParse(TextBox_P2WCoeff2.Text, out Calib_Coeffs.PixelToWavelengthCoeffs[2]) == false) ||
                    (Double.TryParse(TextBox_ShiftVectCoeff0.Text, out Calib_Coeffs.ShiftVectorCoeffs[0]) == false) ||
                    (Double.TryParse(TextBox_ShiftVectCoeff1.Text, out Calib_Coeffs.ShiftVectorCoeffs[1]) == false) ||
                    (Double.TryParse(TextBox_ShiftVectCoeff2.Text, out Calib_Coeffs.ShiftVectorCoeffs[2]) == false))
                {
                    Message.ShowError("Not Numeric!");
                    return;
                }

                if (Device.SendCalibStruct(Calib_Coeffs) == SDK.PASS)
                {
                    Button_CalReadCoeffs_Click(sender, e);
                    //should set config to update DMD pattern
                    if (DevCurCfg_IsTarget)
                    {
                        SetScanConfig(ScanConfig.TargetConfig[DevCurCfg_Index], true, DevCurCfg_Index);
                    }
                    else
                    {
                        SetScanConfig(LocalConfig[DevCurCfg_Index], false, DevCurCfg_Index);                  
                    }
                }             
                else
                {
                    Message.ShowError("Write Failed!");
                }
            }
              
        }

        private void CheckBox_CalWriteEnable_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox_CalWriteEnable.Checked == true)
            {
                Button_CalWriteCoeffs.Enabled = true;
                Button_CalWriteGenCoeffs.Enabled = true;
                if (IsActivated && GetFW_LEVEL() >= FW_LEVEL.LEVEL_2)
                    Button_CalRestoreDefaultCoeffs.Enabled = true;
                else
                    Button_CalRestoreDefaultCoeffs.Enabled = false;
            }
            else
            {
                Button_CalWriteCoeffs.Enabled = false;
                Button_CalWriteGenCoeffs.Enabled = false;
                Button_CalRestoreDefaultCoeffs.Enabled = false;
            }
        }
        #endregion
        #region Device Information
        //Device Information
        private void GetDeviceInfo()
        {
            if (!Device.IsConnected())
                return;
            String GUIRev = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            GUIRev = GUIRev.Substring(0, GUIRev.LastIndexOf('.'));

            String DLPCRev = "";
            DLPCRev = Device.DevInfo.DLPCRev[0].ToString() + "."
                    + Device.DevInfo.DLPCRev[1].ToString() + "."
                    + Device.DevInfo.DLPCRev[2].ToString();

            String SpecLibRev = String.Empty;
            if (Device.DevInfo.SpecLibRev[3].ToString() == "0")
            {
                SpecLibRev = Device.DevInfo.SpecLibRev[0].ToString() + "."
                           + Device.DevInfo.SpecLibRev[1].ToString() + "."
                           + Device.DevInfo.SpecLibRev[2].ToString();
            }
            else
            {
                SpecLibRev = Device.DevInfo.SpecLibRev[0].ToString() + "."
                           + Device.DevInfo.SpecLibRev[1].ToString() + "."
                           + Device.DevInfo.SpecLibRev[2].ToString() + "."
                           + Device.DevInfo.SpecLibRev[3].ToString();
            }
            String HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;
            String Detector_Board_HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(2, 1) : String.Empty;

            label_DevInfoGUIVer.Text = GUIRev;
            label_DevInfoDLPCVer.Text = DLPCRev;
            label_DevInfoSpecLibVer.Text = SpecLibRev;
            label_DevInfoMainBoardVer.Text = HWRev;
            label_DevInfoDetectorBoardVer.Text = Detector_Board_HWRev;
            label_DevInfoModelName.Text = Device.DevInfo.ModelName;
            label_DevInfoDevSerNum.Text = Device.DevInfo.SerialNumber;

            String TivaRev = String.Empty;
            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_3)
            {
                TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                        + Device.DevInfo.TivaRev[1].ToString() + "."
                        + Device.DevInfo.TivaRev[2].ToString();
            }
            else
            {
                if(Device.DevInfo.TivaRev[3] == 0)
                {
                    TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                            + Device.DevInfo.TivaRev[1].ToString() + "."
                            + Device.DevInfo.TivaRev[2].ToString();
                }
                else
                {
                    TivaRev = Device.DevInfo.TivaRev[0].ToString() + "."
                            + Device.DevInfo.TivaRev[1].ToString() + "."
                            + Device.DevInfo.TivaRev[2].ToString() + "."
                            + Device.DevInfo.TivaRev[3].ToString();
                }
            }
            label_DevInfoTivaSWVer.Text = TivaRev;

            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_2)
            {
                String Manu_Seri_Num = Device.DevInfo.Manufacturing_SerialNumber;
                if (!Manu_Seri_Num.Contains("70UB1") && !Manu_Seri_Num.Contains("95UB1"))
                    Manu_Seri_Num = "NA";
                label_DevInfoManfacSerNum.Text = Manu_Seri_Num;

                String Lamp_Usage = "";
                if (Device.ReadLampUsage() == 0)
                    Lamp_Usage = GetLampUsage();
                else
                    Lamp_Usage = "NA";
                label_DevInfoLampUsageValue.Text = Lamp_Usage;
            }
            else
            {
                label_DevInfoManfacSerNum.Text = String.Empty;
                label_DevInfoLampUsageValue.Text = String.Empty;
            }

            if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_1)
            {
                String UUID = BitConverter.ToString(Device.DevInfo.DeviceUUID).Replace("-", ":");
                label_DevInfoUUID.Text = UUID;
            }
            else
            {
                label_DevInfoUUID.Text = String.Empty;
            }
        }

        private void CheckLampFuncUseful()
        {
            if (!Device.IsConnected())
                return;

            String Model = (!String.IsNullOrEmpty(Device.DevInfo.ModelName)) ? Device.DevInfo.ModelName : String.Empty;
            Model = (Model != String.Empty) ? Model.Substring(Model.LastIndexOf('-') + 1, 1) : String.Empty;

            if (Model == "F" || Model == "f")
            {
                GroupBox_LampUsage.Visible = false;
                label_DevInfoLampUsage.Visible = false;
                label_DevInfoLampUsageValue.Visible = false;
                Label_SensorLampVM1.Visible = false;
                Label_SensorLampVM1Value.Visible = false;
                Label_SensorLampVM2.Visible = false;
                Label_SensorLampVM2Value.Visible = false;
                Label_SensorLampCM1.Visible = false;
                Label_SensorLampCM1Value.Visible = false;
                Label_SensorLampCM2.Visible = false;
                Label_SensorLampCM2Value.Visible = false;
            }
            else
            {
                GroupBox_LampUsage.Visible = true;
                label_DevInfoLampUsage.Visible = true;
                label_DevInfoLampUsageValue.Visible = true;
                Label_SensorLampVM1.Visible = true;
                Label_SensorLampVM1Value.Visible = true;
                Label_SensorLampCM1.Visible = true;
                Label_SensorLampCM1Value.Visible = true;
                if (Model == "R")
                {
                    Label_SensorLampVM2.Visible = true;
                    Label_SensorLampVM2Value.Visible = true;
                    Label_SensorLampCM2.Visible = true;
                    Label_SensorLampCM2Value.Visible = true;
                }
                else
                {
                    Label_SensorLampVM2.Visible = false;
                    Label_SensorLampVM2Value.Visible = false;
                    Label_SensorLampCM2.Visible = false;
                    Label_SensorLampCM2Value.Visible = false;
                }

                if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_2)
                {
                    if (IsActivated)
                    {
                        GroupBox_LampUsage.Enabled = true;
                        label_DevInfoLampUsage.Enabled = true;
                        label_DevInfoLampUsageValue.Enabled = true;
                    }
                    else
                    {
                        GroupBox_LampUsage.Enabled = false;
                        label_DevInfoLampUsage.Enabled = false;
                        label_DevInfoLampUsageValue.Enabled = false;
                    }

                    // Convert main board version to ASCII
                    Byte[] HWRev = Encoding.ASCII.GetBytes(Device.DevInfo.HardwareRev);
                    Int32 MB_Ver = HWRev[0];

                    if (GetFW_LEVEL() >= FW_LEVEL.LEVEL_4 && MB_Ver >= 'F')
                    {
                        if (Model == "R")
                        {
                            Label_SensorLampVM1.Text = "Lamp 1 Voltage";
                            Label_SensorLampCM1.Text = "Lamp 1 Current";
                            Label_SensorLampVM2.Text = "Lamp 2 Voltage";
                            Label_SensorLampCM2.Text = "Lamp 2 Current";
                        }
                        else
                        {
                            Label_SensorLampVM1.Text = "Lamp Voltage";
                            Label_SensorLampCM1.Text = "Lamp Current";
                            Label_SensorLampVM2.Text = String.Empty;
                            Label_SensorLampCM2.Text = String.Empty;
                        }
                    }
                    else
                    {
                        Label_SensorLampVM1.Text = "Lamp Intensity";
                        Label_SensorLampCM1.Text = String.Empty;
                        Label_SensorLampVM2.Text = String.Empty;
                        Label_SensorLampCM2.Text = String.Empty;
                    }
                    Label_SensorLampVM1Value.Text = String.Empty;
                    Label_SensorLampCM1Value.Text = String.Empty;
                    Label_SensorLampVM2Value.Text = String.Empty;
                    Label_SensorLampCM2Value.Text = String.Empty;
               }
            }
        }
        #endregion

        #region Activation Key
        //Activation Key
        private void button_KeySet_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to set it?", "Key", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                String[] StrKey = textBox_Key.Text.Split(new char[] { ' ', ':', ';', '-', '_' });
                Byte[] ByteKey = new Byte[12];

                for (int i = 0; i < StrKey.Length; i++)
                {
                    try { ByteKey[i] = Convert.ToByte(StrKey[i], 16); }
                    catch { ByteKey[i] = 0; }
                }

                Device.SetActivationKey(ByteKey);

                if (IsActivated)
                {
                    label_ActivateStatus.Text = "Activated!";
                    GUI_Handler((int)MainWindow.GUI_State.KEY_ACTIVATE);
                }
                else
                {
                    label_ActivateStatus.Text = "Not activated!";
                    GUI_Handler((int)MainWindow.GUI_State.KEY_NOT_ACTIVATE);
                }

                CheckLampFuncUseful();
            }        
        }

        private void GetActivationKeyStatus()
        {
            if (IsActivated)
            {
                label_ActivateStatus.Text = "Activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_ACTIVATE);
            }
            else
            {
                label_ActivateStatus.Text = "Not activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_NOT_ACTIVATE);
            }
        }

        private void button_KeyClear_Click(object sender, EventArgs e)
        {
            String buf = "000000000000";
            String[] StrKey = buf.Split(new char[] { ' ', ':', ';', '-', '_' });
            Byte[] ByteKey = new Byte[12];

            for (int i = 0; i < StrKey.Length; i++)
            {
                try { ByteKey[i] = Convert.ToByte(StrKey[i], 16); }
                catch { ByteKey[i] = 0; }
            }

            Device.SetActivationKey(ByteKey);

            if (IsActivated)
            {
                label_ActivateStatus.Text = "Activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_ACTIVATE);
            }
            else
            {
                label_ActivateStatus.Text = "Not activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_NOT_ACTIVATE);
            }

            CheckLampFuncUseful();
        }

        private void button_manage_Click(object sender, EventArgs e)
        {
            ActiveKeyManageForm window = new ActiveKeyManageForm { Owner = this };
            window.ShowDialog();
            if (IsActivated)
            {
                label_ActivateStatus.Text = "Activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_ACTIVATE);
            }
            else
            {
                label_ActivateStatus.Text = "Not activated!";
                GUI_Handler((int)MainWindow.GUI_State.KEY_NOT_ACTIVATE);
            }

            CheckLampFuncUseful();
        }
        #endregion

        #region Reset Device
        //Device
        //Reset Device
        private void button_DeviceResetSys_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to reset it?", "Reset System", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (!Device.IsConnected())
                    return;

                bwTivaReset = new BackgroundWorker
                {
                    WorkerReportsProgress = false,
                    WorkerSupportsCancellation = true
                };
                bwTivaReset.DoWork += new DoWorkEventHandler(bwTivaReset_DoWork);
                bwTivaReset.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwTivaReset_DoWorkCompleted);

                SDK.IsConnectionChecking = false;
                bwTivaReset.RunWorkerAsync();
            }         
        }
        private BackgroundWorker bwTivaReset;
        private static void bwTivaReset_DoWork(object sender, DoWorkEventArgs e)
        {
            int ret = Device.ResetTiva(false);
        }
        private static void bwTivaReset_DoWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SDK.IsConnectionChecking = true;
        }
        #endregion
        #region Update Reference Data
        //Update Reference Data
        private void button_DeviceUpdateRef_Click(object sender, EventArgs e)
        {
            if (Device.IsConnected())
            {
                DialogResult result = Message.ShowQuestion("IMPORTANT!!!\n\nThis will REPLACE your FACTORY REFERENCE DATA \nand could NOT be REVERTED.\n\nAre you sure you want to do this ? ",null, MessageBoxButtons.YesNo);
                if(result == DialogResult.Yes)
                {
                    result = Message.ShowQuestion("User Agreements:\n\n" +
                    "1. I am well aware of the purpose of factory reference data\n" +
                    "    and have been well trained to replace it.\n" +
                    "2. I fully understand that the factory reference data can be replaced\n" +
                    "    but not revertible.\n" +
                    "3. I agree to pay extra fee to recover the factory reference data\n" +
                    "    if I make anything wrong.\n\n" +
                    "I agree with above terms and would like to continue the process.\n"
                    ,null, MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        result = Message.ShowQuestion("IMPORTANT!!!\n\nPlease confirm again with this process.\n\nDo you still want to do this?",null, MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            result = Message.ShowQuestion("Please place the reference sample and press 'OK' to start the reference scan...",null, MessageBoxButtons.OKCancel);
                            if (result == DialogResult.OK)
                            {
                                bwRefScanProgress = new BackgroundWorker
                                {
                                    WorkerReportsProgress = false,
                                    WorkerSupportsCancellation = true
                                };
                                bwRefScanProgress.DoWork += new DoWorkEventHandler(bwRefScanProgress_DoWork);
                                bwRefScanProgress.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwRefScanProgress_DoWorkCompleted);
                                bwRefScanProgress.RunWorkerAsync();
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
                
            }
            else
            {
                String text = "No device is connected!";
                MessageBox.Show(text, "Warning");
            }
        }

        public ScanConfig.SlewScanConfig tmpCfg;//backup current config before update reference
        private void bwRefScanProgress_DoWork(object sender, DoWorkEventArgs e)
        {
            Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);
            tmpCfg = ScanConfig.GetCurrentConfig();
            ScanConfig.SlewScanConfig scanCfg = new ScanConfig.SlewScanConfig();
            scanCfg.head.config_name = "UserReference";
            scanCfg.head.scan_type = 2;
            scanCfg.head.num_sections = 1;
            scanCfg.head.num_repeats = 30;
            scanCfg.section = new ScanConfig.SlewScanSection[5];
            scanCfg.section[0].section_scan_type = 0;
            scanCfg.section[0].wavelength_start_nm = 900;
            scanCfg.section[0].wavelength_end_nm = 1700;
            scanCfg.section[0].width_px = 6;
            scanCfg.section[0].num_patterns = 228;
            scanCfg.section[0].exposure_time = 0;

            int ret = ScanConfig.SetScanConfig(scanCfg);
            if (ret != 0)
            {
                e.Result = -3;
                return;
            }

            Thread.Sleep(200);
            ret = Scan.PerformScan(Scan.SCAN_REF_TYPE.SCAN_REF_NEW);
            if (ret == 0)
            {
                ret = Scan.SaveReferenceScan();
                if (ret == 0)
                    e.Result = 0;
                else
                    e.Result = -2;
            }
            else
                e.Result = -1;
        }

        private void bwRefScanProgress_DoWorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            int ret = (int)e.Result;
            if (ret == 0)
            {
                String text = "Reference Scan Completed Seccessfully!\n\nPlease start a new scan to check the result.";
                MessageBox.Show(text, "Success");
                GetBuildInRefTime();
                if(RadioButton_RefFac.Checked)
                {
                    RadioButton_RefFac_CheckedChanged(null, null);
                }
            }
            else if (ret == -1)
            {
                String text = "Scan Failed!";
                MessageBox.Show(text, "Error");
            }
            else if (ret == -2)
            {
                String text = "Save Reference Sacn Failed!";
                MessageBox.Show(text, "Error");
            }
            else if (ret == -3)
            {
                String text = "Set Reference Sacn Configuration Failed!";
                MessageBox.Show(text, "Error");
            }
            else
            {
                String text = "Unknow Error Occured!";
                MessageBox.Show(text, "Error");
            }
            ScanConfig.SetScanConfig(tmpCfg);//set current config after update reference
        }
        #endregion
        #region Back up factory reference
        //Back up factory reference
        private void DeviceConnectBackUpRef()
        {
            if (Device.IsConnected())
            {
                int ret;
                string serNum = Device.DevInfo.SerialNumber.ToString();
                ret = Device.Backup_Factory_Reference(serNum);
                if (ret < 0)
                {
                    switch (ret)
                    {
                        case -1:
                            label11.Text = "(Out of memory!)";
                            label11.ForeColor = System.Drawing.Color.Red;
                            button_DeviceBackUpFacRef.Enabled = true;
                            break;
                        case -2:
                            label11.Text = "(System I/O error!)";
                            label11.ForeColor = System.Drawing.Color.Red;
                            button_DeviceBackUpFacRef.Enabled = true;
                            break;
                        case -3:
                            label11.Text = "(Device communcation error!)";
                            label11.ForeColor = System.Drawing.Color.Red;
                            button_DeviceBackUpFacRef.Enabled = true;
                            break;
                        case -4:
                            label11.Text = "(Device does not have the original factory reference data!)";
                            label11.ForeColor = System.Drawing.Color.Red;
                            button_DeviceBackUpFacRef.Enabled = false;
                            break;
                    }
                }
                else
                {
                    button_DeviceBackUpFacRef.Enabled = true;
                }
            }
            else
            {
                label11.ForeColor = System.Drawing.Color.Red;
                label11.Text = "No device connected for backup factory reference!";
                button_DeviceBackUpFacRef.Enabled = false;
            }
        }
        private void button_DeviceBackUpFacRef_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to backup it?", "Backup Factory Reference", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                DeviceConnectBackUpRef();
            }
        }
        private void button_DeviceRestoreFacRef_Click(object sender, EventArgs e)
        {
            DialogResult result = Message.ShowQuestion("Do you want to restore it?", "Restore Factory Reference", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (Device.IsConnected())
                {
                    int ret;
                    string serNum = Device.DevInfo.SerialNumber.ToString();
                    ret = Device.Restore_Factory_Reference(serNum);
                    if (ret < 0)
                    {
                        switch (ret)
                        {
                            case -1:
                                String text = "Factory reference data restore FAILED!\n\nOut of memory.";
                                MessageBox.Show(text, "Error");
                                break;
                            case -2:
                                text = "Factory reference data restore FAILED!\n\nBackup directory not found";
                                MessageBox.Show(text, "Error");
                                break;
                            case -3:
                                text = "Factory reference data restore FAILED!\n\nRead file error";
                                MessageBox.Show(text, "Error");
                                break;
                            case -4:
                                text = "Factory reference data restore FAILED!\n\nReference data currupted";
                                MessageBox.Show(text, "Error");
                                break;
                            case -5:
                                text = "Factory reference data restore FAILED!\n\nDevice communcation error";
                                MessageBox.Show(text, "Error");
                                break;
                            case -6:
                                text = "Factory reference data restore FAILED!\n\nData was NOT the original factory reference data";
                                MessageBox.Show(text, "Error");
                                break;
                        }
                    }
                    else
                    {
                        String text = "Factory reference data has been restored successfully!\n\nPlease start a new scan to check the result.";
                        MessageBox.Show(text, "Success");
                        GetBuildInRefTime();
                        if (RadioButton_RefFac.Checked)
                        {
                            RadioButton_RefFac_CheckedChanged(null, null);
                        }
                        //ClearScanPlotsEvent();
                    }
                }
                else
                {
                    String text = "No device connected for restoring factory reference!";
                    MessageBox.Show(text, "Warning");
                }
            }
               
        }
        #endregion

        #region Button Lock/Unlock

        private void Button_LockButton_Click(object sender, EventArgs e)
        {
            if (Device.SetButtonLock(true) == SDK.PASS)
            {
                Int32 status = Device.GetButtonLockStatus();
                if (status == 1)
                    Label_ButtonStatus.Text = "Button Status: Locked!";
                else if (status == 0)
                    Label_ButtonStatus.Text = "Button Status: Unlocked!";
                else
                    Label_ButtonStatus.Text = "Button Status: Read Failed!";
            }
            else
            {
                Label_ButtonStatus.Text = "Button Status: Lock Failed!";
            }
        }

        private void Button_UnlockButton_Click(object sender, EventArgs e)
        {
            if (Device.SetButtonLock(false) == SDK.PASS)
            {
                Int32 status = Device.GetButtonLockStatus();
                if (status == 1)
                    Label_ButtonStatus.Text = "Button Status: Locked!";
                else if (status == 0)
                    Label_ButtonStatus.Text = "Button Status: Unlocked!";
                else
                    Label_ButtonStatus.Text = "Button Status: Read Failed!";
            }
            else
            {
                Label_ButtonStatus.Text = "Button Status: Unlock Failed!";
            }
        }

        #endregion

        #region BLE Advertising Name

        private void Button_Clear_BLE_Display_Name_Click(object sender, EventArgs e)
        {
            if (Device.WriteBleDispName() == SDK.PASS)
                Button_Get_BLE_Display_Name_Click(sender, e);
            else
                TextBox_BLE_Display_Name.Text = "Clear Failed!";
        }

        private void Button_Set_BLE_Display_Name_Click(object sender, EventArgs e)
        {
            String BLE_Name = TextBox_BLE_Display_Name.Text;

            DialogResult result = Message.ShowQuestion("Do you want to write it?", "Bluetooth LE Advertising Name", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                Regex rgx = new Regex("[^a-zA-Z0-9 <>[]{}_-]");

                if (rgx.IsMatch(BLE_Name))
                {
                    Message.ShowError("BLE Name can only be alpha numeric characters, please enter again in the correct format!");
                    TextBox_BLE_Display_Name.Text = String.Empty;
                    return;
                }
                else if (BLE_Name.Length > 24)
                {
                    Message.ShowError("The max. BLE Name length should be less than 24 characters!");
                    TextBox_BLE_Display_Name.Text = BLE_Name.Substring(0, 24);
                    return;
                }

                if (Device.WriteBleDispName(BLE_Name) == SDK.PASS)
                    Button_Get_BLE_Display_Name_Click(sender, e);
                else
                    TextBox_BLE_Display_Name.Text = "Write Failed!";
            }
        }

        private void Button_Get_BLE_Display_Name_Click(object sender, EventArgs e)
        {
            StringBuilder pOutBuf = new StringBuilder(128);

            if (Device.ReadBleDispName(pOutBuf) == SDK.PASS)
                TextBox_BLE_Display_Name.Text = pOutBuf.ToString();
            else
                TextBox_BLE_Display_Name.Text = "Read Failed!";

            pOutBuf.Clear();
        }

        #endregion

        #region About
        //About
        private void button_AboutLicense_Click(object sender, EventArgs e)
        {
            LicenseWindow window = new LicenseWindow { Owner = this };
            window.ShowDialog();
        }

        private void button_About_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start("http://www.inno-spectra.com/");
            }
            catch { }
        }

        #endregion

        #endregion

        private void Button_SaveDirChange_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                fbd.SelectedPath = Scan_Dir;
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    TextBox_SaveDirPath.Text = fbd.SelectedPath;
                    Scan_Dir = fbd.SelectedPath;
                }
            }
        }

        private void tabScanPage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabScanPage.SelectedIndex == 2)
            {
                Check_Overlay.Enabled = true;
            }
            else if (tabScanPage.SelectedIndex == 0 && RadioButton_RefPre.Checked && !Scan.IsLocalRefExist)
            {
                RadioButton_RefNew.PerformClick();
            }
            Button_CfgCancel_Click(this, e);
        }

        private void UI_no_connection()
        {
            ControlAllControls(this, false);
            ControlSingleControl(panel_Saved_Scan, true);
            ControlPanelContents(panel_Saved_Scan, true);
            ControlSingleControl(MyChart, true);
            ControlSingleControl(RadioButton_Reflectance, true);
            ControlSingleControl(RadioButton_Absorbance, true);
            ControlSingleControl(RadioButton_Intensity, true);
            ControlSingleControl(RadioButton_Reference, true);
            ControlSingleControl(label123, true);
            ControlSingleControl(label122, true);
            ControlSingleControl(button_AboutLicense, true);
            ControlSingleControl(button_About, true);
            ControlSingleControl(label6, true);
            ControlSingleControl(TextBox_TivaFWPath,true);
            ControlSingleControl(Button_TivaFWBrowse, true);
            ListBox_LocalCfgs.BackColor = System.Drawing.Color.White;
            ListBox_TargetCfgs.BackColor = System.Drawing.Color.White;
        }

        private void UI_on_connection()
        {
            ControlAllControls(this, true);
            if (RadioButton_RefNew.Checked)
            {
                GroupBox_ContScan.Enabled = false;
            }
            if(!RadioButton_RefNew.Checked &&(int.TryParse(Text_ContScan.Text, out int repeat) && repeat > 1) )
            {
                CheckBox_SaveOneCSV.Enabled = true;
            }
            else
            {
                CheckBox_SaveOneCSV.Enabled = false;
            }

            CheckLampFuncUseful();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Button_CfgCancel_Click(this, e);
            GetDeviceInfo();
            if(tabControl1.SelectedIndex!=0)
            {
                Scan.SetLamp(Scan.LAMP_CONTROL.OFF_SCAN);
                Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);
            }
            else if(tabControl1.SelectedIndex ==0)
            {
                CheckBox_LampOn.Checked = false;
                RadioButton_LampStableTime.Checked = true;
                Scan.SetLamp(Scan.LAMP_CONTROL.AUTO);
            }
        }

        private void ControlAllControls(Control con, bool enable)
        {
            foreach (Control c in con.Controls)
            {
                ControlAllControls(c, enable);
            }
            con.Enabled = enable;
        }

        private void ControlSingleControl(Control con, bool enable)
        {
            if (con != null)
            {
                con.Enabled = enable;
                ControlSingleControl(con.Parent, enable);
            }
        }

        private void ControlPanelContents(Panel panel, bool enabled)
        {
            foreach (Control ctrl in panel.Controls)
            {
                ctrl.Enabled = enabled;
            }
        }
        #region TIVA FW update timer
        private static int TivaUpdateTime = 1;
        static System.Threading.Timer timer;
        private static void TimerTask(object obj)
        {
            TivaUpdateTime++;
            if (TivaUpdateTime >= 30)
            {
                timer.Dispose();
                if (Device.IsConnected())
                    Device.ResetTiva(true);
                else
                    Device.ResetTiva(null);
                MessageBox.Show("Tiva FW Update timeout!", "Tiva FW Update");
            }
        }
        #endregion

        private void Check_Overlay_CheckedChanged(object sender, EventArgs e)
        {
            if (Check_Overlay.Checked)
                MyChart.DisableAnimations = true;
            else
                MyChart.DisableAnimations = false;

            Chart_Refresh();
            SpectrumPlot();
        }
        private void Chart_Refresh()
        {
            MyChart.Series.Clear();
            MyChart.AxisX.Clear();
            MyChart.AxisY.Clear();

            MyChart.Series.Add(new LineSeries
            {
                Values = new ChartValues<ObservablePoint>(),
                Title = "Intensity",
                StrokeThickness = 1,
                Fill = null,
                LineSmoothness = 0,
                PointGeometry = null,
                PointGeometrySize = 0,
            });

            MyChart.AxisX.Add(new Axis { Title = "Wavelength (nm)" });
            MyChart.AxisY.Add(new Axis { Title = "Intensity" });
        }
        private void ListBox_DeviceScanConfig_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (IsCfgLegal(true) == SDK.FAIL)
            {
                Message.ShowError("Invalid Dig. Resolution value in configuration!");
                return;
            }
            if (TargetCfg_SelIndex>=0)
            {
                SetScanConfig(ScanConfig.TargetConfig[TargetCfg_SelIndex], true, TargetCfg_SelIndex);
                RefreshLocalCfgList();
                RefreshTargetCfgList();
                ListBox_TargetCfgs.SelectedIndex = DevCurCfg_Index;
            }           
        }
        #region Local Config
        private void LoadLocalCfgList()
        {
            String FileName = Path.Combine(ConfigDir, "ConfigList.xml");

            if (File.Exists(FileName) == true)
            {
                XmlSerializer xml = new XmlSerializer(typeof(List<ScanConfig.SlewScanConfig>));
                TextReader reader = new StreamReader(FileName);
                LocalConfig = (List<ScanConfig.SlewScanConfig>)xml.Deserialize(reader);
                reader.Close();
                RefreshLocalCfgList();
            }
            else
            {
                LocalConfig.Clear();
            }
        }
        private void RefreshLocalCfgList()
        {
            ListBox_LocalCfgs.Items.Clear();
            if (LocalConfig.Count > 0)
            {
                for (Int32 i = 0; i < LocalConfig.Count; i++)
                {
                    ListBox_LocalCfgs.Items.Add(LocalConfig[i].head.config_name);
                    /*if (DevCurCfg_IsTarget == false)
                        Item.Foreground = (DevCurCfg_Index == i) ? Brushes.DarkOrange : Brushes.Black;
                    Item.MouseDoubleClick += new System.Windows.Input.MouseButtonEventHandler(ListBoxItem_LocalConfig_MouseDoubleClick);
                    ListBox_LocalCfgs.Items.Add(Item);*/
                }
            }
        }


        private void ListBox_LocalCfgs_SelectedIndexChanged(object sender, EventArgs e)
        {
            isSelectConfig = true;
            if (NewConfig == true || EditConfig == true)
            {
                EditConfig = false;
                NewConfig = false;
                Button_CfgCancel_Click(this, e);
            }
            LocalCfg_Last_SelIndex = LocalCfg_SelIndex;
            LocalCfg_SelIndex = ListBox_LocalCfgs.SelectedIndex;
            if (LocalCfg_SelIndex < 0 || LocalConfig.Count == 0)
            {
                if (ListBox_TargetCfgs.SelectedIndex == -1 && ListBox_LocalCfgs.SelectedIndex == -1)//new config situation
                {
                    return;
                }
                else
                {
                    ListBox_LocalCfgs.BackColor = System.Drawing.Color.White;
                    ListBox_TargetCfgs.BackColor = System.Drawing.Color.AliceBlue;
                    Button_SetActive.Enabled = true;
                    return;
                }            
            }  
            else
            {
                if (!Scan.IsLocalRefExist && ReferenceSelect != Scan.SCAN_REF_TYPE.SCAN_REF_BUILT_IN)
                {
                    RadioButton_RefPre.Checked = false;
                    RadioButton_RefNew.Checked = true;
                }
            }

            SelCfg_IsTarget = false;
            FillCfgDetailsContent();
            OpenColseScanConfigButton(nameof(ScanConfigMode.INITIAL));
            // Clear target listbox index after local config data refreshed.
            if (ListBox_TargetCfgs.SelectedIndex != -1)
            {
                ListBox_TargetCfgs.SelectedIndex = -1;
            }  
            SetDetailColorWhite();
            isSelectConfig = false;
            IsCfgLegal(true);
        }
        private void ListBox_LocalCfgs_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (IsCfgLegal(true) == SDK.FAIL)
            {
                Message.ShowError("Invalid Dig. Resolution value in configuration!");
                return;
            } 
            if (LocalCfg_SelIndex>=0)
            {
                SetScanConfig(LocalConfig[LocalCfg_SelIndex], false, LocalCfg_SelIndex);
                RefreshLocalCfgList();
                RefreshTargetCfgList();
                ListBox_LocalCfgs.SelectedIndex = DevCurCfg_Index;
            }
        }
        private void Button_CopyCfgL2T_Click(object sender, EventArgs e)
        {
            int index = 0;
            if (LocalCfg_SelIndex < 0)
            {
                Message.ShowWarning("No item selected.");
                return;
            }

            if (ScanConfig.TargetConfig.Count >= 20)//Confirm the current number of device configuration before saving
            {
                Message.ShowWarning("Number of scan configs in device cannot exceed 20.");
                return;
            }
            if(CheckConfigName(ListBox_LocalCfgs.Items[ListBox_LocalCfgs.SelectedIndex].ToString(), false,ref index))
            {
                DialogResult Result = Message.ShowQuestion("The config has exist, do you want to overwrite?", null, MessageBoxButtons.YesNo);
                if (Result == DialogResult.No)
                {
                    return;
                }
                if (Result == DialogResult.Yes)
                {
                    ScanConfig.TargetConfig[index] = LocalConfig[LocalCfg_SelIndex];
                    RefreshTargetCfgList();
                    SaveCfgToLocalOrDevice(true);
                    ListBox_TargetCfgs.SelectedIndex = index;
                    return;
                }
            }
            ScanConfig.TargetConfig.Add(LocalConfig[LocalCfg_SelIndex]);
            RefreshTargetCfgList();
            SaveCfgToLocalOrDevice(true);
        }
        private Int32 SaveCfgToLocalOrDevice(Boolean IsTarget)
        {
            Int32 ret = SDK.FAIL;

            if (IsTarget == true)
            {
                if ((ret = ScanConfig.SetConfigList()) == 0)
                    Message.ShowInfo("Device Config List Update Success!");
                else
                    Message.ShowError("Device Config List Update Failed!");
                if(EditConfig)
                {
                    ListBox_TargetCfgs.SelectedIndex = EditSelectIndex;
                }
                else
                {
                    ListBox_TargetCfgs.SelectedIndex = ListBox_TargetCfgs.Items.Count - 1;
                }
            }
            else
            {
                if (!Directory.Exists(ConfigDir))
                {
                    Directory.CreateDirectory(ConfigDir);
                }
                String FileName = Path.Combine(ConfigDir, "ConfigList.xml");
                XmlSerializer xml = new XmlSerializer(typeof(List<ScanConfig.SlewScanConfig>));
                TextWriter writer = new StreamWriter(FileName);
                xml.Serialize(writer, LocalConfig);
                writer.Close();
                Message.ShowInfo("Local Config List Update Success!");
                ret = SDK.PASS;
                if(EditConfig)
                {
                    ListBox_LocalCfgs.SelectedIndex = EditSelectIndex;
                }
                else
                {
                    ListBox_LocalCfgs.SelectedIndex = ListBox_LocalCfgs.Items.Count - 1;
                }
            }

            return ret;
        }
        private void Button_CopyCfgT2L_Click(object sender, EventArgs e)
        {
            int index = 0;
            if (TargetCfg_SelIndex < 0)
            {
                Message.ShowWarning("No item selected.");
                return;
            }
            if (CheckConfigName(ListBox_TargetCfgs.Items[ListBox_TargetCfgs.SelectedIndex].ToString(), true, ref index))
            {
                DialogResult Result = Message.ShowQuestion("The config has exist, do you want to overwrite?", null, MessageBoxButtons.YesNo);
                if (Result == DialogResult.No)
                {
                    return;
                }
                if (Result == DialogResult.Yes)
                {
                    LocalConfig[index] = ScanConfig.TargetConfig[TargetCfg_SelIndex];
                    RefreshLocalCfgList();
                    SaveCfgToLocalOrDevice(false);
                    ListBox_LocalCfgs.SelectedIndex = index;
                    return;
                }
            }
            LocalConfig.Add(ScanConfig.TargetConfig[TargetCfg_SelIndex]);
            RefreshLocalCfgList();
            SaveCfgToLocalOrDevice(false);
        }
        private void Button_MoveCfgL2T_Click(object sender, EventArgs e)
        {
            int index = 0;
            if (LocalCfg_SelIndex < 0)
            {
                Message.ShowWarning("No item selected.");
                return;
            }

            if (ScanConfig.TargetConfig.Count >= 20)//Confirm the current number of device configuration before saving
            {
                Message.ShowWarning("Number of scan configs in device cannot exceed 20.");
                return;
            }
            if (CheckConfigName(ListBox_LocalCfgs.Items[ListBox_LocalCfgs.SelectedIndex].ToString(), false, ref index))
            {
                DialogResult Result = Message.ShowQuestion("The config has exist, do you want to overwrite?", null, MessageBoxButtons.YesNo);
                if (Result == DialogResult.No)
                {
                    return;
                }
                if (Result == DialogResult.Yes)
                {
                    if (DevCurCfg_Index == LocalCfg_SelIndex)
                    {
                        Message.ShowWarning("Device current configuration will be moved,\n" +
                                               "please set a new one to device later,\n" +
                                               "or you can not do scan.");

                        // Clear previous scan data
                        ClearScanPlotsUI();

                        Button_Scan.Enabled = false;
                    }
                    ScanConfig.TargetConfig[index] = LocalConfig[LocalCfg_SelIndex];
                    LocalConfig.RemoveAt(LocalCfg_SelIndex);
                    RefreshLocalCfgList();
                    SaveCfgToLocalOrDevice(false);
                    RefreshTargetCfgList();
                    SaveCfgToLocalOrDevice(true);
                    ListBox_TargetCfgs.SelectedIndex = index;
                    if (DevCurCfg_Index == TargetCfg_SelIndex)
                    {
                        SetScanConfig(ScanConfig.TargetConfig[TargetCfg_SelIndex], true, TargetCfg_SelIndex);
                    }
                }
            }
            else
            {
                if (DevCurCfg_Index == LocalCfg_SelIndex)
                {
                    Message.ShowWarning("Device current configuration will be moved,\n" +
                                           "please set a new one to device later,\n" +
                                           "or you can not do scan.");

                    // Clear previous scan data
                    ClearScanPlotsUI();
                    Button_Scan.Enabled = false;
                }
                ScanConfig.TargetConfig.Add(LocalConfig[LocalCfg_SelIndex]);
                LocalConfig.RemoveAt(LocalCfg_SelIndex);
                RefreshLocalCfgList();
                SaveCfgToLocalOrDevice(false);
                RefreshTargetCfgList();
                ListBox_TargetCfgs.SelectedIndex = ScanConfig.TargetConfig.Count - 1;
                SaveCfgToLocalOrDevice(true);       
            }  
        }
        private void Button_MoveCfgT2L_Click(object sender, EventArgs e)
        {
            int index = 0;
            if (TargetCfg_SelIndex < 0)
            {
                Message.ShowWarning("No item selected.");
                return;
            }

            if (TargetCfg_SelIndex == 0)
            {
                Message.ShowWarning("The built-in default configuration cannot be moved.");
                return;
            }
            if (CheckConfigName(ListBox_TargetCfgs.Items[ListBox_TargetCfgs.SelectedIndex].ToString(), true, ref index))
            {
                DialogResult Result = Message.ShowQuestion("The config has exist, do you want to overwrite?", null, MessageBoxButtons.YesNo);
                if (Result == DialogResult.No)
                {
                    return;
                }
                if (Result == DialogResult.Yes)
                {
                    if (DevCurCfg_Index == TargetCfg_SelIndex)
                    {
                        Message.ShowWarning("Device current configuration will be moved,\n" +
                                               "please set a new one to device later,\n" +
                                               "or you can not do scan.");

                        // Clear previous scan data
                        ClearScanPlotsUI();
                        Button_Scan.Enabled = false;
                    }
                    Int32 ActiveIndex = ScanConfig.GetTargetActiveScanIndex();

                    LocalConfig[index] = ScanConfig.TargetConfig[TargetCfg_SelIndex];
                    ScanConfig.TargetConfig.RemoveAt(TargetCfg_SelIndex);
                    if (TargetCfg_SelIndex == ActiveIndex)
                        ActiveIndex = 0;
                    else if (TargetCfg_SelIndex < ActiveIndex)
                        ActiveIndex--;
                    ScanConfig.SetTargetActiveScanIndex(ActiveIndex);

                    RefreshTargetCfgList();
                    SaveCfgToLocalOrDevice(true);
                    RefreshLocalCfgList();       
                    SaveCfgToLocalOrDevice(false);
                    ListBox_LocalCfgs.SelectedIndex = index;
                    if (DevCurCfg_Index == LocalCfg_SelIndex)
                    {
                        SetScanConfig(LocalConfig[LocalCfg_SelIndex], false, LocalCfg_SelIndex);
                    }
                }
            }
            else
            {
                if (DevCurCfg_Index == TargetCfg_SelIndex)
                {
                    Message.ShowWarning("Device current configuration will be moved,\n" +
                                           "please set a new one to device later,\n" +
                                           "or you can not do scan.");

                    // Clear previous scan data
                    ClearScanPlotsUI();

                    Button_Scan.Enabled = false;
                }

                Int32 ActiveIndex = ScanConfig.GetTargetActiveScanIndex();

                LocalConfig.Add(ScanConfig.TargetConfig[TargetCfg_SelIndex]);
                ScanConfig.TargetConfig.RemoveAt(TargetCfg_SelIndex);
                if (TargetCfg_SelIndex == ActiveIndex)
                    ActiveIndex = 0;
                else if (TargetCfg_SelIndex < ActiveIndex)
                    ActiveIndex--;
                ScanConfig.SetTargetActiveScanIndex(ActiveIndex);

                RefreshTargetCfgList();
                SaveCfgToLocalOrDevice(true);
                RefreshLocalCfgList();
                ListBox_LocalCfgs.SelectedIndex = LocalConfig.Count - 1;
                SaveCfgToLocalOrDevice(false);
            }
          
        }
        private Boolean CheckConfigName(String name,Boolean checklocal,ref int index)
        {
            Boolean isExist = false;
            if(checklocal)
            {
                for(int i=0;i< ListBox_LocalCfgs.Items.Count;i++)
                {
                    if(name == ListBox_LocalCfgs.Items[i].ToString())
                    {
                        isExist = true;
                        index = i;
                        return isExist;
                    }
                }
            }
            else
            {
                for (int i = 0; i < ListBox_TargetCfgs.Items.Count; i++)
                {
                    if (name == ListBox_TargetCfgs.Items[i].ToString())
                    {
                        isExist = true;
                        index = i;
                        return isExist;
                    }
                }
            }
            return isExist;
        }
        #endregion
        public static FW_LEVEL GetFW_LEVEL()
        {
            if (Device.IsConnected())
            {
                String HWRev = String.Empty;
                UInt32 curVer = 0;
                
                curVer = (UInt32)Device.DevInfo.TivaRev[0] << 16 | (UInt32)Device.DevInfo.TivaRev[1] << 8 | Device.DevInfo.TivaRev[2];
                if (Device.IsConnected())
                    HWRev = (!String.IsNullOrEmpty(Device.DevInfo.HardwareRev)) ? Device.DevInfo.HardwareRev.Substring(0, 1) : String.Empty;

                if (HWRev != "D" && HWRev != "B" && HWRev != "F")
                {
                    /* 
                     * TI EVM Board 
                     */
                    return FW_LEVEL.LEVEL_0;
                }
                else if (curVer >= (2 << 16 | 4 << 8 | 0))  // >= 2.4.0
                {
                    /*
                     * New Applications:
                     * 1. Support H/W Ver.F to store the four lamp ADC values
                     * 2. Bluetooth LE Advertising Name Read/Write
                     * 3. Button Lock/Unlock
                     * 4. Update error status and error code
                     */
                    return FW_LEVEL.LEVEL_4;
                }
                else if (curVer >= (2 << 16 | 1 << 8 | 2))  // >= 2.1.2
                {
                    /*
                     * New Applications:
                     */
                    return FW_LEVEL.LEVEL_3;
                }
                else if (curVer >= (2 << 16 | 1 << 8 | 0))  // >= 2.1.0.X
                {
                    /*
                     * New Applications:
                     * 1. Manufacture Serial Number Read
                     * 2. Activation Key Read/Write
                     * 3. Auto PGA Gain in Lamp On/Off
                     * 4. Check BLE Board Exist
                     * 5. Restore Calibration Coefficients
                     * 6. Lamp Usage Read/Write
                     */
                    return FW_LEVEL.LEVEL_2;
                }
                else if (curVer <= (2 << 16 | 0 << 8 | 22))  // <= 2.0.22
                {
                    /*
                     * New Applications:
                     * 1. Model Name Read/Write
                     * 2. Fixed PGA Gain Control
                     * 3. Flash UUID Read
                     */
                    return FW_LEVEL.LEVEL_1;
                }
            }
            return FW_LEVEL.LEVEL_0;
        }
        private void ClearScanPlotsUI()
        {       
            Scan.Intensity.Clear();
            Scan.ReferenceIntensity.Clear();
            Scan.Reflectance.Clear();
            Scan.Absorbance.Clear();
            Label_ScanStatus.Text = String.Empty;
            Label_CurrentConfig.Text = String.Empty;
            Label_EstimatedScanTime.Text = String.Empty;
        }
        private void OpenColseScanConfigButton(String mode)
        {
            switch(mode)
            {
                case nameof(ScanConfigMode.INITIAL):
                    Button_CfgNew.Enabled = true;
                    Button_CfgEdit.Enabled = true;
                    Button_CfgDelete.Enabled = true;
                    Button_CfgSave.Enabled = false;
                    Button_CfgCancel.Enabled = false;
                    break;
                case nameof(ScanConfigMode.NEW):
                    Button_CfgEdit.Enabled = false;
                    Button_CfgDelete.Enabled = false;
                    Button_CfgSave.Enabled = true;
                    Button_CfgCancel.Enabled = true;
                    break;
                case nameof(ScanConfigMode.EDIT):
                    Button_CfgNew.Enabled = false;
                    Button_CfgDelete.Enabled = false;
                    Button_CfgSave.Enabled = true;
                    Button_CfgCancel.Enabled = true;
                    break;
                case nameof(ScanConfigMode.DELETE):
                    Button_CfgNew.Enabled = false;
                    Button_CfgEdit.Enabled = false;
                    Button_CfgSave.Enabled = false;
                    Button_CfgCancel.Enabled = false;
                    break;
                case nameof(ScanConfigMode.SAVE):
                    Button_CfgNew.Enabled = false;
                    Button_CfgEdit.Enabled = false;
                    Button_CfgDelete.Enabled = false;
                    Button_CfgCancel.Enabled = false;
                    break;
                case nameof(ScanConfigMode.CANCEL):
                    Button_CfgNew.Enabled = true;
                    Button_CfgEdit.Enabled = true;
                    Button_CfgDelete.Enabled = true;
                    Button_CfgSave.Enabled = false;
                    Button_CfgCancel.Enabled = false;
                    break;
            }
        }
        private void SaveSettings()
        {
            /*
             * <?xml version="1.0" encoding="utf-8"?>
             * <Settings>
             *   <ScanDir>     Scan_Dir     </ScanDir>
             *   <DisplayDir>  Display_Dir  </DisplayDir>
             *   <FileFormats> ScanFile_Formats </FileFormats>
             * </Settings>
             */

            if (Scan_Dir == String.Empty && Display_Dir == String.Empty && ScanFile_Formats == 0)
                return;

            XmlDocument XmlDoc = new XmlDocument();
            XmlDeclaration XmlDec = XmlDoc.CreateXmlDeclaration("1.0", "utf-8", "");
            XmlDoc.PrependChild(XmlDec);

            // Create root element
            XmlElement Root = XmlDoc.CreateElement("Settings");
            XmlDoc.AppendChild(Root);

            // Create scan dir node under root element
            XmlElement ScanDir = XmlDoc.CreateElement("ScanDir");
            ScanDir.AppendChild(XmlDoc.CreateTextNode(Scan_Dir));
            Root.AppendChild(ScanDir);

            // Create display dir node under root element
            XmlElement DisplayDir = XmlDoc.CreateElement("DisplayDir");
            DisplayDir.AppendChild(XmlDoc.CreateTextNode(Display_Dir));
            Root.AppendChild(DisplayDir);

            // Create file format node under root element
            XmlElement FileFormats = XmlDoc.CreateElement("FileFormats");
            FileFormats.AppendChild(XmlDoc.CreateTextNode(ScanFile_Formats.ToString()));
            Root.AppendChild(FileFormats);

            // Save XML file
            String FilePath = Path.Combine(ConfigDir, "ScanPageSettings.xml");
            String dirpath = Path.GetDirectoryName(FilePath);
            if (!Directory.Exists(dirpath))
            {
                Directory.CreateDirectory(dirpath);
            }              
            XmlDoc.Save(FilePath);
        }

        private void Text_ContScan_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(Text_ContScan.Text, out int repeat) && repeat > 1)
            {
                CheckBox_SaveOneCSV.Enabled = true;
            }
            else
            {
                CheckBox_SaveOneCSV.Enabled = false;
                CheckBox_SaveOneCSV.Checked = false;
            }
        }

        private void Button_ClearAllErrors_Click(object sender, EventArgs e)
        {
            if (!Device.IsConnected())
                return;

            if (Device.ResetErrorStatus() == 0)
                RefreshErrorStatus();
        }

        private void CheckBox_SaveFileFormat_Click(object sender, System.EventArgs e)
        {
            var checkBox = sender as CheckBox;

            if (checkBox.Name.ToString() == "CheckBox_SaveDAT")
            {
                if (CheckBox_SaveDAT.Checked == false)
                {
                    DialogResult Result = Message.ShowQuestion("Are you sure to cancel saving *.dat?\n" +
                                                                      "It will not be able to display in saved scan.", null, MessageBoxButtons.YesNo);                   
                    if (Result == DialogResult.No)
                        CheckBox_SaveDAT.Checked = true;
                }
            }

            ScanFile_Formats = (CheckBox_SaveCombCSV.Checked == true) ? (ScanFile_Formats | 0x01) : (ScanFile_Formats & (~0x01));
            ScanFile_Formats = (CheckBox_SaveICSV.Checked == true) ? (ScanFile_Formats | 0x02) : (ScanFile_Formats & (~0x02));
            ScanFile_Formats = (CheckBox_SaveACSV.Checked == true) ? (ScanFile_Formats | 0x04) : (ScanFile_Formats & (~0x04));
            ScanFile_Formats = (CheckBox_SaveRCSV.Checked == true) ? (ScanFile_Formats | 0x08) : (ScanFile_Formats & (~0x08));
            ScanFile_Formats = (CheckBox_SaveIJDX.Checked == true) ? (ScanFile_Formats | 0x10) : (ScanFile_Formats & (~0x10));
            ScanFile_Formats = (CheckBox_SaveAJDX.Checked == true) ? (ScanFile_Formats | 0x20) : (ScanFile_Formats & (~0x20));
            ScanFile_Formats = (CheckBox_SaveRJDX.Checked == true) ? (ScanFile_Formats | 0x40) : (ScanFile_Formats & (~0x40));
            ScanFile_Formats = (CheckBox_SaveDAT.Checked == true) ? (ScanFile_Formats | 0x80) : (ScanFile_Formats & (~0x80));
        }

        private void SetDefaultConfig()
        {
            ScanConfig.SlewScanConfig CurConfig = new ScanConfig.SlewScanConfig
            {
                section = new ScanConfig.SlewScanSection[5]
            };
            CurConfig.head.ScanConfig_serial_number = Device.DevInfo.SerialNumber;
            CurConfig.head.config_name = "Column 1";
            CurConfig.head.scan_type = 2;
            CurConfig.head.num_sections = 1;
            CurConfig.head.num_repeats = 6;
            CurConfig.section[0].wavelength_start_nm = 900;
            CurConfig.section[0].wavelength_end_nm = 1700;
            CurConfig.section[0].num_patterns = 228;
            CurConfig.section[0].section_scan_type = 0;
            CurConfig.section[0].width_px = 6;
            CurConfig.section[0].exposure_time = 0;

            ScanConfig.TargetConfig.Add(CurConfig);
            
            int ret;
            if ((ret = ScanConfig.SetConfigList()) == 0)
            {
                Message.ShowInfo("Add default device config success!");
                ScanConfig.SetTargetActiveScanIndex(0);
                RefreshTargetCfgList();
                SetScanConfig(ScanConfig.TargetConfig[0], true, 0);
            }                
            else
            {
                Message.ShowError("Add default device config success failed!");
                ScanConfig.TargetConfig.Clear();
            }              
        }
        private ProgressBar pbw = null;
        public static double ProgressWindow_left = 0;//record the last window position when move
        public static double ProgressWindow_top = 0;//record the last window position when move
        private void ProgressWindowStart(String title, String Content)
        {
            Dispatcher dispUI = Dispatcher.CurrentDispatcher;
            Thread t = new Thread(delegate ()
            {
                dispUI.BeginInvoke(System.Windows.Threading.DispatcherPriority.Send, new Action(() =>
                {
                    pbw = new ProgressBar(title, Content)
                    {

                    };
                    pbw.ShowDialog();
                }));
            })
            {
                Priority = ThreadPriority.Highest
            };
            t.Start();
        }

        private void ProgressWindowCompleted()
        {
            Dispatcher dispUI = Dispatcher.CurrentDispatcher;
            dispUI.BeginInvoke(System.Windows.Threading.DispatcherPriority.Send, new Action(() =>
            {
                try
                {
                    pbw.Close();
                }
                catch { }
            }));
        }

        private void dataGridView_savescan_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView_savescan.SelectedRows.Count <= 0)
                return;

            String item = dataGridView_savescan.Rows[dataGridView_savescan.SelectedRows[0].Index].Cells[0].Value.ToString();
            String FileName = Path.Combine(TextBox_DisplayDirPath.Text, item);

            // Read scan result and populate to the buffer
            if (Scan.ReadScanResultFromBinFile(FileName) == SDK.FAIL)
            {
                DBG.WriteLine("Read file failed!");
                //MainWindow.ShowError("Read file failed!\nThis file may not match the format!");
                String text = "Read file failed!\nThis file may not match the format!";
                MessageBox.Show(text, "Warning");
                return;
            }

            Scan.GetScanResult();

            // Draw the scan result
            SpectrumPlot();

            // Populate config data
            ClearSavedScanCfgItems();

            for (Int32 i = 0; i < Scan.ScanConfigData.head.num_sections; i++)
            {
                Label_SavedScanType[i].Text = Helper.ScanTypeIndexToMode(Scan.ScanConfigData.section[i].section_scan_type).Substring(0, 3);
                Label_SavedRangeStart[i].Text = Scan.ScanConfigData.section[i].wavelength_start_nm.ToString();
                Label_SavedRangeEnd[i].Text = Scan.ScanConfigData.section[i].wavelength_end_nm.ToString();
                Label_SavedWidth[i].Text = Math.Round(Helper.CfgWidthPixelToNM(Scan.ScanConfigData.section[i].width_px), 2).ToString();
                Label_SavedDigRes[i].Text = Scan.ScanConfigData.section[i].num_patterns.ToString();
                Label_SavedExposure[i].Text = Helper.CfgExpIndexToTime(Scan.ScanConfigData.section[i].exposure_time).ToString();
            }
            Label_SavedAvg.Text = Scan.ScanConfigData.head.num_repeats.ToString();
        }

        private void button_search_Click(object sender, EventArgs e)
        {
            if(!String.IsNullOrEmpty(textBox_filter.Text))
            {
                LoadSavedScanList(true);
            }
        }

        private void button_clear_Click(object sender, EventArgs e)
        {
            LoadSavedScanList(false);
            textBox_filter.Text = "";
        }

        private void checkBox_tooltip_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_tooltip.Checked)
            {
                var tooltip = new DefaultTooltip
                {
                    SelectionMode = TooltipSelectionMode.SharedXInSeries
                };
                MyChart.DataTooltip = tooltip;
                MyChart.Hoverable = true;
            }
            else
            {
                MyChart.DataTooltip = null;
                MyChart.Hoverable = false;
            }
        }

        private void checkBox_zoom_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_zoom.Checked)
            {
                MyChart.Zoom = ZoomingOptions.Xy;
            }
            else
            {
                MyChart.Zoom = ZoomingOptions.None;
                MyChart.AxisX.Clear();
                MyChart.AxisY.Clear();
                if (Scan.ScanConfigData.section != null)
                {
                    int min = MAX_WAVELENGTH, max = MIN_WAVELENGTH;
                    GetMaxMinWav(ref min, ref max);
                    MyChart.AxisX.Add(new Axis { Title = "Wavelength (nm)", MinValue = min, MaxValue = max });
                }
                else
                {
                    MyChart.AxisX.Add(new Axis { Title = "Wavelength (nm)", MinValue = MIN_WAVELENGTH, MaxValue = MAX_WAVELENGTH });
                }
            }
        }

        private void ListBox_LocalCfgs_MouseClick(object sender, MouseEventArgs e)
        {
            if(ListBox_LocalCfgs.Items.Count == 0)
            {
                ListBox_LocalCfgs.BackColor = System.Drawing.Color.AliceBlue;
                ListBox_TargetCfgs.BackColor = System.Drawing.Color.White;
                ListBox_TargetCfgs.SelectedIndex = -1;
                Button_CfgEdit.Enabled = false;
                Button_CfgDelete.Enabled = false;
                ClearDetailValue();
            }
        }
        private void ClearDetailValue()
        {
            TextBox_CfgName.Text = "";
            TextBox_CfgAvg.Text = "";
            for(int i=0;i<5;i++)
            {
                TextBox_CfgRangeStart[i].Text = "";
                TextBox_CfgRangeEnd[i].Text = "";
                TextBox_CfgDigRes[i].Text = "";
            }
        }
    }
}
