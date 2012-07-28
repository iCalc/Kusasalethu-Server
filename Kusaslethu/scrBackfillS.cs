using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.IO;
using Analysis = clsAnalysis;
using TB = clsTable;
using DB = clsDBase;
using Base = clsMain;
using System.Threading;
using System.Data.OleDb;
using MetaReportRuntime;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;
using System.Net;
 
namespace Kusasalethu
{
    public partial class scrBackfillS : Form
    {
        #region Declarations
        int columnnr = 0;
        int intNoOfDays = 0;
        int noOFDay = 0;
        DateTime sheetfhs = new DateTime();
        DateTime sheetlhs = new DateTime();
        int importdone = 0;
        DataTable fixShifts = new DataTable();
        int intStartDay = 0;
        int intEndDay = 0;
        int intStopDay = 0;
        int workedShiftsFixedClockedShift = 0;
        int exitValue = 0;
        string searchEmplNr = "";
        string searchEmplName = "";
        string searchEmplGang = "";
        bool blTablenames = true;
        string Path = string.Empty;
        string strMO = "";
        string strMonthShifts = string.Empty;

        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        clsShared Shared = new clsShared();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        clsTableFormulas TBFormulas = new clsTableFormulas();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsAnalysis.clsAnalysis Analysis = new clsAnalysis.clsAnalysis();
        SqlConnection myConn = new SqlConnection();
        SqlConnection AConn = new SqlConnection();
        SqlConnection AAConn = new SqlConnection();
        SqlConnection BaseConn = new SqlConnection();
        System.Collections.Hashtable buttonCollection = new System.Collections.Hashtable();
        Dictionary<string, string> dictPrimaryKeyValues = new Dictionary<string, string>();
        Dictionary<string, string> dictGridValues = new Dictionary<string, string>();

        Dictionary<string, string> dict = new Dictionary<string, string>();
        Dictionary<string, string> GangTypes = new Dictionary<string, string>();

        string strEarningsCode = string.Empty;
        string strprevPeriod = string.Empty;
        string prevDatabaseName = string.Empty;
        string strWhere = string.Empty;
        string strActivity = string.Empty;
        string strMiningIndicator = string.Empty;
        string strServerPath = string.Empty;
        string strName = string.Empty;
        string strWhereSection = string.Empty;
        string strWherePeriod = string.Empty;
        string strMetaReportCode = "BSFnupmWkNxm8ZAA1ZhlOgL8fNdMdg4zhJj/j6T0vEyG9aSzk/HPwYcrjmawRGou66hBtseT7qJE+9hbEq9jces6bcGJmtz4Ih8Fic4UIw0Kt2lEffc05nFdiD2aQC0m";

        string dbPath = string.Empty;

        string[] ClockedShifts = new string[5];
        string[] OffShifts = new string[5];
        int intFiller = 0;
        int intCounter = 0;

        List<string> lstNames = new List<string>();
        List<string> lstTableColumns = new List<string>();
        List<string> lstPrimaryKeyColumns = new List<string>();
        List<string> lstColumnNames = new List<string>();

        Int64 intProcessCounter = 0;
        StringBuilder strSqlAlter = new StringBuilder();

        DataTable Survey = new DataTable();
        DataTable Labour = new DataTable();
        DataTable Workers = new DataTable();
        DataTable Miners = new DataTable();
        DataTable SupportLink = new DataTable();
        DataTable Designations = new DataTable();
        DataTable Offdays = new DataTable();
        DataTable Clocked = new DataTable();
        DataTable Rates = new DataTable();
        DataTable EmplPen = new DataTable();
        DataTable Configs = new DataTable();
        DataTable Factors = new DataTable();
        DataTable EmployeeTotalShifts = new DataTable();
        DataTable MineParameters = new DataTable();
        DataTable Drillers = new DataTable();
        DataTable Crews = new DataTable();
        DataTable Calendar = new DataTable();
        DataTable Production = new DataTable();
        DataTable Participants = new DataTable();
        DataTable PayrollSends = new DataTable();
        DataTable earningsCodes = new DataTable();
        DataTable Status = new DataTable();
        DataTable BonusShifts = new DataTable();
        DataTable newDataTable = new DataTable();
        DataTable _formulas = new DataTable();

        string[] arrArgs = new string[1] { "" };

        SqlDataAdapter minersTA = new SqlDataAdapter();
        BindingSource bSource = new BindingSource();
        SqlCommandBuilder _cmdBuilder = new SqlCommandBuilder();

        private ExcelDataReader.ExcelDataReader spreadsheet = null;

        ToolTip tooltip = new ToolTip();
        #endregion

        public scrBackfillS()
        {
            InitializeComponent();
            //string[] args = Program.Args;
            //arrArgs = args;
            //newdatabase(arrArgs);

        }

        internal void scrBackfillSLoad(string Period, string Region, string BussUnit, string Userid, string MiningType, string BonusType, string Environment)
        {
            #region disable all functions
            //Disable all menu functions.
            foreach (ToolStripMenuItem IT in menuStrip1.Items)
            {
                if (IT.DropDownItems.Count > 0)
                {
                    foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                    {
                        if (ITT.DropDownItems.Count > 0)
                        {
                            foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                            {
                                ITTT.Enabled = false;
                            }
                        }
                        else
                        {
                            ITT.Enabled = false;
                        }
                    }
                }
                else
                {
                    IT.Enabled = false;
                }
            }
            #endregion

            #region declarations
            BusinessLanguage.Period = Period;
            BusinessLanguage.Region = Region;
            BusinessLanguage.BussUnit = BussUnit;
            BusinessLanguage.Userid = Userid;
            BusinessLanguage.MiningType = MiningType;
            BusinessLanguage.BonusType = BonusType;
            txtMiningType.Text = MiningType;
            txtBonusType.Text = BonusType;
            strServerPath = Environment;
            BusinessLanguage.Env = Environment;
            lblEnvironment.Text = Environment;
            txtDatabaseName.Text = "BACSER8000";
            //Display dbname in text box.
            txtDatabaseName.Text = txtDatabaseName.Text.Trim();
            Base.DBName = txtDatabaseName.Text.Trim();
            Base.Period = BusinessLanguage.Period;

            //Setup the environment BEFORE the databases are moved to the classes.  This is because the environment path forms
            //part of the fisical name of the db

            setEnvironment();

            Base.DBName = txtDatabaseName.Text.Trim();
            TB.DBName = txtDatabaseName.Text.Trim();

            #endregion

            #region Connections
            //Open Connections and create classes

            AAConn = Analysis.AnalysisConnection;
            AAConn.Open();
            BaseConn = Base.BaseConnection;
            BaseConn.Open();

            #endregion

            DataTable useraccess = Base.SelectAccessByUserid(BusinessLanguage.Userid, Base.BaseConnectionString);

            #region Assign useraccess

            //BusinessLanguage.BussUnit = useraccess.Rows[0]["BUSSUNIT"].ToString().Trim();
            BusinessLanguage.Resp = useraccess.Rows[0]["RESP"].ToString().Trim();

            foreach (DataRow dr in useraccess.Rows)
            {
                string strCodeName = dr[6].ToString().Trim();
                foreach (ToolStripMenuItem IT in menuStrip1.Items)
                {
                    if (IT.DropDownItems.Count > 0)
                    {
                        foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                        {
                            if (ITT.DropDownItems.Count > 0)
                            {
                                foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                                {
                                    if (ITTT.Name.Trim() == strCodeName)
                                    {
                                        ITTT.Enabled = true;
                                    }
                                }
                            }
                            else
                                if (ITT.Name.Trim() == strCodeName)
                                {
                                    ITT.Enabled = true;
                                }
                        }
                    }
                    else
                    {
                        if (IT.Name.Trim() == strCodeName)
                        {
                            IT.Enabled = true;
                        }

                    }
                }

            }
            #endregion

            #region General
            //Display user details
            txtUserDetails.Text = BusinessLanguage.Userid + " - " + BusinessLanguage.Region + " - " + BusinessLanguage.BussUnit;
            //txtDatabaseName.Text = BusinessLanguage.BussUnit;

            txtPeriod.Text = BusinessLanguage.Period;

            // Set up the delays for the ToolTip.
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            //Force the ToolTip text to be displayed whether or not the form is active.
            tooltip.ShowAlways = true;

            //Set up the ToolTip text for the Button and Checkbox.
            tooltip.SetToolTip(this.btnImportADTeam, "Clocked Shifts");
            tooltip.SetToolTip(this.tabLabour, "Bonus Shifts");
            tooltip.SetToolTip(this.btnSearch, "Search");

            listBox2.Enabled = false;
            listBox3.Enabled = false;


            #endregion

            #region Status button collection

            //Add the buttons needed for this bonus scheme and that are on the STATUS tab.
            buttonCollection["tabCalendar"] = btnLockCalendar; 
            buttonCollection["tabLabour"] = btnLockBonusShifts;
            buttonCollection["tabParticipants"] = btnLockGangLink;
            buttonCollection["tabEmplPen"] = btnLockEmplPen;
            buttonCollection["ParticipantsEarn10"] = btnBaseCalcs;
            buttonCollection["ParticipantsEarn50"] = btnGangCalcs;
            buttonCollection["ParticipantsEarn60"] = btnSupportLinkCalc;
            buttonCollection["ParticipantsEarn90"] = btnBonusShiftsCalcs;
            buttonCollection["Input Process"] = btnInputProcess;
            buttonCollection["Paysend"] = btnLockPaysend; 
            #endregion

            #region BaseData Extracts

            //Extract Base data
            extractDesignations();
            extractConfiguration();

            #endregion

            //Extract Tab Info
            loadInfo();


        }

        private void extractPrimaryKeys(clsMain.clsMain main)
        {
            //A threat is started to extract the primary keys of selected tables.
            //The primary keys are stored in clsMain.
            //When the user select one of the selected tables tab on the front-end,
            //the list are passed from clsMain into the primary keys list.
            //No extracts are done from the databases and that makes the audit table fast.
            //ExtractKeys(Base);
            Thread t = new Thread(ExtractKeys);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractKeys(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.extractPrimaryKey();

        }

        private void extractEmployeeTotalShifts(clsMain.clsMain main)
        {
            //A threat is started to extract the employees total shifts from bonusshifts
            //The datatable will be on  clsMain. 
            //The PARTICIPANTS will not finalize if there are any employees with more shifts booked
            //than monthshifts.
            
            Thread t = new Thread(ExtractTotalShifts);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractTotalShifts(object main)
        {
            
            clsMain.clsMain M = (clsMain.clsMain)main;
            M.ExtractTotalShifts();

        }

        private void setEnvironment()
        {

            Base.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Base.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Base.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Base.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Base.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.BaseConnectionString = Base.ServerName;
            Base.Directory = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerPath"])).Trim();

            Analysis.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Analysis.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Analysis.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Analysis.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Analysis.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Analysis.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.ADTeamConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.ClockConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.DBConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.StopeConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.BackupPath = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "BackupPath"])).Trim();

            #region oleDBConnectionStringBuildeR

            if (strServerPath.ToString().Contains("Development") || strServerPath.ToString().Contains("Support"))
            {
                strServerPath = "Development";

                Base.DBConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.StopeConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.BaseConnectionString = Environment.MachineName.Trim() + Base.ServerName;
                Base.ADTeamConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
            }

            OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
            builder.ConnectionString = @"Data Source=" + Base.ServerName;
            builder.Add("Provider", "SQLOLEDB.1");
            builder.Add("Initial Catalog", Base.DBName);
            //builder.Add("Persist Security Info", "False");
            builder.Add("User ID", Base.Userid);
            builder.Add("Password", Base.PWord);

            string strdb = Base.DBName;
            //string strPath = Base.Directory.Replace("data\\", "reports\\") + strdb.Replace(BusinessLanguage.Period, "").Replace("1000", "Conn") + ".udl";
            //string strPath = "z:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            string strPath = "c:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            //MessageBox.Show("MEtatreport path en connfile :" + strPath.Trim());

            FileInfo fil = new FileInfo(strPath);

            try
            {
                File.Delete(strPath);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("delete of udl failed: " + ex.Message);
            }

            switch (strServerPath)
            {
                case "Test":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;


                case "Development":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Integrated Security", "SSPI");
                    builder.Add("Trusted_Connection", "True");
                    break;

                case "Production":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;

            }

            //MessageBox.Show("Path: " + strPath);
            bool _check = Shared.CreateUDLFile(strPath, builder);

            if (_check)
            { }
            else
            {
                MessageBox.Show("Error in creation of UDL file", "ERROR", MessageBoxButtons.OK);
            }
            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            #endregion

            myConn.ConnectionString = Base.DBConnectionString;

        }

        private void DoDataExtract(string Where)
        {
            connectToDB();
            if (Where.Trim().Length == 0)
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);
            }
            else
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName, Where);

            }
        }

        static void CreateUDLFile(string FileName, OleDbConnectionStringBuilder builder)
        {
            try
            {

                string conn = Convert.ToString(builder);
                MSDASC.DataLinksClass aaa = new MSDASC.DataLinksClass();
                aaa.WriteStringToStorage(FileName, conn, 1);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in creation of UDL file - " + ex.Message, "ERROR", MessageBoxButtons.OK);
            }
        }

        private void extractEarningsCodes()
        {

            earningsCodes = Base.SelectEarningsCodes(Base.DBConnectionString);

        }

        public void extractDBTableNames(ListBox lstbox)
        {

            lstbox.Items.Clear();
            switch (Base.DBTables.Count)
            {
                case 0:
                    lstbox.Items.Add("No tables in database");
                    break;
                default:
                    foreach (string s in Base.DBTables)
                    {
                        lstbox.Items.Add(s);
                    }
                    break;
            }
        }

        private void extractConfiguration()
        {

            Configs = Base.SelectConfigs(Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType,BusinessLanguage.BussUnit);

            grdConfigs.DataSource = Configs;

            foreach (DataRow dr in Configs.Rows)
            {
                //This extract the value identifying the first 3 digits that the gang must conform to.
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING"
                    && dr["PARM1"].ToString().Trim() == "MININGTYPE")
                {
                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMiningIndicator = strMiningIndicator + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strMiningIndicator = "(" + strMiningIndicator.Trim().Substring(1) + ")";

                }


                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING"
                    && dr["PARM1"].ToString().Trim() == "ACTIVITY")
                {
                    strActivity = string.Empty;

                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strActivity = strActivity + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strActivity = "(" + strActivity.Trim().Substring(1) + ")";
                }

                

                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGTYPE"
                        && dr["PARM1"].ToString().Trim() == "INDICATOR")
                {

                    GangTypes.Add(dr["PARM2"].ToString().Trim(), dr["PARM3"].ToString().Trim());

                }

            }


        }

        private void loadMO()
        {
            strMO = "";
            foreach (DataRow dr in Configs.Rows)
            {
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING" &&
                    dr["PARM1"].ToString().Trim() == "MO"
                    && dr["PARM2"].ToString().Trim() == txtSelectedSection.Text)
                {
                    for (int i = 6; i <= Configs.Columns.Count - 1; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMO = strMO + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }


                }
            }

            strMO = "(" + strMO.Trim().Substring(1) + ")";

        }

        private void extractDesignations()
        {
            
        }

        private void extractEarningsCode()
        {
            //Extract the records by miningtype, bonustype and paymethod
            DataTable t = Base.GetDataByMintypeBontypePaytype(txtMiningType.Text, txtBonusType.Text, "3", Base.BaseConnectionString);
            strEarningsCode = t.Rows[0]["EARNINGSCODE"].ToString().Trim();
        }

        private void evaluateOffDays()
        {
            // Display die Offday Info
            Offdays.Rows.Clear();

            loadOffdays();
        }

        private void loadOffdays()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "OFFDAYS");

            if (intCount > 0)
            {
                //YES
                string strSQL = "Select * from OFFDAYS " + strWhere;
                Offdays = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            }
            else
            {
                TB.createOffday(Base.DBConnectionString);
                TB.TBName = "OFFDAYS";
                Offdays = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from OFFDAYS");

            }

            grdOffDays.DataSource = Offdays;

            grdOffDays.Refresh();
        }

        private void evaluateFactors()
        {
            // Display die Rates info
            Factors.Rows.Clear();

            loadFactors();

        }

        private void loadFactors()
        {
            //Check if Factors exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Factors");

            if (intCount > 0)
            {
                //YES

                Factors = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Factors", " where period = '" + BusinessLanguage.Period + "'");


            }
            else
            {
                //NO - Factors DOES NOT EXIST 
            }

            grdFactors.DataSource = Factors;

            grdFactors.Refresh();
            cboVarName.Items.Clear();

            foreach (DataRow row in Factors.Rows)
            {
                cboVarName.Items.Add(row["VARNAME"].ToString().Trim());
            }
        }

        private void loadInfo()
        {
            strWherePeriod = "  where period = '" + BusinessLanguage.Period + "'";
            //Check if records in calendar exists with the selected period
            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "CALENDAR", strWherePeriod);

            if (Calendar.Rows.Count > 0)
            {

                //Run the extraction of the primary keys on its own threat.
                Shared.extractPrimaryKeys(Base);

                //Run the extraction of the views.
                Shared.createViews(Base);

                //Check if formulas exist.  If not, copy
                Shared.copyFormulas(Base);

                if (myConn.State == ConnectionState.Open)
                {
                    evaluateAll();
                }
                else
                {
                    connectToDB();
                    evaluateCalendar();
                    //Create the tab names
                    foreach (TabPage tp in tabInfo.TabPages)
                    {
                        tp.Text = tp.Tag.ToString();
                    }

                    listBox2.SelectedIndex = 0;
                }

            }
            else
            {
                //NO....
                //1. Get Previous months info  ==> MAG NIE MEER HIERIN GAAN NIE!!!!!!!!!!!!!!!!!!!!!!!

                getHistory();

                //2. Check if PREVIOUS months DB exists
                //if (BusinessLanguage.checkIfFileExists(Base.Directory + "\\" + prevDatabaseName + Base.DBExtention))
                //{
                //3. If exist - Create this selected DB and copy Formulas, Rates and Factors to the new database.
                DialogResult result = MessageBox.Show("Do you want to start a new Bonus Period: " + BusinessLanguage.Period + "?",
                                       "Information", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        this.Cursor = Cursors.WaitCursor;
                        backupAndRestoreDB();
                        copyFormulas();
                        extractDBTableNames(listBox1);
                     

                        //Run the extraction of the primary keys on its own threat.
                        Shared.extractPrimaryKeys(Base);
                        evaluateCalendar();
                        //Create the tab names
                        foreach (TabPage tp in tabInfo.TabPages)
                        {
                            tp.Text = tp.Tag.ToString();
                        }

                        listBox2.SelectedIndex = 0;

                        this.Cursor = Cursors.Arrow;
                        break;

                    case DialogResult.No:
                        btnSelect_Click("METHOD", null);
                        break;
                }
            }
        }

        private void evaluateAll()
        {

            evaluateCalendar();
            //evaluateOffDays();
            //evaluateParticipants();
            //evaluateClockedShifts();
            //evaluateLabour();

            //evaluateEmployeePenalties();
            //evaluateFactors();
            //evaluateRates();
            //evaluateCrews(); 
            extractDBTableNames(listBox1);
        }


        private void copyFormulas()
        {
            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable dtBaseFormulas = Analysis.SelectAllFormulasPerDatabaseName(Base.DBCopyName, Base.AnalysisConnectionString);
            if (dtBaseFormulas.Rows.Count > 0)
            {
                foreach (DataRow row in dtBaseFormulas.Rows)
                {
                    //Check if the receiving table already contains this formula.
                    object intCount = Analysis.countcalcbyname(Base.DBName + BusinessLanguage.Period.Trim(), row["TABLENAME"].ToString(),
                                      row["CALC_NAME"].ToString(), Base.AnalysisConnectionString);

                    if ((int)intCount > 0)
                    {
                        //rename the formula name to be inserted to NEW

                    }
                    else
                    {
                        //insert the formula.
                        Base.CopyFormulas(Base.DBName + strprevPeriod.Trim(),
                                          Base.DBName + BusinessLanguage.Period.Trim(),
                                          Analysis.AnalysisConnectionString);
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("No formulas exist on " + "\n" + "database: " + Base.DBCopyName + "\n" + "tablename: " + TB.TBCopyName +
                                "\n" + "therefor" + "\n" + "nothing will be copied", "Information", MessageBoxButtons.OK);
            }
        }


        private void evaluateMineParameters()
        {
            // Display die HOD info
            MineParameters.Rows.Clear();

            loadMineParameters();

            hideColumnsOfGrid("grdMineParameters");
        }

        private void loadMineParameters()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "MineParameters");

            if (intCount > 0)
            {
                //YES

                MineParameters = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MineParameters ", strWhere);
            }

            grdMineParameters.DataSource = MineParameters;
            grdMineParameters.Refresh();

            hideColumnsOfGrid("grdMineParameters");
        }

        private void evaluateSupportLink()
        {
            // Display die Ganglink info
            SupportLink.Rows.Clear();
            loadSupportLink();
        }

        private void loadSupportLink()
        {
            //Check if Payroll exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "SUPPORTLINK");

            if (intCount > 0)
            {
                //YES
                SupportLink = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "SupportLink", strWhere);
            }

            //grdSupportLink.DataSource = SupportLink;

            hideColumnsOfGrid("grdSupportLink");
        }

        private void confirmCopyandCreate()
        {
            listBox2.Items.Add("No sections found");

            this.Cursor = Cursors.WaitCursor;

            #region Create the new DB
            //Create the new database
            Base.createDatabase(Base.DBName, Base.ServerName);

            myConn = Base.DBConnection;
            myConn.Open();

            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createCalendarTable(Base.DBConnectionString);
            TB.createOffday(Base.DBConnectionString);
            TB.createEmployeePenalties(Base.DBConnectionString);

            //Extract Calendar again and insert into 
            DataTable calendar = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Calendar");
            grdCalendar.DataSource = calendar;

            listBox2.Items.Clear();
            listBox2.Items.Add("No sections exist yet");

            panel2.Enabled = false;
            panel3.Enabled = false;
            panel4.Enabled = false;

            this.Cursor = Cursors.Arrow;

            #endregion

        }

        private void getHistory()
        {
            #region Generate previous months db name
            //Calculate the previous months db name
            string Year = txtPeriod.Text.Trim().Substring(0, 4);
            strprevPeriod = txtPeriod.Text.Trim();

            if (txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2) == "01")
            {
                strprevPeriod = Convert.ToString(Convert.ToInt16(Year) - 1) + "12";
                prevDatabaseName = Base.DBName + strprevPeriod.Trim();
            }
            else
            {
                string strMonth = Convert.ToString(Convert.ToInt16(txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2)) - 1);
                if (strMonth.Length == 1)
                {
                    strMonth = "0" + strMonth;
                }

                strprevPeriod = Year + strMonth;
                prevDatabaseName = Base.DBName + strprevPeriod.Trim();
            }

            Base.DBCopyName = prevDatabaseName;

            #endregion

        }

        private void createAndCopyCalendar()
        {

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");

            foreach (DataRow rr in Calendar.Rows)
            {
                rr["FSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(1)).ToString("yyyy-MM-dd");
                rr["LSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(31)).ToString("yyyy-MM-dd");
            }

            TB.saveCalculations2(Calendar, Base.DBConnectionString, "", "CALENDAR");
            this.Cursor = Cursors.Arrow;
        }

        private void createAndCopyStatus()
        {
            getHistory();

            TB.createStatusTable(Base.DBConnectionString);
            myConn.Close();

            //create the Status datatable from the previous periods'table.
            Base.DBName = Base.DBCopyName;
            connectToDB();

            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");

            #region signoff from previous months DB and signon to this new DB

            myConn.Close();

            Base.DBName = TB.DBName;

            //Connect to the database that you want to copy from and load the tables into the listbox2.  Afterwards, change the db.dbname to the main database name.
            connectToDB();

            #endregion

            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            foreach (DataRow rr in Status.Rows)
            {
                strSQL.Append("insert into Status values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                              "','" + rr["STATUS"].ToString().Trim() + "','" + rr["LOCKED"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            Application.DoEvents();
            TB.InsertData(Base.DBConnectionString, "Update Status set status = 'N', locked = '0'");
            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");
            Application.DoEvents();
            this.Cursor = Cursors.Arrow;
        }

        private void backupAndRestoreDB()
        {
            //copy the data of the previous period to the current period.
            this.Cursor = Cursors.WaitCursor;
            Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period, strprevPeriod);
            this.Cursor = Cursors.Arrow;

        }

        private void evaluateInputProcessStatus()
        {

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere + " and category = 'Input Process'");

            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

            }
            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = 'Input Process'" +
                                      " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                btnLock.Text = "Lock";

            }

            evaluateStatus();
        }

        private void evaluateStatus()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "STATUS");

            if (intCount > 0)
            {
                //Status exists,  
                loadStatus();
            }
            else
            {
                createAndCopyStatus();
            }
        }

        private void statusChangeButtonColors()
        {
            foreach (DataRow rr in Status.Rows)
            {
                if (rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "Exit")
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        btnRefresh.Visible = false;
                        btnx.Visible = false;

                        pictBox.Visible = false;
                        pictBox2.Visible = false;
                        calcTime.Enabled = false;
                    }
                }
                else
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        string strButtonName = rr["PROCESS"].ToString().Trim();
                        Control c = (Control)buttonCollection[strButtonName];
                        c.BackColor = Color.LightGreen;

                    }
                    else
                    {
                        if (rr["STATUS"].ToString().Trim() == "P")
                        {
                            string strButtonName = rr["PROCESS"].ToString().Trim();
                            Control c = (Control)buttonCollection[strButtonName];
                            c.BackColor = Color.Orange;
                        }
                        else
                        {
                            if (rr["STATUS"].ToString().Trim() == "N" &&
                                pictBox.Visible == true &&
                                rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "CALC")
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.Orange;
                            }
                            else
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.PowderBlue;
                            }
                        }
                    }
                }

                Application.DoEvents();
            }
        }

        private void evaluateClockedShifts()
        {

            Clocked = Base.CShifts;
            grdClocked.DataSource = Clocked;

        }

        private void evaluateLabour()
        {

                Labour = Base.Labour;

                //Load distinct employees into lstParticipants

                DataTable Names = TB.loadDistinctValuesFrom2Columns(Labour, "EMPLOYEE_NO", "EMPLOYEE_NAME");

               
                grdLabour.DataSource = Labour;

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_No");
                //cboMinersEmpName.Items.Clear();
                cboEmplPenEmployeeNo.Items.Clear();

                foreach (string s in lstNames)
                {

                    //cboMinersEmpName.Items.Add(s.Trim());
                    cboEmplPenEmployeeNo.Items.Add(s.Trim());

                }

                lstNames = TB.loadDistinctValuesFromColumn(Labour, "GANG");
                cboCrewLinkingGang.Items.Clear();

                foreach (string s in lstNames)
                {
                    cboCrewLinkingGang.Items.Add(s.Trim());
                }              
      

            hideColumnsOfGrid("grdLabour");
        }

        private void hideColumnsOfGrid(string gridname)
        {

            switch (gridname)
            {
                

                case "grdLabour":

                    if (grdLabour.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdLabour.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("MININGTYPE"))
                    {
                        this.grdLabour.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdLabour.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdRates":
                    if (grdRates.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdRates.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("MININGTYPE"))
                    {
                        this.grdRates.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdRates.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdCalendar":
                    if (grdCalendar.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdCalendar.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("MININGTYPE"))
                    {
                        this.grdCalendar.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdCalendar.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdActiveSheet":
                    if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                    {
                        this.grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                    }

                    break;

                

                case "grdDrillers":
                    if (grdParticipants.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdParticipants.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("MININGTYPE"))
                    {
                        this.grdParticipants.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdParticipants.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdCrews":
                    if (grdCrews.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdCrews.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdCrews.Columns.Contains("MININGTYPE"))
                    {
                        this.grdCrews.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdCrews.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdCrews.Columns["BONUSTYPE"].Visible = false;
                    }
                    if (grdCrews.Columns.Contains("SECTION"))
                    {
                        this.grdCrews.Columns["SECTION"].Visible = false;
                    }
                    break;

                case "grdParticipants":
                    if (grdParticipants.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdParticipants.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("MININGTYPE"))
                    {
                        this.grdParticipants.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdParticipants.Columns["BONUSTYPE"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("SECTION"))
                    {
                        this.grdParticipants.Columns["SECTION"].Visible = false;
                    }
                    if (grdParticipants.Columns.Contains("PERIOD"))
                    {
                        this.grdParticipants.Columns["PERIOD"].Visible = false;
                    }
                    break;
            }
        }

        private void evaluateParticipants()
        {
            Participants.Rows.Clear();

            loadParticipants();

        }

        private void loadParticipants()
        {
            //Check if ganglinking exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PARTICIPANTS");

            if (intCount > 0)
            {
                //YES
                Participants = TB.createDataTableWithAdapter(Base.DBConnectionString,
                           "SELECT * FROM PARTICIPANTS WHERE SECTION = '" + txtSelectedSection.Text.Trim() + 
                           "' and period = '" + BusinessLanguage.Period + "'");
                txtAutoDGang.Clear();
            }
            else
            {

            }

            grdParticipants.DataSource = Participants;

            hideColumnsOfGrid("grdParticipants");

        }

        private void evaluateCalendar()
        {
            panel3.Enabled = true;
            panel4.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true;

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "CALENDAR");

            if (intCount > 0)
            {
                //Calendar exists,
                loadCalendar();
                loadDatePickers(0);
                loadSectionsFromCalendar();
            }
            else
            {
                createAndCopyCalendar();
            }
        }

        private void loadCalendar()
        {
            // Display die calendar info

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", " where period = '" + BusinessLanguage.Period + "'");

            grdCalendar.DataSource = Calendar;


        }

        private void loadStatus()
        {
            // Display die STATUS info
            //XXXXXXXXXXXXXXXXX
            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);  //XXXXXXXXXXXXXXXXX
            if (Status.Rows.Count > 0)
            {
                statusChangeButtonColors();
            }
            else
            {
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status");
                string tempSection = Status.Rows[0]["SECTION"].ToString().Trim();

                DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "STATUS",
                                 "Where section = '" + tempSection + "' and period = '" + BusinessLanguage.Period + "'");
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("BEGIN transaction; ");

                foreach (DataRow rr in temp.Rows)
                {
                    strSQL.Append("insert into Status values('" + rr["BUSSUNIT"].ToString().Trim() + "','" + rr["MININGTYPE"].ToString().Trim() +
                                    "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + txtSelectedSection.Text +
                                  "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                                  "','N','0');");

                }

                strSQL.Append("Commit Transaction;");
                TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
                Application.DoEvents();
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
            }
        }

        private void loadDatePickers(int Position)
        {
            //xxxxxxxxxxxxxxxx
            if (Calendar.Rows.Count > 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[Position]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[Position]["LSH"].ToString().Trim());
            }
            intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);
            lstOffDayValue.Items.Clear();
            //Load the possible dates that the user can select in this measuring period for the offday calendar
            for (DateTime i = dateTimePicker1.Value; i <= dateTimePicker2.Value; i = i.AddDays(1))
            {
                lstOffDayValue.Items.Add(i.ToString("yyyy-MM-dd"));
            }
        }

        private void UpdateClockedShifts()
        {
            //This process runs only when clocked shifts are imported.

            #region Extract dates
            //Load the section's first and last shift date
            DateTime dteFSH = dateTimePicker1.Value;
            DateTime dteLSH = dateTimePicker2.Value;

            string tempdte = Clocked.Rows[1]["FSH"].ToString().Trim();
            DateTime dteDateFrom = Convert.ToDateTime(tempdte.Trim());

            tempdte = Clocked.Rows[1]["LSH"].ToString().Trim();
            DateTime dteDateEnd = Convert.ToDateTime(tempdte.Trim());

            int intstart = dteDateFrom.Subtract(dteFSH).Days + 1;
            int intend = dteLSH.Subtract(dteDateFrom).Days + 2;

            #endregion

            evaluateOffDays();

            foreach (DataRow dr in Offdays.Rows)
            {
                string offdate = dr["OFFDAYVALUE"].ToString();
                if (offdate.Trim() == "2009-01-01")
                {
                }
                else
                {

                    DateTime dteOffdate = Convert.ToDateTime(dr["OFFDAYVALUE"].ToString());
                    int intOffday = 0;

                    if (intstart <= 0)
                    {
                        intOffday = dteOffdate.Subtract(dteDateFrom).Days;
                    }
                    else
                    {
                        intOffday = dteOffdate.Subtract(dteFSH).Days;
                    }

                    Base.UpdateOffdays(Base.DBConnectionString, intOffday);

                    Application.DoEvents();
                }
            }

        }

        private void loadSectionsFromCalendar()
        {
            lstNames = TB.loadDistinctValuesFromColumn(Calendar, "SECTION");

            if (lstNames.Count > 0)
            {

                //xxxxxxxxxxxxxxxxx
                txtSelectedSection.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                cboOffDaysSection.Text = txtSelectedSection.Text.Trim();
                label15.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + 
                           "' and period = '" + BusinessLanguage.Period + "'";
                strWhereSection = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "'";
                listBox2.Items.Clear();

                if (lstNames.Count > 1)
                {
                    foreach (string s in lstNames)
                    {
                        if (s != "XXX")
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
                else
                {
                    if (lstNames.Count == 1)
                    {
                        foreach (string s in lstNames)
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
            }
        }

        private void evaluateCrews()
        {
            // Display die Ganglink info
            Crews.Rows.Clear();
            loadCrews();
        }

        private void loadCrews()
        {
            //Check if Payroll exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Crews");

            if (intCount > 0)
            {
                //YES
                Crews = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Crews", strWhere);
            }

            lstNames.Clear();
            lstNames = TB.loadDistinctValuesFromColumn(Crews, "CREW");
            cboCrew.Items.Clear();
            cboParticipantsFilterCREWS.Items.Clear();
            cboCrewsFilterCREWS.Items.Clear();


            foreach (string s in lstNames)
            {

                cboCrew.Items.Add(s.Trim());
                cboParticipantsFilterCREWS.Items.Add(s.Trim());
                cboCrewsFilterCREWS.Items.Add(s.Trim());

            }
             
            lstNames = TB.loadDistinctValuesFromColumn(Crews, "GANG");

            cboParticipantsFilterGANGS.Items.Clear();
            

            foreach (string s in lstNames)
            {
                cboParticipantsFilterGANGS.Items.Add(s.Trim());
                
            }

            grdCrews.DataSource = Crews;
            hideColumnsOfGrid("grdCrews");

            lstNames.Clear();
            lstNames = TB.loadDistinctValuesFromColumn(Crews, "CREWTYPE");
            cboCrewType.Items.Clear();
            cboParticipantsCrewType.Items.Clear(); 


            foreach (string s in lstNames)
            {

                cboCrewType.Items.Add(s.Trim());
                cboParticipantsCrewType.Items.Add(s.Trim()); 
            }
        }

        private void evaluatePayroll()
        {
            // Display die Ganglink info
            PayrollSends.Rows.Clear();

            loadPayroll();

        }

        private void loadPayroll()
        {
            //Check if Payroll exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");

            if (intCount > 0)
            {
                //YES
                PayrollSends = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Payroll", strWhere);
            }


        }

        private void evaluateRates()
        {
            // Display die Abnormal info
            Rates.Rows.Clear();

            loadRates();

        }

        private void evaluateEmployeePenalties()
        {
            // Display die EmployeePenalties info
            EmplPen.Rows.Clear();

            loadEmployeePenalties();

        }

        private void loadEmployeePenalties()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            if (intCount > 0)
            {
                //YES

                EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            }
            else
            {
                //NO
                //Check if Bonusshifts Exists

                intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

                if (intCount > 0)
                {
                    TB.createEmployeePenalties(Base.DBConnectionString);
                    TB.TBName = "EMPLOYEEPENALTIES";
                    EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES ", strWhere);

                }
                else
                {
                }

            }

            grdEmplPen.DataSource = EmplPen;

            grdEmplPen.Refresh();

            hideColumnsOfGrid("grdEmplPen");

        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            Application.Exit();

          

        }

        private void connectToDB()
        {

            if (myConn.State == ConnectionState.Closed)
            {
                try
                {
                    myConn.Open();
                }
                catch (SystemException eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }
        }

        private void importGangLink()
        {
            DataTable newGangLink = TB.extractGangLinkFromSurvey(Base.DBConnectionString, txtSelectedSection.Text.Trim(),"STOPING",BusinessLanguage.Period);
            TB.saveCalculations2(newGangLink, Base.DBConnectionString, strWhere, "GANGLINK");
        }

        private void loadRates()
        {
            //Check if ABNORMAL exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Rates");

            if (intCount > 0)
            {
                //YES

                Rates = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Rates", " where period = '" + BusinessLanguage.Period + "'");

            }
            else
            {
                //NO - Rates DOES NOT EXIST 
            }

            grdRates.DataSource = Rates;

            grdRates.Refresh();

        }

        private void importTheSheet(string importFilename)
        {
            string path = BusinessLanguage.InputDirectory + Base.DBName;

            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                string filename = BusinessLanguage.InputDirectory + Base.DBName + importFilename;
                bool fileCheck = BusinessLanguage.checkIfFileExists(filename);

                if (fileCheck)
                {
                    FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                    spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                    fs.Close();
                    //If the file was SURVEY, all sections production data will be on this datatable.
                    //Only the selected section's data must be saved.

                    saveTheSpreadSheetToTheDatabase();
                }
                else
                {
                    MessageBox.Show("File " + filename + " - does not exist", "Check", MessageBoxButtons.OK);
                }

                //Check if file exists
                //If not  = Message
                //If exists ==>  Import
            }
            catch
            {
                MessageBox.Show("File " + importFilename + " - is inuse by another package?", "Check", MessageBoxButtons.OK);
            }
        }

        private void saveTheSpreadSheetToTheDatabase()
        {
            foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
            {
                if (dt.TableName == "SURVEY" || dt.TableName == "Survey")
                {
                    for (int i = 1; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][3].ToString().Trim() == txtSelectedSection.Text.Trim())
                        {
                        }
                        else
                        {
                            dt.Rows[i].Delete();

                        }
                    }

                }

                dt.AcceptChanges();
                //checker = true;

                TB.TBName = dt.TableName.ToString().ToUpper();
                TB.recreateDataTable();

                //Extract column names
                string strColumnHeadings = TB.getFirstRowValues(dt, Base.AnalysisConnectionString);

                switch (strColumnHeadings)
                {
                    case null:
                        break;

                    case "":
                        break;

                    default:


                        if (myConn.State == ConnectionState.Closed)
                        {
                            try
                            {
                                myConn = Base.DBConnection;
                                myConn.Open();

                                //create a table
                                bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);

                                if (tableCreate)
                                {
                                    MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    MessageBox.Show("Try again after correction of spreadsheet - input data.", "Information", MessageBoxButtons.OK);
                                }

                                //checker = false;
                            }
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.GetHashCode() + " " + ex.ToString(), "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //create a table
                            bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                            if (tableCreate)
                            {
                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);
                                MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);

                            }
                            else
                            {
                                MessageBox.Show("Data was not imported.", "Information", MessageBoxButtons.OK);
                            }
                        }

                        break;
                }
            }
        }

        private void saveXXXMiners()
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");
            string coy = "";
            string designation = "";

            #region tabMiners
            foreach (DataRow rr in Miners.Rows)
            {
                if (rr["EMPLOYEE_NO"].ToString().Trim().Contains("-"))
                {

                    coy = rr["EMPLOYEE_NO"].ToString().Substring(0, rr["EMPLOYEE_NO"].ToString().IndexOf("-")).Trim();
                }
                else
                {
                    coy = rr["EMPLOYEE_NO"].ToString().Trim();
                }

                if (rr["DESIGNATION"].ToString().Contains("-"))
                {
                    designation = rr["DESIGNATION"].ToString().Substring(0, rr["DESIGNATION"].ToString().IndexOf("-")).Trim();
                }
                else
                {
                    designation = rr["DESIGNATION"].ToString().Trim();
                }

                string test = rr["EMPLOYEE_NO"].ToString().Trim();

                strSQL.Append("insert into Miners values('" + rr["SECTION"].ToString().Trim() + "','" + rr["PERIOD"].ToString().Trim() +
                              "','xxx','" + coy + "','" + designation +
                              "','" + rr["PAYSHIFTS"].ToString().Trim() + "','" + rr["AWOP_SHIFTS"].ToString().Trim() +
                              "','" + rr["SAFETYIND"].ToString().Trim() + "');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        public String[] GetExcelSheetNames(string excelFile)
        {
            //MessageBox.Show(excelFile);
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" + excelFile + ";Extended Properties=Excel 12.0;";
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {

                    //MessageBox.Show(row["TABLE_NAME"].ToString());
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception exx)
            {
                MessageBox.Show(exx.Message);
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void btnImportADTeam_Click(object sender, EventArgs e)
        {
            DataTable temp = new DataTable();

            int intCalendarProcesses = checkLockCalendarProcesses();

            if (intCalendarProcesses > 0)
            {
                MessageBox.Show("Please finalize Calendar before importing your shifts.");
            }
            else
            {
                if (Labour.Rows.Count > 0)
                {
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                                  where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                                  where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        if (intNoOfDays <= 45)
                        {
                            loadMO();
                            refreshLabour();
                                                       
                        }
                        else
                        {
                            MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                MessageBoxButtons.OK);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("No records on Status for the Section,Period and tabLabour");
                    }
                }
                else
                {
                    evaluateStatus();
                    IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                                  where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                                  where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                                  where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                                  select locks;

                    try
                    {
                        temp = query1.CopyToDataTable<DataRow>();
                        if (temp.Rows.Count > 0)
                        {
                            if (temp.Rows[0]["STATUS"].ToString().Trim() == "N")
                            {

                                //loadDatePickers(0);
                                if (intNoOfDays <= 45)
                                {
                                    loadMO();
                                    refreshLabour();
                                   
                                }
                                else
                                {
                                    MessageBox.Show("Shifts cannot be imported.  Please fix the shifts on calendar.", "Information",
                                        MessageBoxButtons.OK);
                                }

                            }
                            else
                            {
                                MessageBox.Show("BonusShifts is locked. Please unlock before refresh.  You WILL loose all previous updates.",
                                    "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Please re-select your section.", "Information", MessageBoxButtons.OK);
                    }
                }

            }
        }

        private void createParticipants()
        {
            this.Cursor = Cursors.WaitCursor;

            if (Participants.Rows.Count > 0)
            {
                DataTable temp = TB.extractParticipantsForBacSer8000(Base.DBConnectionString, txtSelectedSection.Text.Trim(), BusinessLanguage.BussUnit,
                                                                 BusinessLanguage.Period, "");

                //This button will refresh the shifts of each employees shifts_worked and awop_shifts.
                //New employees imported from the spreadsheet, will be added to participants.

                DataTable Transfers = TB.createDataTableWithAdapter(Base.DBConnectionString, 
                                      "Select Employee_no,gang,wagecode,crew,CREWTYPE from PARTICIPANTS  " + 
                                      " where period = '" + BusinessLanguage.Period + "' order by employee_no,gang,crew,crewtype");

                DataTable Awops = TB.createDataTableWithAdapter(Base.DBConnectionString, 
                    "Select Employee_no,gang,wagecode,crew,crewtype,Awop_Shifts from PARTICIPANTS where period = '" + BusinessLanguage.Period +
                    "'  order by employee_no,gang,crew,crewtype");

                #region Update Transfers
                //Update the transfers
                foreach (DataRow empl in Transfers.Rows)
                {
                    foreach (DataRow dt in temp.Rows)
                    {
                        if (empl["EMPLOYEE_NO"].ToString().Trim() == dt["EMPLOYEE_NO"].ToString().Trim() &&
                            empl["GANG"].ToString().Trim() == dt["GANG"].ToString().Trim() &&
                            empl["WAGECODE"].ToString().Trim() == dt["WAGECODE"].ToString().Trim())
                        {
                            dt["CREW"] = empl["Crew"];
                            dt["CREWTYPE"] = empl["CrewType"];
                            break;
                        }
                    }

                    temp.AcceptChanges();
                }

                #endregion Transfers

                #region Update Awops
                //Update the AWOPS Shifts with the old awop shifts from previously imported Participants on the newly imported file
                foreach (DataRow awp in Awops.Rows)
                {
                    foreach (DataRow dt in temp.Rows)
                    {
                        if (awp["EMPLOYEE_NO"].ToString().Trim() == dt["EMPLOYEE_NO"].ToString().Trim() &&
                            awp["GANG"].ToString().Trim() == dt["GANG"].ToString().Trim() &&
                            awp["WAGECODE"].ToString().Trim() == dt["WAGECODE"].ToString().Trim() &&
                            awp["CREW"].ToString().Trim() == dt["CREW"].ToString().Trim() &&
                            awp["CREWTYPE"].ToString().Trim() == dt["CREWTYPE"].ToString().Trim())
                        {
                            if (Convert.ToInt16(awp["AWOP_SHIFTS"].ToString()) < Convert.ToInt16(dt["AWOP_SHIFTS"].ToString()))
                            {
                                dt["AWOP_SHIFTS"] = awp["AWOP_SHIFTS"];
                            }

                            break;
                        }

                    }

                    temp.AcceptChanges();
                }

                #endregion

                TB.saveCalculations2(temp, Base.DBConnectionString, " where period = '" + BusinessLanguage.Period + "'", "PARTICIPANTS");
                evaluateParticipants();
                MessageBox.Show(@"Participants were updated.", @"Information", MessageBoxButtons.OK);
            }
            else
            {
            //Extract a new participants table for BONUSSHIFTS
            DataTable temp = TB.extractParticipantsForBacSer8000(Base.DBConnectionString, txtSelectedSection.Text.Trim(), BusinessLanguage.BussUnit,
                                                                 BusinessLanguage.Period,"");


            TB.saveCalculations2(temp, Base.DBConnectionString, " where period = '" + BusinessLanguage.Period + "'", "PARTICIPANTS");
            evaluateParticipants();
            MessageBox.Show(@"Participants were created.", @"Information", MessageBoxButtons.OK);
            }

            this.Cursor = Cursors.Arrow;
        }

        private int checkLockCalendarProcesses()
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("PROCESS").TrimEnd() == "tabCalendar"
                                          where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period
                                          select locks;

            try
            {
                int intcount = query1.Count<DataRow>();

                return intcount;
            }
            catch
            {
                MessageBox.Show("Error in checkLockCalendarProcess.");
                return 0;
            }


        }

        private void refreshLabour2()
        {
            #region extract the  FSH from the database
            this.Cursor = Cursors.WaitCursor;

            extractMeasuringDates();

            //This is the refresh from the ADTeam database.
            SqlConnection _ADTeamConn = new SqlConnection();

            _ADTeamConn = Base.ADTeamConnection;
            _ADTeamConn.Open();

            DataTable ADTeam = TB.createDataTableWithAdapter(Base.ADTeamConnectionString, "select TOP 1 *  from FREEGOLD_EMPLOYEEDETAIL");
            DateTime _lastRunDate = Convert.ToDateTime(ADTeam.Rows[0]["lastrundate"]);

            int intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);
            int intStart = Base.calcNoOfDays(_lastRunDate, dateTimePicker1.Value) + 1;

            if (intStart > 100)
            {
                intStart = 100;
            }

            int intEnd = intStart - intNoOfDays;

            if (intEnd <= 0)
            {
                intEnd = 1;
            }

            #region create list of gangs

            string strListOfGangs = string.Empty;
            lstNames = TB.loadDistinctValuesFromColumn(Crews, "GANG");

            if (lstNames.Count > 0)
            {
                strListOfGangs = "Where GANG in ('";
                foreach (string s in lstNames)
                {
                    strListOfGangs = strListOfGangs.Trim() + s.Trim() + "','";

                }

                strListOfGangs = strListOfGangs.ToString().Trim().Substring(0, strListOfGangs.ToString().Trim().Length -2) + ")";
            }
            else
            {
            }

            

            #endregion

            DataTable dt = TB.ExtractADTeamShifts(Base.ADTeamConnectionString, intStart, intEnd, dateTimePicker1.Value,
                                                  dateTimePicker2.Value, intStart, intEnd,
                                                  BusinessLanguage.Period, txtSelectedSection.Text.Trim(), BusinessLanguage.MiningType,
                                                  BusinessLanguage.BonusType, BusinessLanguage.BussUnit, strListOfGangs);


            foreach (DataRow row in dt.Rows)
            {

                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }


            MessageBox.Show("Save Clocked shifts");

            string tst = string.Empty;
            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                tst = tst.Trim() + "-" + dt.Columns[i].ColumnName.Trim();
            }

            MessageBox.Show(intStart.ToString().Trim() + "-" + intEnd.ToString().Trim() + "-" + intNoOfDays.ToString().Trim() + tst);

            string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

            TB.saveCalculations2(dt, Base.DBConnectionString, strDelete, "CLOCKEDSHIFTS");

            MessageBox.Show("Clocked shifts were saved!");

            //========================================================================
            //string tst = string.Empty;
            //for (int i = 0; i <= dt.Columns.Count - 1; i++)
            //{
            //    tst = tst.Trim() + "-" + dt.Columns[i].ColumnName.Trim();
            //}

            //MessageBox.Show(intStart.ToString().Trim() + "-" + intEnd.ToString().Trim() + "-" + intNoOfDays.ToString().Trim() + "-" + tst);

            #endregion

            #region Apply offdays
            if (dt.Rows.Count > 0)
            {
                Clocked = dt.Copy();
                //Update clockedshifts with offday calendar data
                UpdateClockedShifts();
                dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                Application.DoEvents();

                #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                                txtSelectedSection.Text.Trim() + "' ORDER BY EMPLOYEE_NO";

                string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

                fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);
                BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                int intStopDay = 0;

                if (intStartDay < 0)
                {
                    //The calendarFSH falls outside the startdate of the sheet.
                    intStartDay = 0;
                }
                else
                {

                }

                if (intEndDay < 0 && intEndDay < -44)
                {
                    intStopDay = 0;
                }
                else
                {
                    if (intEndDay < 0)
                    {
                        //the LSH of the measuring period falls within the spreadsheet
                        intStopDay = intNoOfDays + intEndDay;

                    }
                    else
                    {
                        //The LSH of the measuring period falls outside the spreadsheet
                        intStopDay = 44;
                    }

                    //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                    //were not imported.

                    #region count the shifts
                    //Count the shifts

                    DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                    switch (result)
                    {
                        case DialogResult.OK:
                            extractAndCalcShifts(0, intNoOfDays);
                            MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                            break;

                        case DialogResult.Cancel:
                            MessageBox.Show("No changes was made!", "Information", MessageBoxButtons.OK);
                            break;

                    }

                    #endregion

                #endregion

                    this.Cursor = Cursors.Arrow;


                }
            }
            else
            {
                MessageBox.Show("No shifts were imported. Please check the parameters for the section.", "Information", MessageBoxButtons.OK);
                this.Cursor = Cursors.Arrow;

            }
            #endregion
        }

        private void refreshLabour()
        {
            #region create list of gangs

            string strListOfGangs = string.Empty;
            lstNames = TB.loadDistinctValuesFromColumn(Crews, "GANG");

            if (lstNames.Count > 0)
            {
                strListOfGangs = "Where GANG in ('";
                foreach (string s in lstNames)
                {
                    strListOfGangs = strListOfGangs.Trim() + s.Trim() + "','";

                }

                strListOfGangs = strListOfGangs.ToString().Trim().Substring(0, strListOfGangs.ToString().Trim().Length - 2) + ")";
            }
            else
            {
            }

            #endregion

            #region extract the sheet name and FSH and LSH of the extract
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls");

 

            if (XLSX_exists.Equals(true))
            {
                //MessageBox.Show("nou in xlsx filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xlsx", "");

                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {
                //MessageBox.Show("nou in xls filepath");
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\master" + BusinessLanguage.Period.Trim() + ".xls", "");
             
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                 
                 
                excel.CloseFile();
                 
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xls";
                 
            }
            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\adteam_" + BusinessLanguage.Period.Trim() + ".xlsx";
                //MessageBox.Show("gebruik die xlsx filepath");
            }

            
            string[] sheetNames = GetExcelSheetNames(FilePath);
             
            string sheetName = sheetNames[0];

            string testString = sheetName.Substring(0, 3).ToString().Trim();


            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[1];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[2];
            }

            if (sheetName.Substring(0, 3).ToString().Trim() != "'20")
            {
                sheetName = sheetNames[3];
            }
            #endregion

            #region import spreadsheet
            //this.Cursor = Cursors.WaitCursor;
            //DataTable dt = new DataTable();
            //OleDbConnection con = new OleDbConnection();
            //OleDbDataAdapter da;

            //con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
            //        + FilePath + ";Extended Properties='Excel 8.0;'";

            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();
            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;

            con.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 12.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            //da = new OleDbDataAdapter("select * from [" + sheetName + "]  " + strListOfGangs, con);
            da = new OleDbDataAdapter("select * from [" + sheetName + "]   where MID([GANG NAME],1,4) IN " + strMO, con);

            da.Fill(dt);

             
            #endregion

            if (dt.Rows.Count > 0)
            {

                #region remove invalid records

                //extract the column names with length less than 3.  These columns must be deleted.
                string[] columnNames = new String[dt.Columns.Count];

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (dt.Columns[i].ColumnName.Length <= 2)
                    {
                        columnNames[i] = dt.Columns[i].ColumnName;
                    }
                }

                for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
                {
                    if (string.IsNullOrEmpty(columnNames[i]))
                    {

                    }
                    else
                    {
                        dt.Columns.Remove(columnNames[i].ToString().Trim());
                        dt.AcceptChanges();
                    }
                }

                dt.Columns.Remove("INDUSTRY NUMBER");
                dt.AcceptChanges();

                if (dt.Columns.Contains("BONUS 1"))
                {
                    dt.Columns.Remove("BONUS 1");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 2"))
                {
                    dt.Columns.Remove("BONUS 2");
                    dt.AcceptChanges();
                }
                if (dt.Columns.Contains("BONUS 3"))
                {
                    dt.Columns.Remove("BONUS 3");
                    dt.AcceptChanges();
                }

                #endregion

                #region set dates
                string strSheetFSH = string.Empty;
                string strSheetLSH = string.Empty;

                //Extract the dates from the spreadsheet - the name of the spreadsheet contains the start and enddate of the extract
                string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
                string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

                //Correct the dates and calculate the number of days extracted.
                if (strSheetFSHx.Substring(6, 1) == "-")
                {
                    strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
                }
                else
                {
                    strSheetFSH = strSheetFSHx;
                }

                if (strSheetLSHx.Substring(6, 1) == "-")
                {
                    strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
                }
                else
                {
                    strSheetLSH = strSheetLSHx;
                }

                DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
                DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

                //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
                int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

                if (intNoOfDays <= 44)
                {
                    for (int j = intNoOfDays + 1; j <= 44; j++)
                    {
                        dt.Columns.Add("DAY" + j);
                    }
                }
                else
                {

                }
                #endregion

                #region Change the column names
                //Change the column names to the correct column names.
                Dictionary<string, string> dictNames = new Dictionary<string, string>();
                DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                     "Select * from varnames");
                dictNames.Clear();

                dictNames = TB.loadDict(varNames, dictNames);
                int counter = 0;

                //If it is a column with a date as a name.
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Substring(0, 1) == "2")
                    {
                        if (counter == 0)
                        {
                            strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            if (column.Ordinal == dt.Columns.Count - 1)
                            {

                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;

                            }
                            else
                            {
                                column.ColumnName = "DAY" + counter;
                                counter = counter + 1;
                            }
                        }


                    }
                    else
                    {
                        if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                        {
                            column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                        }

                    }
                }

                //Add the extra columns
                dt.Columns.Add("BUSSUNIT");
                dt.Columns.Add("FSH");
                dt.Columns.Add("LSH");
                dt.Columns.Add("SECTION");
                dt.Columns.Add("EMPLOYEETYPE");
                dt.Columns.Add("PERIOD");      //xxxxxxxx
                dt.AcceptChanges();


                foreach (DataRow row in dt.Rows)
                {
                    row["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                    row["FSH"] = strSheetFSH;
                    row["LSH"] = strSheetLSH;
                    row["MININGTYPE"] = "BACKFILL";
                    row["BONUSTYPE"] = "SERVICES";
                    row["PERIOD"] = BusinessLanguage.Period;   //xxx
                    if (row["GANG"].ToString().Length > 0)
                    {
                        row["SECTION"] = txtSelectedSection.Text.Trim();
                    }
                    else
                    {
                        row["SECTION"] = "XXX";
                    }
                    if (row["WAGECODE"].ToString().Trim() == "")
                    {
                        row["WAGECODE"] = "00000";
                    }
                    else
                    {
                    }
                    row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                    for (int i = 0; i <= dt.Columns.Count - 1; i++)
                    {
                        if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                        {
                            row[i] = "-";
                        }
                    }
                }

                //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
                DataColumn dcBussunit = new DataColumn();
                dcBussunit.ColumnName = "BUSSUNIT";
                dt.Columns.Remove("BUSSUNIT");
                dt.AcceptChanges();
                InsertAfter(dt.Columns, dt.Columns["BONUSTYPE"], dcBussunit);

                foreach (DataRow dr in dt.Rows)
                {
                    dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                }

                #endregion


                //Write to the database
                
                TB.saveCalculations2(dt, Base.DBConnectionString, " WHERE PERIOD = '" + BusinessLanguage.Period.Trim() + "'", "CLOCKEDSHIFTS");
                

                if (dt.Rows.Count > 0)
                {
                    Clocked = dt.Copy();
                    //Update clockedshifts with offday calendar data
                    UpdateClockedShifts();
                    dt = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Clockedshifts");

                    Application.DoEvents();

                    grdClocked.DataSource = dt;

                    #region Calculate the shifts per employee en output to bonusshifts

                    string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                                    "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts " + strListOfGangs + 
                                    " and section = '" +  txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "' order by gang";

                    //string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

                    //fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);
                    BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
                    //exportToExcel("c:\\", BonusShifts);
                    string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                    DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                    sheetfhs = SheetFSH;
                    sheetlhs = SheetLSH;
                    int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
                    int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
                    int intStopDay = 0;

                    if (intStartDay < 0)
                    {
                        //The calendarFSH falls outside the startdate of the sheet.
                        intStartDay = 0;
                    }
                    else
                    {

                    }

                    if (intEndDay < 0 && intEndDay < -44)
                    {
                        intStopDay = 0;
                    }
                    else
                    {
                        if (intEndDay < 0)
                        {
                            //the LSH of the measuring period falls within the spreadsheet
                            intStopDay = intNoOfDays + intEndDay;

                        }
                        else
                        {
                            //The LSH of the measuring period falls outside the spreadsheet
                            intStopDay = 44;
                        }

                        //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                        //were not imported.

                        #region count the shifts
                        //Count the shifts
                        //Start the evaluate here, because has to make sure that the previous savecalc2 actually finished.
                        Shared.evaluateDataTable(Base, "CLOCKEDSHIFTS");


                        DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                        switch (result)
                        {
                            case DialogResult.OK:
                                extractAndCalcShifts(intStartDay, intStopDay);
                                Shared.evaluateDataTable(Base, "BONUSSHIFTS");

                                MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                                evaluateShifts();
                                break;

                            case DialogResult.Cancel:
                                MessageBox.Show("No changes was made!", "Information", MessageBoxButtons.OK);
                                break;

                        }

                        #endregion

                    #endregion
                        
                        this.Cursor = Cursors.Arrow;
                        File.Delete(FilePath);

                    }
                }
                else
                {
                    MessageBox.Show("No shifts were imported. Please check the parameters for the section.", "Information", MessageBoxButtons.OK);
                    this.Cursor = Cursors.Arrow;
                    File.Delete(FilePath);
                }
            }
            else
            {
                MessageBox.Show("No records were imported from spreadsheet.", "Information", MessageBoxButtons.OK);
            }
        }
        
        private void extractAndCalcShifts(int DayStart, int DayEnd)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int shiftsCheck = 0;
            BonusShifts.Columns.Add("TMLEADERIND");

            foreach (DataRow row in BonusShifts.Rows)
            {
                foreach (DataColumn column in BonusShifts.Columns)
                {
                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                        {
                            if (row[column].ToString().Trim() == "U" || 
                                row[column].ToString().Trim() == "u" ||
                                row[column].ToString().Trim() == "W" ||
                                row[column].ToString().Trim() == "w")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                                shiftsCheck = 1;
                            }
                            else
                            {
                                //if (row[column].ToString().Trim() == "A" || row[column].ToString().Trim() == "b")
                                if (row[column].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else { }

                            }
                        }
                        else
                        {
                            row[column] = "*";
                        }
                    }
                    else
                    {
                        if (column.ColumnName == "BONUSTYPE")
                        {
                            row["BONUSTYPE"] = "SERVICES";
                        }
                    }
                }//foreach datacolumn

                //If shifts_worked > monthsshifts then employee_shifts = monthshifts
                if (Convert.ToInt16(intShiftsWorked) > Convert.ToInt16(strMonthShifts))
                {
                    row["SHIFTS_WORKED"] = strMonthShifts.Trim();
                }
                else
                {
                    row["SHIFTS_WORKED"] = Convert.ToString(intShiftsWorked);
                }

                row["AWOP_SHIFTS"] = intAwopShifts;
                row["TMLEADERIND"] = "0";
                intShiftsWorked = 0;
                intAwopShifts = 0;
            }
            //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
            DataColumn dcPeriod = new DataColumn();
            dcPeriod.ColumnName = "PERIOD";
            BonusShifts.Columns.Remove("PERIOD");
            BonusShifts.AcceptChanges();
            InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

            foreach (DataRow dr in BonusShifts.Rows)
            {
                dr["PERIOD"] = BusinessLanguage.Period;
            }

            string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

            TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");

            //if (importdone == 0)
            //{

            //    fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
            //    importdone = 1;

            //}

            Application.DoEvents();
        }

        public void InsertAfter(DataColumnCollection columns, DataColumn currentColumn, DataColumn newColumn)
        {
            if (columns.Contains(currentColumn.ColumnName))
            {
                columns.Add(newColumn);
                //add the new column after the current one 
                columns[newColumn.ColumnName].SetOrdinal(currentColumn.Ordinal + 1);
            }
            else
            {
                 
            }
        }

        private void extractGangLink()
        {
            //Add the rigging, equipping and tramming gangs to the ganglinking.
            DataTable TmpGanglink = Base.extractGanglink(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType, txtSelectedSection.Text.Trim());

            if (TmpGanglink.Rows.Count > 0)
            {

                TB.saveCalculations2(TmpGanglink, Base.DBConnectionString, strWhere, "GANGLINK");
                Application.DoEvents();

            }
            else
            {
                MessageBox.Show("No records for ganglinking were extracted for section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
        }

        private void btnLock_Click(object sender, EventArgs e)
        {

            string strProcess = tabInfo.SelectedTab.Name;

            if (btnLock.Text == "Lock")
            {
                if (strProcess == "tabParticipants")
                {
                    newDataTable = TB.extractEmployeesWithShiftsMoreThanMonthShifts(Base.DBConnectionString, "PARTICIPANTS", strMonthShifts, 
                                                                                    BusinessLanguage.Period);
                    if (newDataTable.Rows.Count > 0)
                    {
                        grdParticipants.DataSource = newDataTable;
                        MessageBox.Show("The following employees have more shifts that the allowed measuring shifts." + Environment.NewLine + 
                                        "The Participants input screen will not finalize until shifts have been fixed.", "Warning", MessageBoxButtons.OK);
                    }
                    else
                    {
                        TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                     "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                        btnLock.Text = "Unlock";
                        evaluateInputProcessStatus();
                        openTab(tabProcess);

                        Application.DoEvents();
                    }
                }
                else
                {
                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                      "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                btnLock.Text = "Unlock";
                evaluateInputProcessStatus();
                openTab(tabProcess);

                Application.DoEvents();
                }

            }

            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = '" + strProcess +
                                      "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                btnLock.Text = "Lock";
                evaluateInputProcessStatus();
                openTab(tabProcess);

                Application.DoEvents();

            }


        }
        
        private void grdMiners_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            string strSQL = string.Empty;
            string strName = string.Empty;
            string strDesignation = string.Empty;
            string strDesignationDesc = string.Empty;

            switch (tabInfo.SelectedTab.Name)
            {
             
                case "tabEmplPen":
                    #region tabEmployee Penalties
                    if (cboEmplPenEmployeeNo.Text.Trim().Length > 0 &&
                        txtPenaltyValue.Text.Trim().Length > 0 && cboPenaltyInd.Text.Trim().Length > 0)
                    {
                        DataRow dr;
                        dr = EmplPen.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboEmplPenEmployeeNo.Text.Trim();
                        dr["PENALTYVALUE"] = txtPenaltyValue.Text.Trim();
                        dr["PENALTYIND"] = cboPenaltyInd.Text.Trim();

                        EmplPen.Rows.Add(dr);

                        strSQL = "Insert into EmployeePenalties values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + cboEmplPenEmployeeNo.Text.Trim() + "', '" + txtPenaltyValue.Text.Trim() +
                                 "', '" + cboPenaltyInd.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                    }
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {
                        DataRow dr;
                        dr = Rates.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["RATE_TYPE"] = txtRateType.Text.Trim();
                        dr["LOW_VALUE"] = txtLowValue.Text.Trim();
                        dr["HIGH_VALUE"] = txtHighValue.Text.Trim();
                        dr["RATE"] = txtRate.Text.Trim();

                        int rowindex = grdRates.CurrentCell.RowIndex;
                        strSQL = "Insert into Rates values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtRateType.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + txtLowValue.Text.Trim() + "', '" + txtHighValue.Text.Trim() +
                                 "', '" + txtRate.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdRates.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabOffdays":
                    #region tabOffdays
                    this.Cursor = Cursors.WaitCursor;
                    if (cboOffDaysSection.Text.Trim().Length > 0)
                    {

                        //Get the layout of the offday file.
                        DataTable temp = new DataTable();
                        temp = Offdays.Copy();

                        int intRow = 0;
                        if (grdOffDays.CurrentRow == null)
                        {
                            intRow = 0;
                        }
                        else
                        {
                            intRow = grdOffDays.CurrentCell.RowIndex;
                        }

                        //Clear the input temp table
                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();
                        }

                        temp.AcceptChanges();

                        if (lstOffDayValue.SelectedItems.Count == 0)
                        {
                            MessageBox.Show("Please select dates from the listbox", "Information", MessageBoxButtons.OK);
                        }
                        else
                        {

                            for (int i = 0; i < lstOffDayValue.SelectedItems.Count; i++)
                            {
                                DataRow dr = temp.NewRow();
                                dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();             //xxxxxxxxxxxxxxxx
                                dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();//xxxxxxxxxxxxxxxx
                                dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();//xxxxxxxxxxxxxxxx
                                dr["SECTION"] = cboOffDaysSection.Text.Trim();
                                dr["PERIOD"] = BusinessLanguage.Period;//xxxxxxxxxxxxxxxx
                                dr["GANG"] = cboOffDaysGang.Text.Trim();
                                dr["OFFDAYVALUE"] = lstOffDayValue.SelectedItems[i].ToString();

                                temp.Rows.Add(dr);

                            }
                            //Create a invalid delete that will execute in the savecalculation2 method.
                            string strDelete = " where section = '999'";

                            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "OFFDAYS");
                            evaluateOffDays();

                            grdOffDays.FirstDisplayedScrollingRowIndex = intRow;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }

                    this.Cursor = Cursors.Arrow;
                    break;

                    #endregion

                case "tabCrews":
                    #region tabCrews
                    this.Cursor = Cursors.WaitCursor;
                    if (cboCrew.Text.Trim().Length > 0 && cboCrewLinkingGang.Text.Trim().Length > 0)
                    {

                        //Get the layout of the crew file.
                        DataTable temp = new DataTable();
                        temp = Crews.Copy();

                        int intRow = 0;
                        if (grdCrews.CurrentRow == null)
                        {
                            intRow = 0;
                        }
                        else
                        {
                            intRow = grdCrews.CurrentCell.RowIndex;
                        }

                        //Clear the input temp table
                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();
                        }

                        temp.AcceptChanges();

                        
                        DataRow dr = temp.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit;
                        dr["MININGTYPE"] = BusinessLanguage.MiningType;
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType;
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = BusinessLanguage.Period.Trim();
                        dr["GANG"] = cboCrewLinkingGang.Text.Trim();
                        dr["CREW"] = cboCrew.Text.Trim();
                        dr["CREWTYPE"] = cboCrewType.Text.Trim();
                        dr["TONS"] = txtTons.Text.Trim();
                        dr["SAFETYIND"] = cboSafetyInd.Text.Trim();

                        temp.Rows.Add(dr);


                            //Create a invalid delete that will execute in the savecalculation2 method.
                            string strDelete = " where section = '999'";

                            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "CREWS");
                            evaluateCrews();

                            grdCrews.FirstDisplayedScrollingRowIndex = intRow;
                    }
                  
                    else
                    {
                        MessageBox.Show("Supply all input data. Please check that all input boxes contain data.", "Error", MessageBoxButtons.OK);
                    }

                    this.Cursor = Cursors.Arrow;
                    break;

                    #endregion
            }
        }

        private string checkSQL(int intCounter, string strSQL)
        {
            if (intCounter > 0)
            {
                for (int i = 0; i <= intCounter - 1; i++)
                {
                    strSQL = strSQL.Trim() + ",'0'";
                }
                strSQL = strSQL.Trim() + ")";
            }
            else
            {
                strSQL = strSQL.Trim() + "')";
            }

            return strSQL;
        }

        private void grdCrews_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            DataTable temp = new DataTable();

            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboCrewLinkingGang.Text = grdCrews["GANG", e.RowIndex].Value.ToString().Trim();
                cboCrew.Text = grdCrews["CREW", e.RowIndex].Value.ToString().Trim();
                cboCrewType.Text = grdCrews["CREWTYPE", e.RowIndex].Value.ToString().Trim();
                txtTons.Text = grdCrews["TONS", e.RowIndex].Value.ToString().Trim();
                 
                //Indirect teams get the tons from Mine parameters and user is not allowed to change it on the Crews tab.

                if (grdCrews["CREWTYPE", e.RowIndex].Value.ToString().Trim() == "INDIRECT")
                {
                    txtTons.Enabled = false;
                }
                else
                {
                    txtTons.Enabled = true;
                }

            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdCrews[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdCrews[s, e.RowIndex].Value.ToString().Trim());
                }
            }
            #endregion
        }

        private void tabInfo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text == "***")
            {
                MessageBox.Show("Please select a section.", "Information", MessageBoxButtons.OK);
            }
            else
            {
                btnInsertRow.Enabled = true;
                btnUpdate.Enabled = true;

                btnDeleteRow.Enabled = false;
                listBox1.Enabled = false;                               //HJ
                btnLoad.Enabled = false;
                dateTimePicker1.Enabled = false;                        //HJ
                dateTimePicker2.Enabled = false;                        //HJ
                btnPrint.Enabled = false;
                btnLock.Enabled = false;
                panelLock.BackColor = Color.Lavender;

                int intCount = checkLock(tabInfo.SelectedTab.Name);
                if (intCount > 0)
                {
                    btnLock.Text = "Unlock";
                }
                else
                {
                    btnLock.Text = "Lock";
                }

                switch (tabInfo.SelectedTab.Name)
                {
                    #region tabCalendar
                    case "tabCalendar":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnLoad.Enabled = true;
                        dateTimePicker1.Enabled = true;                 //HJ
                        dateTimePicker2.Enabled = true;                 //HJ
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        lstPrimaryKeyColumns.Clear();
                        extractPrimaryKey(Calendar, "CALENDAR");

                        break;
                    #endregion
         
                    #region tabClockShifts
                    case "tabClockShifts":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnPrint.Enabled = true;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;
                    #endregion

                    #region tabLabour
                    case "tabLabour":

                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;

                        evaluateLabour();

                        extractPrimaryKey(Labour, "BONUSSHIFTS");
                        break;
                    #endregion

                    #region tabCrews
                    case "tabCrews":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateCrews();
                        extractPrimaryKey(Crews, "CREWS");
                        break;
                    #endregion

                    #region tabSupportLinking
                    case "tabSupportLink":
                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateSupportLink();
                        extractPrimaryKey(SupportLink, "SUPPORTLINK");
                        break;

                    #endregion

                    #region tabConfig
                    case "tabConfig":

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;

                        extractPrimaryKey(Configs, "CONFIGURATION");
                        break;

                    #endregion

                    #region tabEmplPen
                    case "tabEmplPen":

                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        extractPrimaryKey(EmplPen, "EMPLOYEEPENALTY");
                        break;

                    #endregion

                    #region tabSelected
                    case "tabSelected":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        listBox1.Enabled = true;                            //HJ
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        extractDBTableNames(listBox1);
                        hideColumnsOfGrid("grdActiveSheet");
                        break;

                    #endregion

                    #region tabMineParameter

                    case "tabMineParameter":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = true;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabRates
                    case "tabRates":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        extractPrimaryKey(Rates, "RATES");
                        break;

                    #endregion

                    #region tabFactors

                    case "tabFactors":

                        evaluateFactors();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = true;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.LightGray;
                        panelUpdate.BackColor = Color.LightGray;
                        panelDelete.BackColor = Color.LightGray;
                        panelPreCalcReport.BackColor = Color.LightGray;
                        extractPrimaryKey(Factors, "FACTORS");
                        break;

                    #endregion

                    #region tabOffdays

                    case "tabOffdays":

                        btnInsertRow.Enabled = true;
                        btnUpdate.Enabled = true;
                        btnDeleteRow.Enabled = true;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.LightGray;
                        panelUpdate.BackColor = Color.LightGray;
                        panelDelete.BackColor = Color.LightGray;
                        break;

                    #endregion

                    #region tabParticipants
                    case "tabParticipants":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        btnShowAll.Visible = false;
                        btnShowEmpl.Visible = false;
                        extractPrimaryKey(Participants, "PARTICIPANTS");
                        break;

                    #endregion

                    #region tabFactors

                    case "tabMineParameters":

                        evaluateFactors();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = true;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.LightGray;
                        panelUpdate.BackColor = Color.LightGray;
                        panelDelete.BackColor = Color.LightGray;
                        panelPreCalcReport.BackColor = Color.LightGray;
                        break;

                    #endregion
                }
            }
        }

        private void extractPrimaryKey(DataTable p, string tablename)
        {
            //List Names contains the primary key columns of the selected table
            lstPrimaryKeyColumns.Clear();
            switch (tablename)
            {
                case "CALENDAR":
                    lstPrimaryKeyColumns = Base.listCalendarPrimaryKey;
                    break;

                case "BONUSSHIFTS":
                    lstPrimaryKeyColumns = Base.listBonusShiftsPrimaryKey;
                    break;
            
                case "PARTICIPANTS":
                    lstPrimaryKeyColumns = Base.listParticipantsPrimaryKey;
                    break;

                case "RATES":
                    lstPrimaryKeyColumns = Base.listRatesPrimaryKey;
                    break;

                case "FACTORS":
                    lstPrimaryKeyColumns = Base.listFactorsPrimaryKey;
                    break;

                case "CONFIGURATION":
                    lstPrimaryKeyColumns = Base.listConfigurationPrimaryKey;
                    break;


            }

            ////lstTableColumns contains all the column names of the table excluding "BUSSUNIT","MININGTYPE","BONUSTYPE","PERIOD")
            ////Do this extract on the table in memory, because much quicker.
            lstTableColumns.Clear();
            DataTable temp = p.Copy();
            deleteAllCalcColumnsFromTempTable(tablename, temp);

            if (temp.Columns.Count > 0)
            {
                foreach (DataColumn col in temp.Columns)
                {
                    if (col.ColumnName == "BUSSUNIT" || col.ColumnName == "MININGTYPE" || col.ColumnName == "BONUSTYPE" || col.ColumnName == "PERIOD")
                    {
                    }
                    else
                    {
                        lstTableColumns.Add(col.ColumnName.ToString().Trim());
                    }
                }
            }
        }

        private int checkLock(string processToBeChecked)
        {
            //Lynx....LINQ
            DataTable contactTable = TB.getDataTable(TB.TBName);

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "Y"
                                          where locks.Field<string>("PROCESS").TrimEnd() == processToBeChecked
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;


            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();
            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }

        private int checkLockInputProcesses()
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }
        
        private void grdEmplPen_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboEmplPenEmployeeNo.Text = grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtPenaltyValue.Text = grdEmplPen["PENALTYVALUE", e.RowIndex].Value.ToString().Trim();
                cboPenaltyInd.Text = grdEmplPen["PENALTYIND", e.RowIndex].Value.ToString().Trim();
                if (grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXXXXXXXXXXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }
            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdEmplPen[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdEmplPen[s, e.RowIndex].Value.ToString().Trim());
                }
            }
            #endregion

            Cursor.Current = Cursors.Arrow;

        }

        private void grdLabour_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
               


            }

            Cursor.Current = Cursors.Arrow;

        }

        private void grdConfigs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboParameterName.Text = grdConfigs["PARAMETERNAME", e.RowIndex].Value.ToString().Trim();
                cboParm1.Text = grdConfigs["PARM1", e.RowIndex].Value.ToString().Trim();
                cboParm2.Text = grdConfigs["PARM2", e.RowIndex].Value.ToString().Trim();
                cboParm3.Text = grdConfigs["PARM3", e.RowIndex].Value.ToString().Trim();
                cboParm4.Text = grdConfigs["PARM4", e.RowIndex].Value.ToString().Trim();
                cboParm5.Text = grdConfigs["PARM5", e.RowIndex].Value.ToString().Trim();
                cboParm6.Text = grdConfigs["PARM6", e.RowIndex].Value.ToString().Trim();
                cboParm7.Text = grdConfigs["PARM7", e.RowIndex].Value.ToString().Trim();
            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdConfigs[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdConfigs[s, e.RowIndex].Value.ToString().Trim());
                }
            }
            #endregion

            Cursor.Current = Cursors.Arrow;
        }

        private void grdOffdays_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        #region AutoSize

        private void autoSizeGrid(DataGridView DG)
        {
            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader.ToString())
            {
                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                {
                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                }
                else
                {
                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.ColumnHeader.ToString())
                    {
                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    }
                    else
                    {
                        if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                        {
                            DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                        }
                        else
                        {
                            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                            {
                                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                            }
                            else
                            {
                                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.Fill.ToString())
                                {
                                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                                }
                                else
                                {
                                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                                    {
                                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void grdActiveSheet_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdActiveSheet);
            }
        }

        private void grdCalendar_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdCalendar);
            }
        }

        private void grdClocked_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdClocked);
            }
        }

        private void grdLabour_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdLabour);
            }
        }

        private void grdConfigs_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdConfigs);
            }
        }

        private void DoDataExtract()
        {
            connectToDB();
            TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);

        }
        #endregion

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {//xxxxxxxxxxxxxxxxxxxx
            string FormulaTableName = string.Empty;

            TB.TBName = (string)listBox1.SelectedItem;

            if (TB.TBName.Trim().ToUpper().Contains("EARN") && TB.TBName.Trim().ToUpper().Contains("20"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));   //xxxxxxxxxxxxxxxxxx
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            TB.DBName = Base.DBName;

            connectToDB();
            cboColumnValues.Items.Clear();
            cboColumnNames.Items.Clear();
            cboColumnNames.Text = string.Empty;
            cboColumnValues.Text = string.Empty;

            List<string> lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, TB.TBName);

            foreach (string s in lstColumnNames)
            {
                cboColumnNames.Items.Add(s.Trim());
                 
            }

            TB.ListOfSelectedTableColumns = lstColumnNames;

            DoDataExtract(strWhere);
            newDataTable = TB.getDataTable(TB.TBName);
            if (newDataTable == null)
            {
                DoDataExtract(strWherePeriod);
                newDataTable = TB.getDataTable(TB.TBName);

            }
            //if newdatatable is still null
            if (newDataTable == null)
            {
                DoDataExtract("");
                newDataTable = TB.getDataTable(TB.TBName);

            }

            grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            hideColumnsOfGrid("grdActiveSheet");
        }

        private void exportToExcel(string path, DataTable dt)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = path + "\\" + TB.TBName + ".xls";
                try
                {
                    StreamWriter SW = new StreamWriter(OPath);
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {
                            grid.RenderControl(HTMLWriter);
                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();
                    MessageBox.Show("Your spreadsheet was created at: " + OPath, "Information", MessageBoxButtons.OK);
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void TBExport_Click(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void saveTheSpreadSheet()
        {
            string path = @"c:\" + TB.DBName + "\\" + TB.TBName;
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                DoDataExtract();
                DataTable outputTable = TB.getDataTable(TB.TBName);
                exportToExcel(path, outputTable);
                MessageBox.Show("Successfully Downloaded.", "Information", MessageBoxButtons.OK);

            }
            catch (Exception ee)
            {
                Console.WriteLine("The process failed: {0}", ee.ToString());
            }

            finally { }
        }

        private void grdActiveSheet_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            object intCount = Analysis.countcalcbyname(TB.DBName, TB.TBName, TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);
            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetails(TB.DBName, TB.TBName, TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS");
                DataTable dtFactors = TB.getDataTable("FACTORS");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        MessageBox.Show("SQL extract", "Not available to be tested", MessageBoxButtons.OK);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {
                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void loadDict(DataTable _datatable)
        {
            foreach (DataRow _row in _datatable.Rows)
            {
                string str = _row[0].ToString().Trim();
                if (dict.ContainsKey(str))
                {
                    dict.Remove(str);
                    dict.Add(str, _row[1].ToString().Trim());
                }
                else
                {
                    dict.Add(str, _row[1].ToString().Trim());
                }
            }
            dict.Remove("X");
            dict.Add("X", "0");

        }

        private void buildDisplaySQL(string strwhere, decimal decValue)
        {
            string strSQL = "";

            strSQL = "Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                         TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + TBFormulas.A + TBFormulas.B + TBFormulas.C + TBFormulas.D + TBFormulas.E + TBFormulas.F + TBFormulas.G + TBFormulas.H + " " + strwhere;
            strSQL = strSQL.Replace("#", "").Replace(":and:", "and").Replace(" from ", "\n from ").Replace(" and ", "\n and ").Replace(" where ", "\n where ");

            General.textTestSQL = strSQL;
            scrQuerySQL testsql = new scrQuerySQL();
            testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
            testsql.ShowDialog();

        }

        private void userProfile_Click(object sender, EventArgs e)
        {
            scrProfile userProfile = new scrProfile();
            userProfile.FormLoad(BusinessLanguage, BaseConn);
            userProfile.Show();
        }

        private void grantAccessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scrSecurity useraccess = new scrSecurity();
            useraccess.userAccessLoad(myConn, Base, TB, BusinessLanguage.Userid, strServerPath.ToString().ToUpper());
            useraccess.Show();
        }

        private void btn0_Click(object sender, EventArgs e)
        {

            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "0";

        }

        private void btn1_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "1";
        }

        private void btn2_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "2";
        }

        private void btn3_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "3";
        }

        private void btn4_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "4";
        }

        private void btn5_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "5";
        }

        private void btn6_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "6";
        }

        private void btn7_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "7";
        }

        private void btn8_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "8";
        }

        private void btn9_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "9";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = "";
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataTable searchEmpl = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from ClockedShifts where employee_no like '%" + txtSearchEmpl.Text.Trim() + "%'");

            if (searchEmpl.Rows.Count > 0)
            {
                //amp
                string strLSH = Clocked.Rows[0]["LSH"].ToString().Trim();
                DateTime LSH = Convert.ToDateTime(strLSH);
                string Mnth = string.Empty;
                string Day = string.Empty;
                foreach (DataColumn dc in searchEmpl.Columns)
                {
                    if (dc.Caption.Substring(0, 3) == "DAY")
                    {
                        double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                        string strTemp = Clocked.Rows[0]["FSH"].ToString().Trim();
                        DateTime temp = Convert.ToDateTime(strTemp);
                        temp = temp.AddDays(d);
                        if (temp > LSH)  //remember the days start at 0
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                        }
                        else
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                        }
                    }
                }
            }
            grdClocked.DataSource = searchEmpl;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            grdClocked.DataSource = Clocked;
        }

        private void grdActiveSheet_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            string FormulaTableName = string.Empty;

            if (TB.TBName.Trim().ToUpper().Contains("EARN"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            object intCount = Analysis.countcalcbyname(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                       TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);

            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetailsDCript(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName,
                                                                    TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                //TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS"," Where period = '" + BusinessLanguage.Period + "'");
                DataTable dtFactors = TB.createDataTableWithAdapter(Base.DBConnectionString,
                                    "Select Varname,Varvalue from FACTORS where period = '" + BusinessLanguage.Period + "'");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        string strWhere = " ";
                        for (int i = 0; i < grdActiveSheet.Columns.Count - 1; i++)
                        {

                            strWhere = strWhere.Trim() + " and t1." + grdActiveSheet.Columns[i].HeaderText.Trim() +
                                       " = '" + (string)(grdActiveSheet[i, e.RowIndex].Value).ToString().Trim() + "'";

                        }

                        buildDisplaySQL(strWhere, decValue);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {

                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + BusinessLanguage.Period.Trim() + '\n' + "Table Name:           " + FormulaTableName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex >= 0)
            {
                this.Cursor = Cursors.WaitCursor;
                txtSelectedSection.Text = listBox2.SelectedItem.ToString().Trim();
                cboOffDaysSection.Text = txtSelectedSection.Text.Trim();
                cboOffDaysGang.Text = @"DUMMY";
                Base.Section = txtSelectedSection.Text.Trim();  

                //Start Threads
                Shared.evaluateDataTable(Base, "BONUSSHIFTS");
                Shared.evaluateDataTable(Base, "CLOCKEDSHIFTS");
                Shared.extractListOfTableNames(Base);

                label15.Text = listBox2.SelectedItem.ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() +
                "' and period = '" + BusinessLanguage.Period + "'";                                  //xxxxxxxxxxxxxxx
                loadMO();
 
                evaluateStatus();

                evaluateParticipants();
                evaluateEmployeePenalties();
                evaluateFactors();
                evaluateOffDays();
                evaluateRates();
                evaluateCrews(); 
                evaluateMineParameters();
                extractMeasuringDates();

                evaluateShifts();

                this.Cursor = Cursors.Arrow; 
            }
        }

        private void evaluateShifts()
        {

            System.Threading.Thread.Sleep(3000);

            evaluateClockedShifts();
            evaluateLabour();
        }

        private void extractMeasuringDates()
        {

            IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                          where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                          where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();
            dateTimePicker1.Value = Convert.ToDateTime(temp.Rows[0]["FSH"].ToString().Trim());
            dateTimePicker2.Value = Convert.ToDateTime(temp.Rows[0]["LSH"].ToString().Trim());
            strMonthShifts = temp.Rows[0]["MONTHSHIFTS"].ToString().Trim();

            lstOffDayValue.Items.Clear();
            //Load the possible dates that the user can select in this measuring period for the offday calendar
            for (DateTime i = dateTimePicker1.Value; i <= dateTimePicker2.Value; i = i.AddDays(1))
            {
                lstOffDayValue.Items.Add(i.ToString("yyyy-MM-dd"));
            }


        }

        private void btnEmployeeCalc_Click(object sender, EventArgs e)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        private void dataSort_Click(object sender, EventArgs e)
        {

        }

        private void DataPrintCrewPrint_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {

                case "tabFactors":
                    #region tabFactors

                    //HJ
                    if (cboVarName.Text.Trim().Length != 0 && txtVarValue.Text.Trim().Length != 0)
                    {
                        intRow = grdFactors.CurrentCell.RowIndex;
                        intColumn = grdFactors.CurrentCell.ColumnIndex;

                        if (grdFactors[0, intRow].Value.ToString().Trim() != "XXX")
                        {

                            strSQL = "BEGIN transaction; Update Factors set VarName = '" + cboVarName.Text.Trim() +
                                             "', VarValue = '" + txtVarValue.Text.Trim() + "' Where " +
                                             " VarName = '" + grdFactors["VARNAME", intRow].Value.ToString().Trim() +
                                             "' and VarValue = '" + grdFactors["VARVALUE", intRow].Value.ToString().Trim() +
                                             "' and period = '" + BusinessLanguage.Period + "' ;Commit Transaction;";


                            grdFactors["VARNAME", intRow].Style.BackColor = Color.LightBlue;
                            grdFactors["VARNAME", intRow].Value = cboVarName.Text.Trim();
                            grdFactors["VARVALUE", intRow].Style.BackColor = Color.LightBlue;
                            grdFactors["VARVALUE", intRow].Value = txtVarValue.Text.Trim();

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdFactors[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("FACTORS", "U - Update", s, dictGridValues[s], grdFactors[s, intRow].Value.ToString().Trim());

                                }

                            }

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    //HJ
                    if (cboEmplPenEmployeeNo.Text.Trim().Length != 0 &&
                        txtPenaltyValue.Text.Trim().Length != 0 && cboPenaltyInd.Text.Trim().Length != 0)
                    {

                        intRow = grdEmplPen.CurrentCell.RowIndex;
                        intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                        if (cboEmplPenEmployeeNo.Text.Contains("-"))
                        {
                            strName = cboEmplPenEmployeeNo.Text.Substring(0, cboEmplPenEmployeeNo.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboEmplPenEmployeeNo.Text.Trim();
                        }

                        strSQL = "BEGIN transaction; Update EmployeePenalties set Period = '" + txtPeriod.Text.Trim() +
                                             "', Employee_No = '" + strName + "', PenaltyValue = '" + txtPenaltyValue.Text.Trim() +
                                             "', PenaltyInd = '" + cboPenaltyInd.Text.Trim() + "'" +
                                             " Where Section = '" + grdEmplPen["SECTION", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdEmplPen["PERIOD", intRow].Value.ToString().Trim() +
                                             "' and Employee_No = '" + grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                             "' and PenaltyValue = '" + grdEmplPen["PENALTYVALUE", intRow].Value.ToString().Trim() +
                                             "' and PenaltyInd = '" + grdEmplPen["PENALTYIND", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXXXXXXXXXXX")
                        {
                            grdEmplPen["Section", intRow].Value = txtSelectedSection.Text.Trim();
                            grdEmplPen["Section", intRow].Style.BackColor = Color.LightBlue;
                            grdEmplPen["Period", intRow].Value = txtPeriod.Text.Trim();
                            grdEmplPen["Period", intRow].Style.BackColor = Color.LightBlue;
                            grdEmplPen["Employee_No", intRow].Value = cboEmplPenEmployeeNo.Text.Trim();
                            grdEmplPen["Employee_No", intRow].Style.BackColor = Color.LightBlue;
                            grdEmplPen["PenaltyValue", intRow].Value = txtPenaltyValue.Text.Trim();
                            grdEmplPen["PenaltyValue", intRow].Style.BackColor = Color.LightBlue;
                            grdEmplPen["PenaltyInd", intRow].Value = cboPenaltyInd.Text.Trim();
                            grdEmplPen["PenaltyInd", intRow].Style.BackColor = Color.LightBlue;

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            clearAllCalcValues("Ganglink", txtSelectedSection.Text.Trim());
                            clearAllCalcValues("Miners", txtSelectedSection.Text.Trim());
                            clearAllCalcValues("Bonusshifts", txtSelectedSection.Text.Trim());
                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdEmplPen[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("EmplPen", "U - Update", s, dictGridValues[s], grdEmplPen[s, intRow].Value.ToString().Trim());

                                }

                            }

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabConfig":
                    #region tabConfiguration

                    //HJ
                    if (grdConfigs[0, intRow].Value.ToString().Trim() != "XXX")
                    {
                        if (cboParameterName.Text.Trim().Length != 0 && cboParm1.Text.Trim().Length != 0 &&
                            cboParm2.Text.Trim().Length != 0 && cboParm3.Text.Trim().Length != 0 &&
                            cboParm4.Text.Trim().Length != 0 && cboParm5.Text.Trim().Length != 0 &&
                            cboParm6.Text.Trim().Length != 0 && cboParm7.Text.Trim().Length != 0)
                        {

                            intRow = grdConfigs.CurrentCell.RowIndex;
                            intColumn = grdConfigs.CurrentCell.ColumnIndex;

                            InputBoxResult intresult = InputBox.Show("Password: ");

                            if (intresult.ReturnCode == DialogResult.OK)
                            {
                                if (intresult.Text.Trim() == "Moses")
                                {

                                    General.updateConfigsRecord(Base.BaseConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                                     cboParameterName.Text.Trim(), cboParm1.Text.Trim(), cboParm2.Text.Trim(), cboParm3.Text.Trim(), cboParm4.Text.Trim(),
                                     cboParm5.Text.Trim(), cboParm6.Text.Trim(), cboParm7.Text.Trim(), grdConfigs["ParameterName", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm1", intRow].Value.ToString().Trim(), grdConfigs["Parm2", intRow].Value.ToString().Trim(),
                                     grdConfigs["Parm3", intRow].Value.ToString().Trim(), grdConfigs["Parm4", intRow].Value.ToString().Trim());
                                    //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                    foreach (string s in lstTableColumns)
                                    {
                                        if (dictGridValues[s] == grdConfigs[s, intRow].Value.ToString().Trim())
                                        {

                                        }
                                        else
                                        {
                                            //Write out to audit log
                                            writeAudit("CONFIGURATION", "U - Update", s, dictGridValues[s], grdConfigs[s, intRow].Value.ToString().Trim());

                                        }

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Invalid password", "Error", MessageBoxButtons.OK);
                                }
                            }

                            grdConfigs["ParameterName", intRow].Value = cboParameterName.Text.Trim();
                            grdConfigs["ParameterName", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm1", intRow].Value = cboParm1.Text.Trim();
                            grdConfigs["Parm1", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm2", intRow].Value = cboParm2.Text.Trim();
                            grdConfigs["Parm2", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm3", intRow].Value = cboParm3.Text.Trim();
                            grdConfigs["Parm3", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm4", intRow].Value = cboParm4.Text.Trim();
                            grdConfigs["Parm4", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm5", intRow].Value = cboParm5.Text.Trim();
                            grdConfigs["Parm5", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm6", intRow].Value = cboParm6.Text.Trim();
                            grdConfigs["Parm6", intRow].Style.BackColor = Color.LightBlue;
                            grdConfigs["Parm7", intRow].Value = cboParm7.Text.Trim();
                            grdConfigs["Parm7", intRow].Style.BackColor = Color.LightBlue;

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    //HJ
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {

                        InputBoxResult result = InputBox.Show("Password: ", "Rates Inputs are Password Protected!", "*", "0");

                        if (result.ReturnCode == DialogResult.OK)
                        {
                            if (result.Text.Trim() == "Moses")
                            {
                                intRow = grdRates.CurrentCell.RowIndex;
                                intColumn = grdRates.CurrentCell.ColumnIndex;

                                General.updateRatesRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                             txtBonusType.Text.Trim(),
                                                             txtPeriod.Text.ToString().Trim(), txtRateType.Text.Trim(), txtLowValue.Text.Trim(),
                                                             txtHighValue.Text.Trim(), txtRate.Text.Trim(),
                                                             grdRates["Low_Value", intRow].Value.ToString().Trim(), grdRates["High_Value", intRow].Value.ToString().Trim(),
                                                             grdRates["Rate", intRow].Value.ToString().Trim());
                                Application.DoEvents();

                                MessageBox.Show("All calculations will becleared.  Recalculations have to be done.", "Information", MessageBoxButtons.OK);
                                clearAllCalcValues("Ganglink", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Miners", txtSelectedSection.Text.Trim());
                                clearAllCalcValues("Bonusshifts", txtSelectedSection.Text.Trim());

                                grdRates["Low_Value", intRow].Value = txtLowValue.Text.Trim();
                                grdRates["Low_Value", intRow].Style.BackColor = Color.LightBlue;
                                grdRates["High_Value", intRow].Value = txtHighValue.Text.Trim();
                                grdRates["High_Value", intRow].Style.BackColor = Color.LightBlue;
                                grdRates["Rate", intRow].Value = txtRate.Text.Trim();
                                grdRates["Rate", intRow].Style.BackColor = Color.LightBlue;

                                //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                foreach (string s in lstTableColumns)
                                {
                                    if (dictGridValues[s] == grdRates[s, intRow].Value.ToString().Trim())
                                    {

                                    }
                                    else
                                    {
                                        //Write out to audit log
                                        writeAudit("RATES", "U - Update", s, dictGridValues[s], grdRates[s, intRow].Value.ToString().Trim());

                                    }

                                }

                            }
                            else
                            {
                                MessageBox.Show("Invalid Password.", "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabParticipants":
                    #region tabParticipants

                    if (cboParticipantsCrew.Text.Trim().Length > 0)
                    {
                        intRow = grdParticipants.CurrentCell.RowIndex;

                        strSQL = "Update PARTICIPANTS set " +
                                     " Shifts_Worked = '" + txtParticipatantsShiftsWorked.Text.Trim() +
                                     "' , Awop_Shifts = '" + txtParticipatantsAwopShifts.Text.Trim() +
                                     "' , Gang = '" + txtAutoDGang.Text.Trim() +
                                     "' , Crew = '" + cboParticipantsCrew.Text.Trim() +
                                     "' , Crewtype = '" + cboParticipantsCrewType.Text.Trim() +
                                     "' where " +
                                     " Shifts_Worked = '" + grdParticipants["SHIFTS_WORKED", intRow].Value +
                                     "' and Awop_Shifts = '" + grdParticipants["AWOP_SHIFTS", intRow].Value +
                                     "' and GANG = '" + grdParticipants["Gang", intRow].Value +
                                     "' and CREW = '" + grdParticipants["Crew", intRow].Value +
                                     "' and CREWTYPE = '" + grdParticipants["Crewtype", intRow].Value +
                                     "' and EMPLOYEE_No = '" + grdParticipants["Employee_no", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdParticipants["GANG", intRow].Value = txtAutoDGang.Text.Trim();
                        grdParticipants["CREW", intRow].Value = cboParticipantsCrew.Text.Trim();
                        grdParticipants["CREWTYPE", intRow].Value = cboParticipantsCrewType.Text.Trim();
                        grdParticipants["SHIFTS_WORKED", intRow].Value = txtParticipatantsShiftsWorked.Text.Trim();
                        grdParticipants["AWOP_SHIFTS", intRow].Value = txtParticipatantsAwopShifts.Text.Trim();

                        for (int i = 0; i <= grdParticipants.Columns.Count - 1; i++)
                        {
                            grdParticipants[i, intRow].Style.BackColor = Color.Plum;
                        }

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdParticipants[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("PARTICIPANTS", "U - Update", s, dictGridValues[s], grdParticipants[s, intRow].Value.ToString().Trim());

                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Please fill all input boxes.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabCrews":
                    #region tabCrews

                    //HJ
                    if (cboCrew.Text.Trim().Length != 0 &&
                        cboCrewLinkingGang.Text.Trim().Length != 0)
                    {

                        intRow = grdCrews.CurrentCell.RowIndex;
                        intColumn = grdCrews.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update Crews set Gang = '" + cboCrewLinkingGang.Text.Trim() +
                                             "', Crew = '" + cboCrew.Text.Trim() +
                                             "', CrewType = '" + cboCrewType.Text.Trim() +
                                             "', Tons = '" + txtTons.Text.Trim() +
                                             "', SafetyInd = '" + cboSafetyInd.Text.Trim() +
                                             "' Where Gang = '" + grdCrews["GANG", intRow].Value.ToString().Trim() +
                                             "' and Crew = '" + grdCrews["CREW", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdCrews["PERIOD", intRow].Value.ToString().Trim() +
                                             "';Commit Transaction;";

                        grdCrews["GANG", intRow].Value = cboCrewLinkingGang.Text.Trim();
                        grdCrews["GANG", intRow].Style.BackColor = Color.LightBlue;
                        grdCrews["CREWTYPE", intRow].Value = cboCrewType.Text.Trim();
                        grdCrews["CREWTYPE", intRow].Style.BackColor = Color.LightBlue;
                        grdCrews["CREW", intRow].Value = cboCrew.Text.Trim();
                        grdCrews["CREW", intRow].Style.BackColor = Color.LightBlue;
                        grdCrews["TONS", intRow].Value = txtTons.Text.Trim();
                        grdCrews["TONS", intRow].Style.BackColor = Color.LightBlue;
                        grdCrews["SafetyInd", intRow].Value = cboSafetyInd.Text.Trim();
                        grdCrews["SafetyInd", intRow].Style.BackColor = Color.LightBlue;

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdCrews[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("Crews", "U - Update", s, dictGridValues[s], grdCrews[s, intRow].Value.ToString().Trim());

                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabMineParameters":
                    #region tabMineParameters

                    if (txtTons_Actual.Text.Trim().Length != 0)
                    {
                        intRow = grdMineParameters.CurrentCell.RowIndex;
                        intColumn = grdMineParameters.CurrentCell.ColumnIndex;

                        strSQL = " Update MineParameters set " +
                                 " Tons_Actual = '" + txtTons_Actual.Text.Trim() +
                                 "' Where section = '" + grdMineParameters["Section", intRow].Value.ToString().Trim() +
                                 "' and Tons_Actual = '" + grdMineParameters["Tons_Actual", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdMineParameters["Period", intRow].Value.ToString().Trim() +
                                 "'; ";

                        strSQL = strSQL.Trim() + " Update Crews set " +
                                 " Tons = '" + txtTons_Actual.Text.Trim() +
                                 "' Where section = '" + grdMineParameters["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdMineParameters["Period", intRow].Value.ToString().Trim() +
                                 "' and Crewtype = 'InDirect'" ;

                        TB.InsertData(Base.DBConnectionString, strSQL);     
                        grdMineParameters["Tons_Actual", intRow].Value = txtTons_Actual.Text.Trim();

                        for (int i = 0; i <= grdMineParameters.Columns.Count - 1; i++)
                        {
                            grdMineParameters[i, intRow].Style.BackColor = Color.LightBlue;
                        }

                        grdMineParameters.FirstDisplayedScrollingRowIndex = intRow;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    evaluateMineParameters();

                    break;
                    #endregion

            }
        }

        private void writeAudit(string tablename, string function, string fieldname, string oldValue, string newValue)
        {
            string PK = string.Empty;
            foreach (string key in dictPrimaryKeyValues.Keys)
            {
                PK = PK + "<" + key.Trim() + "=" + dictPrimaryKeyValues[key] + ">";
            }

            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "AUDIT");
            audit.Clear();

            DataRow dr = audit.NewRow();
            dr["Type"] = function.Substring(0, 1);
            dr["TableName"] = tablename;
            dr["PK"] = PK;
            dr["FieldName"] = fieldname;
            dr["OldValue"] = oldValue;
            dr["NewValue"] = newValue;
            dr["UpdateDate"] = DateTime.Today.ToLongDateString();
            dr["UserName"] = BusinessLanguage.Userid;

            audit.Rows.Add(dr);
            audit.AcceptChanges();

            TB.saveCalculations2(audit, Base.DBConnectionString, " where type = 'x'", "AUDIT");
        }

        private void clearAllCalcValues(string _Tablename, string _Section)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Update " + _Tablename + " set ");
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, _Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                sb.Append(row["CALC_NAME"].ToString() + " = '0',");
            }

            if (sb.Length > 25)
            {
                sb.Append(strWhere);

                string strTemp = Convert.ToString(sb.Replace(",where", " Where"));
                TB.InsertData(Base.DBConnectionString, strTemp);
            }
        }
        
        private void deleteAllColumns(string Tablename)
        {
            //xxxxxxxxxxxxxxxxxxx
            //Create the earnings table
            createTheFile(Tablename);

            //Add the calculation columns.
            createEarningsColumns(Tablename);

            List<string> lstColumnNames = new List<string>();

            //extract the latest data from the base file e.g. Ganglink, Bonusshifts and replace data in the earningsfile.

            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                           " where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

            //Give the tempory file a name
            tb.TableName = Tablename + "EARN" + BusinessLanguage.Period.Trim();

            if (Tablename.ToUpper() == "BONUSSHIFTS")
            {
                #region Remove columns starting with DAY from BONUSSHIFTS
                //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                        lstColumnNames.Add(dc.ColumnName.Trim());
                    }
                    else
                    {

                    }
                }

                foreach (string s in lstColumnNames)
                {
                    tb.Columns.Remove(s);
                    tb.AcceptChanges();
                }

                lstColumnNames.Clear();
                #endregion
            }

            //Save the data to be processed to the earnings table.
            TB.saveCalculations2(tb, Base.DBConnectionString, " where section = '" + txtSelectedSection.Text.Trim() + "'",
                                 tb.TableName.Trim());

            Application.DoEvents();
            //}
        }

        private void createTheFile(string Tablename)
        {
            //Check if earningstable exist - e.g. GangLinkEarn201108....if not...CREATE the table
            List<string> lstColumnNames = new List<string>();

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period.Trim());

            if (intCount > 0)
            {
            }
            else
            {
                //CREATE the earnings table:  GanglinkEarn201108
                //Extract the table into a temp file from the datafile e.g. GANGLINK, BONUSSHIFTS, DRILLERS etc.

                DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                               "where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

                //Give the tempory file a name
                tb.TableName = Tablename + "Earn" + BusinessLanguage.Period.Trim();

                if (Tablename.ToUpper() == "BONUSSHIFTS")
                {
                    #region Remove columns starting with DAY from BONUSSHIFTS
                    //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                    foreach (DataColumn dc in tb.Columns)
                    {
                        if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                        {
                            lstColumnNames.Add(dc.ColumnName.Trim());
                        }
                        else
                        {

                        }
                    }

                    foreach (string s in lstColumnNames)
                    {
                        tb.Columns.Remove(s);
                        tb.AcceptChanges();
                    }

                    lstColumnNames.Clear();
                    #endregion
                }

                strSqlAlter.Remove(0, strSqlAlter.Length);

                //First create the base table.  Why, because all these columns should be NOT NULL.  
                //The Formulas SHOULD be NULL when created
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                    }
                    else
                    {
                        lstColumnNames.Add(dc.ColumnName);
                    }
                }

                //Create the earningstable e.g. BONUSSHIFTSEARN201108T

                TB.createEarningsTable(Base.DBConnectionString, tb.TableName, Tablename, lstColumnNames);

            }
        }

        private void createEarningsColumns(string Tablename)
        {
            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period);

            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(Base.DBName + BusinessLanguage.Period,
                                      Tablename + "EARN", Base.AnalysisConnectionString);

            foreach (DataRow row in tableformulas.Rows)
            {
                if (tb.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                }
                else
                {
                    strSqlAlter = strSqlAlter.Append(" ; Alter table " + Tablename + "EARN" + BusinessLanguage.Period + " add " +
                                                     row["CALC_NAME"].ToString().Trim() + " varchar(50) NULL");
                }
            }

            if (strSqlAlter.ToString().Trim().Length > 0)
            {
                StringBuilder bld = new StringBuilder();
                bld.Append("BEGIN transaction;" + strSqlAlter.ToString().Substring(1).Trim() + ";COMMIT transaction;");
                TB.InsertData(Base.DBConnectionString, bld.ToString().Trim());
                Application.DoEvents();
            }
            else
            {
            }
        }

        private void deleteAllCalcColumns(string Tablename)
        {
            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
            }
        }

        private void deleteAllCalcColumns(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
                }

            }
        }

        private void deleteAllCalcColumnsFromTempTable(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    Table.Columns.Remove(row["CALC_NAME"].ToString().Trim());
                }
            }

            Table.AcceptChanges();
        }

        private void Calcs(string tablename, string phasename, string Delete)
        {
            if (Delete == "Y")
            {
                deleteAllColumns(tablename);
            }

            TB.insertProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period, tablename + "EARN", phasename, txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "N", "N", (string)DateTime.Now.ToLongTimeString(), Convert.ToString(++intProcessCounter));

        }

        private void openTab(TabPage tp)
        {
            this.tabInfo.SelectedTab = tp;

            Application.DoEvents();

        }

        private void executeCostSheetFormulas(string TableName)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);
            string strprevPeriod = TableName;
            strSQL = "BEGIN transaction; insert into monitor values('" + Base.DBName + "','" + strprevPeriod + "','N','0','" + txtSelectedSection.Text.Trim() + "','0','0'); commit transaction; ";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        #region Open Tabs

        private void btnLockCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

        private void btnLockBonusShifts_Click(object sender, EventArgs e)
        {
            openTab(tabLabour);
        }

        private void btnLockGangLink_Click(object sender, EventArgs e)
        {
            openTab(tabParticipants);
        }

        private void btnLockOffday_Click(object sender, EventArgs e)
        {
            openTab(tabFactors);
        }

        private void btnLockEmplPen_Click(object sender, EventArgs e)
        {
            openTab(tabEmplPen);
        }
        #endregion

        private void btnDeleteRow_Click_1(object sender, EventArgs e)
        {

            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {

                case "tabEmplPen":
                    #region tabEmployeePenalty

                    intRow = grdEmplPen.CurrentCell.RowIndex;
                    intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                    if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from EmployeePenalties " +
                                 " Where Section = '" + grdEmplPen["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdEmplPen["Period", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdEmplPen["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and Workplace = '" + grdEmplPen["Workplace", intRow].Value.ToString().Trim() +
                                 "' and PenaltyInd = '" + grdEmplPen["PenaltyInd", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateEmployeePenalties();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabParticipants":
                    #region tabParticipants

                    if (
                        txtAutoDGang.Text.Trim().Length > 0 &&
                        txtAutoEmployee.Text.Trim().Length > 0)
                    {
                        intRow = grdParticipants.CurrentCell.RowIndex;


                        strSQL = "delete from Participants where " +
                                 "Shifts_Worked = '" + txtParticipatantsShiftsWorked.Text.Trim() +
                                 "' and Awop_Shifts = '" + txtParticipatantsAwopShifts.Text.Trim() +
                                 "' and GANG = '" + txtAutoDGang.Text.Trim() +
                                 "' and CREW = '" + grdParticipants["Crew", intRow].Value +
                                 "' and EMPLOYEE_No = '" + grdParticipants["Employee_no", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateParticipants();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

                case "tabCrews":
                    #region tabCrews



                    strSQL = "delete from crews where " +
                             " Gang = '" + cboCrewLinkingGang.Text.Trim() +
                             "' and Crew = '" + cboCrew.Text.Trim() +
                             "' and CrewType = '" + cboCrewType.Text.Trim() +
                             "' and Tons = '" + txtTons.Text.Trim() +
                             "' and SafetyInd = '" + cboSafetyInd.Text.Trim() + "'";

                    TB.InsertData(Base.DBConnectionString, strSQL);
                    evaluateCrews();

                    break;
                    #endregion

                case "tabOffdays":
                    #region tabOffdays

                    intRow = grdOffDays.CurrentCell.RowIndex;

                    if (cboOffDaysGang.Text.Trim().Length > 0 &&
                        cboOffDaysSection.Text.Trim().Length > 0)
                    {
                        strSQL = "delete from Offdays where gang = '" + grdOffDays["Gang", intRow].Value +
                                 "' and section = '" + grdOffDays["Section", intRow].Value +
                                 "' and period = '" + BusinessLanguage.Period +
                                 "' and OffDayValue = '" + grdOffDays["OffdayValue", intRow].Value + "'";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateOffDays();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion

            }
        }

        protected virtual void FrontDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteFullBeginTag("HTML");
            writer.WriteFullBeginTag("Head");
            writer.RenderBeginTag(System.Web.UI.HtmlTextWriterTag.Style);
            writer.Write("<!--");

            StreamReader sr = File.OpenText(strServerPath + ":\\koos.html");
            String input;
            while ((input = sr.ReadLine()) != null)
            {
                writer.WriteLine(input);
            }
            sr.Close();
            writer.Write("-->");
            writer.RenderEndTag();
            writer.WriteEndTag("Head");
            writer.WriteFullBeginTag("Body");
        }
                    
        protected virtual void RearDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteEndTag("Body");
            writer.WriteEndTag("HTML");
        }
                    
        private void printHTML(DataTable dt, string TabName)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = "c:\\icalc\\koos.html";
                try
                {

                    StreamWriter SW = new StreamWriter(OPath);
                    //StringWriter SW = new StringWriter();
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {

                            HTMLWriter.WriteLine("HARMONY - Kusasalethu Mine - " + TabName);
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("==============================");
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();

                            grid.RenderControl(HTMLWriter);
                            //RearDecorator(HTMLWriter);

                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();


                    System.Diagnostics.Process P = new System.Diagnostics.Process();
                    P.StartInfo.WorkingDirectory = strServerPath + ":\\Program Files\\Internet Explorer";
                    P.StartInfo.FileName = "IExplore.exe";
                    P.StartInfo.Arguments = "c:\\icalc\\koos.html";
                    P.Start();
                    P.WaitForExit();


                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }
                    
        private void btnLoad_Click_1(object sender, EventArgs e)
        {
            if (listBox3.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select the number of measuring shifts", "Information", MessageBoxButtons.OK);
            }
            else
            {
                if (txtSelectedSection.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Please select a section and the correct month measuring shifts for the section.", "Information", MessageBoxButtons.OK);
                }
                else
                {
                    string selectedSection = txtSelectedSection.Text.Trim();
                    string grdSection = grdCalendar["SECTION", intFiller].Value.ToString().Trim();
                    if (selectedSection == grdSection)
                    {
                        Base.updateCalendarRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                         txtBonusType.Text.Trim(), txtSelectedSection.Text.Trim(),
                                                         txtPeriod.Text.ToString().Trim(),
                                                         (Convert.ToDateTime(dateTimePicker1.Text)).ToString("yyyy-MM-dd"),
                                                         (Convert.ToDateTime(dateTimePicker2.Text)).ToString("yyyy-MM-dd"),
                                                         listBox3.SelectedItem.ToString().Trim());
                        Application.DoEvents();
                    }

                    else
                    {
                        MessageBox.Show("Selected section not the same as grid section.", "Informations", MessageBoxButtons.OK);
                    }

                    //Extract Calendar again and insert into 
                    Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");
                    grdCalendar.DataSource = Calendar;

                    extractMeasuringDates();
                }
            }
        }

        private void label83_Click(object sender, EventArgs e)
        {
            extractDBTableNames(listBox1);
        }

        private void btnLockPaysend_Click(object sender, EventArgs e)
        {
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();


        }

        private void btnEmployeeCostsheet_Click(object sender, EventArgs e)
        {
            Calcs("Miners", "Miners", "N");
        }

        private void btnPrint_Click_1(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            switch (tabInfo.SelectedTab.Name)
            {
              
                case "tabParticipants":
                    #region tabParticipants
                    if (Participants.Rows.Count > 0 && Participants.Columns.Count < 5)
                    {
                        if (Participants.Rows.Count > 0)
                        {

                            printHTML(Participants, "Participants");
                        }
                        else
                        {
                            MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                        }
                    }

                    else
                    {
                        dt = Base.extractPrintData(Base.DBConnectionString, "GangLink", strWhere);
                        deleteAllCalcColumns("Participants", dt);
                        dt.AcceptChanges();
                        if (dt.Rows.Count > 0)
                        {

                            printHTML(dt, "Participants");
                        }
                        else
                        {
                            MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                        }
                    }
                    break;
                    #endregion

               
                case "tabLabour":
                    #region tabLabour

                    dt = Base.extractPrintData(Base.DBConnectionString, "BonusShifts", strWhere);
                    deleteAllCalcColumns("BonusShifts", dt);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "BonusShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    dt = Base.extractPrintData(Base.DBConnectionString, "EmployeePenalties", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "EmployeePenalties");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }
                    break;
                    #endregion

                case "tabOffday":
                    #region tabOffdays

                    dt = Base.extractPrintData(Base.DBConnectionString, "Offdays", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Offdays");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabCalendar":
                    #region tabCalendar

                    dt = Base.extractPrintData(Base.DBConnectionString, "Calendar", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Calendar");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabClockShifts":
                    #region tabClockShifts

                    dt = Base.extractPrintData(Base.DBConnectionString, "ClockedShifts", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "ClockedShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Rates", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Rates");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

            

            }
        }

        private void calcStopeData()
        {
            Base.Period = txtPeriod.Text.Trim();
            //Base.Period = "200909";

            SqlConnection stopeConn = Base.StopeConnection;
            stopeConn.Open();

            try
            {
                DataTable ContractTotals = TB.getContractCrewOfficialBonus(Base.StopeConnectionString, "STOPING", txtSelectedSection.Text.Trim());

                stopeConn.Close();

                TB.updateDSShiftbossCrewBonus(Base.DBConnectionString, ContractTotals);
            }
            catch { }
        }

        private void startCalcProcess()
        {
            this.Cursor = Cursors.WaitCursor;
            btnx.Visible = true;
            btnx.Enabled = true;
            btnx.Text = "Run";
            TB.deleteProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period);
            //clear the monitor table
            TB.deleteAllExcept(Base.DBConnectionString, "Monitor");
            Calcs("BonusShifts", "BonusShiftsearn10", "Y");
            Calcs("BonusShifts", "BonusShiftsearn20", "N");
            btnBaseCalcs.BackColor = Color.Orange;
            Calcs("Participants", "Participantsearn08", "Y");  
            Calcs("Participants", "Participantsearn09", "N");  
            Calcs("Participants", "Participantsearn10", "N");
            Calcs("Participants", "Participantsearn20", "N");
            Calcs("Participants", "Participantsearn30", "N");
            btnSupportLinkCalc.BackColor = Color.Orange;
            Calcs("Participants", "Participantsearn50", "N");
            Calcs("Participants", "Participantsearn55", "N");
            Calcs("Participants", "Participantsearn60", "N");
            Calcs("Participants", "Participantsearn62", "N");
            Calcs("Participants", "Participantsearn65", "N");
            btnGangCalcs.BackColor = Color.Orange;
            Calcs("Participants", "Participantsearn67", "N");
            Calcs("Participants", "Participantsearn70", "N");
            Calcs("Participants", "Participantsearn80", "N");
            Calcs("Participants", "Participantsearn90", "N"); 
            Calcs("Exit", "Exit", "N");
            btnBonusShiftsCalcs.BackColor = Color.Orange;

            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "ParticipantsEarn10", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), ""); ;
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "ParticipantsEarn50", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "ParticipantsEarn60", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "ParticipantsEarn90", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Exit", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");

            //Base.backupDatabase3(Base.DBConnectionString, Base.DBName, Base.BackupPath);
            this.Cursor = Cursors.Arrow;
        }

        private void btnBaseCalcsHeader_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                //Check if the a calculator is currently running
                Int16 intCount1 = TB.checkTableExist(Base.DBConnectionString, "BonusShiftsEARN");
                Int16 intCount2 = TB.checkTableExist(Base.DBConnectionString, "ParticipantsEARN");
                Int16 intCount3 = TB.checkTableExist(Base.DBConnectionString, "SupportLinkEARN");
                Int16 intCount4 = TB.checkTableExist(Base.DBConnectionString, "DrillersEARN");
                Int16 intCount5 = TB.checkTableExist(Base.DBConnectionString, "MinersEARN");
                Int16 intCount6 = TB.checkTableExist(Base.DBConnectionString, "SectionEarningsEARN");

                if (intCount1 > 0 || intCount2 > 0 || intCount3 > 0 || intCount4 > 0 || intCount5 > 0 || intCount6 > 0)
                {
                    MessageBox.Show("A calculator is currently running for this bonus scheme: " + BusinessLanguage.MiningType +
                                    " " + BusinessLanguage.BonusType);
                }
                else
                {
                    startCalcProcess();

                }

            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
        }

        private void grdActiveSheet_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int columnnr = e.ColumnIndex;
                DialogResult result = MessageBox.Show("Do you want to delete the column:  " + grdActiveSheet.Columns[columnnr].HeaderText + "?", "INFORMATION", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    //int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
                    TB.removeColumn(Base.DBConnectionString, TB.TBName, grdActiveSheet.Columns[columnnr].HeaderText);
                    DoDataExtract("");
                    grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);
                }
                else
                {
                    if (listBox1.SelectedItem.ToString().Trim() == "MONITOR")
                    {

                        string strSQL = "Begin transaction; Delete from monitor; commit transaction";
                        TB.InsertData(Base.DBConnectionString, strSQL);
                        Application.DoEvents();

                    }
                }
            }

            else
            {
                AConn = Analysis.AnalysisConnection;
                AConn.Open();
                DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

                foreach (DataRow dt in tempDataTable.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }



            }
        }

        private void grdCalendar_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["LSH"].ToString().Trim());
                intFiller = e.RowIndex;
            }

            dictPrimaryKeyValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdCalendar[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            dictGridValues.Clear();

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdCalendar[s, e.RowIndex].Value.ToString().Trim());
                }
            }

        }

        private void grdRates_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if (grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;

                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                    btnInsertRow.Enabled = true;
                }

                txtRateType.Text = grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim();
                txtLowValue.Text = grdRates["LOW_VALUE", e.RowIndex].Value.ToString().Trim();
                txtHighValue.Text = grdRates["HIGH_VALUE", e.RowIndex].Value.ToString().Trim();
                txtRate.Text = grdRates["RATE", e.RowIndex].Value.ToString().Trim();
            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictPrimaryKeyValues.Add(s, grdRates[s, e.RowIndex].Value.ToString().Trim());
                }
            }

            foreach (string s in lstTableColumns)
            {
                if (e.RowIndex < 0)
                {
                }
                else
                {
                    dictGridValues.Add(s, grdRates[s, e.RowIndex].Value.ToString().Trim());
                }
            }
            #endregion

        }

        private void payrollSend_Click(object sender, EventArgs e)
        {
            
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();
            //}
            //}

        }

        private void emailInfo_Click(object sender, EventArgs e)
        {

        }

        private void basicGraph_Click(object sender, EventArgs e)
        {

        }

        private void drillDownGraph_Click(object sender, EventArgs e)
        {

        }

        private void dataFilter_Click(object sender, EventArgs e)
        {
            if (General.textTestSQL.ToString().Trim().Length > 0)
            {
                scrQuerySQL testsql = new scrQuerySQL();
                testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
                testsql.Show();
            }
            else
            {
                MessageBox.Show("No SQL to pass", "Information", MessageBoxButtons.OK);
            }
        }

        private void dataPrintTables_Click(object sender, EventArgs e)
        {

        }

        private void dataFormulasImportTable_Click(object sender, EventArgs e)
        {
           

        }

        private void TBCreateSpreadsheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (openDialog.ShowDialog() != DialogResult.OK) return;
                //grpData.Enabled = false;
                string filename = openDialog.FileName;
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                fs.Close();

                if (spreadsheet.WorkbookData.Tables.Count > 0)
                {
                    switch (string.IsNullOrEmpty(Base.DBName))
                    {
                        case true:
                            MessageBox.Show("Create or select a database.", "DATABASE NEEDED!", MessageBoxButtons.OK);
                            break;

                        case false:
                            saveTheSpreadSheetToTheDatabase();
                            MessageBox.Show("Successfully Uploaded.", "Information", MessageBoxButtons.OK);
                            break;
                        default:

                            break;
                    }
                }

              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read file: \n" + ex.Message);
            }
        }

        private void TBDeleteTable_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Delete table: " + TB.TBName + " ? ", "Confirm", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
                    extractDBTableNames(listBox1);
                    TB.deleteDataTableFromCollection(TB.DBName);
                    TB.TBName = "";
                    TBFormulas.Tablename = "";
                    loadInfo();
                    break;


                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteCalcColumns_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Confirm DELETE of calculated columns from table: " + TBFormulas.Tablename + "?", "", MessageBoxButtons.YesNo);

            switch (result1)
            {
                case DialogResult.Yes:

                    DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, TB.TBName, Base.AnalysisConnectionString);
                    foreach (DataRow row in tableformulas.Rows)
                    {
                        TB.removeColumn(Base.DBConnectionString, TB.TBName, row["CALC_NAME"].ToString());

                    }
                    loadInfo();
                    break;

                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteAllTables_Click(object sender, EventArgs e)
        {
            foreach (string s in listBox1.Items)
            {
                TB.TBName = s.Trim();
                bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
            }
            extractDBTableNames(listBox1);
            loadInfo();
        }

        private void DBCreate_Click(object sender, EventArgs e)
        {

        }

        private void createNewDatabase(string Databasename)
        {

        }

        private void DBBackup_Click(object sender, EventArgs e)
        {
            ////The database-tables and formulas will be stored on spreadsheets.

            //if (listBox1.Items.Count == 0)
            //{
            //    MessageBox.Show("No tables to backup", "Backup Failure", MessageBoxButtons.OK);
            //}
            //else
            //{
            //    foreach (string s in listBox1.Items)
            //    {
            //        TB.TBName = s.Trim();
            //        saveTheSpreadSheet();
            //    }
            //}

            ////Extract the formulas of the database
            //extractDatabaseFormulas();
            //TB.TBName = "";

            //Base.backupDatabase3(Base.DBConnectionString, Base.DBName, "D:\\iCalc\\Backups\\Databases");

            //MessageBox.Show("Backup Done to:  D:\\iCalc\\Backups\\Databases ", "Information", MessageBoxButtons.OK);
        }

        private void extractDatabaseFormulas()
        {

        }

        private void DBDeleteList_Click(object sender, EventArgs e)
        {

        }
                    
        private void listDB()
        {
        }

        private void DBList_Click(object sender, EventArgs e)
        {
        }

        private void evaluateStatusButtons()
        {
            btnInsertRow.Enabled = false;
            btnUpdate.Enabled = false;
            btnDeleteRow.Enabled = false;
            btnLoad.Enabled = false;
            btnPrint.Enabled = false;
            btnLock.Enabled = false;

            panelInsert.BackColor = Color.Cornsilk;
            panelUpdate.BackColor = Color.Cornsilk;
            panelDelete.BackColor = Color.Cornsilk;
            panelPreCalcReport.BackColor = Color.Cornsilk;
        }

        private void btnx_Click_1(object sender, EventArgs e)
        {

            btnx.Text = "Running";
            btnx.Enabled = false;
            btnRefresh.Visible = true;
            execute();
            refreshExecution();

        }

        private void refreshExecution()
        {
           
            calcTime.Enabled = true;
            
        }

        private void execute()
        {

            System.Diagnostics.Process P = new System.Diagnostics.Process();

            switch (BusinessLanguage.Env)
            {
                case "Production":
                    strName = "KusasalethuBacSerP";
                    P.StartInfo.WorkingDirectory = @"z:\Harmony\Kusasalethu\Production\Core";
                    P.StartInfo.FileName = strName + ".exe";


                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;

                case "Test":
                    strName = "KusasalethuT";
                    P.StartInfo.WorkingDirectory = "C:\\OEM2";
                    P.StartInfo.FileName = strName + ".exe";

                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;

                case "Development":

                    strName = "Bacser8000D";
                    P.StartInfo.WorkingDirectory = @"C:\iCalc\Harmony\ServerProjects\Kusasalethu\Core";
                    P.StartInfo.FileName = strName + ".exe";

                    pictBox.Visible = true;
                    pictBox2.Visible = true;
                    calcTime.Enabled = true;

                    P.Start();
                    P.Close();
                    break;
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            evaluateStatus();
            evaluateStatusButtons();
        }
        
        private void btnGangTypeAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STPTMGangtypeAuth");
            this.Cursor = Cursors.Arrow;
        }

        private void printST_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            //mm.StartReport("BacSer8000Teams");
            //this.Cursor = Cursors.Arrow;


            this.Cursor = Cursors.WaitCursor;
            Shared.metareportAutoWParameter("Section_value", txtSelectedSection.Text.Trim(), "BacSer8000Teams" + txtSelectedSection.Text.Trim() + ".PDF",
                                   "BacSer8000Teams", strMetaReportCode,  
                                   "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\");
            this.Cursor = Cursors.Arrow;

        }

        private void btnSupportAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            //mm.StartReport("BackfillAuth");

            Shared.metareportAutoWParameter("Section_value", txtSelectedSection.Text.Trim(), "BackfillAuth" + txtSelectedSection.Text.Trim() + ".PDF",
                                   "BackfillAuth", strMetaReportCode,  
                                   "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\");

            this.Cursor = Cursors.Arrow;
        }

        private void btnInStopeAuth_Click(object sender, EventArgs e)
        {
             this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            //mm.StartReport("BackfillAutoInStope");

            Shared.metareportAutoWParameter("Section_value", txtSelectedSection.Text.Trim(), "BackfillAutoInStope" + txtSelectedSection.Text.Trim() + ".PDF",
                                   "BackfillAutoInStope", strMetaReportCode,  
                                   "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\");

            this.Cursor = Cursors.Arrow;
        }
      
        private void TBExport_Click_1(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void btnChangePeriod_Click(object sender, EventArgs e)
        {
            //Gets the name of all open forms in application
            foreach (Form form in Application.OpenForms)
            {
                if (form is scrLogon)
                {
                    form.Show(); //Show the form
                    break;
                }
            }
            exitValue = 2;//Change exit value

            this.Close(); //Close the current window

        }

        private void scrTeamD_FormClosing(object sender, FormClosingEventArgs e)//jvdw
        {

            if (exitValue == 0)
            {
                DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        //this.Close();
                        //scrMain main = new scrMain();
                        //main.MainLoad(BusinessLanguage, DB, Survey, Labour, Miners, Designations, Occupations, Clocked, EmplList, EmplPen, Configs);
                        //main.ShowDialog();
                        myConn.Close();
                        AAConn.Close();
                        AConn.Close();
                        //this.Close();
                        exitValue = 1;
                        Application.Exit();
                        break;

                    case DialogResult.No:
                        e.Cancel = true;
                        break;
                }
                if (exitValue == 2)
                {
                    exitValue = 1;
                    this.Close();
                }
            }
        }

        private void btnAttendance_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            evaluateLabour();
            if (Labour.Rows.Count == 0)
            {
                MessageBox.Show("No Labour records to print for the section: " + txtSelectedSection.Text.Trim(),
                                "Information", MessageBoxButtons.OK);
            }
            else
            {
                DataTable temp = Labour.Copy();
                deleteAllCalcColumnsFromTempTable("BonusShifts", temp);

                temp.Columns.Remove("TMLEADERIND");
                temp.AcceptChanges();

                //create a view of calendar
                string strSQL = " Drop view CalendarV;";

                Base.InsertData(Base.DBConnectionString, strSQL);

                strSQL = "create view CalendarV as SELECT * from Calendar" +
                                " where period = '" + BusinessLanguage.Period.Trim() + "';";

                Base.InsertData(Base.DBConnectionString, strSQL);

                TB.createAttendanceTable_withPeriodandBussUnit(Base.DBConnectionString, temp);

                //MetaReportRuntime.App mm = new MetaReportRuntime.App();
                //mm.Init(strMetaReportCode);
                //mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
                //mm.StartReport("BACSERTA");

                Shared.metareportAutoWParameter("Section_value", txtSelectedSection.Text.Trim(), "BACSERTA_" + txtSelectedSection.Text.Trim() + ".PDF",
                                   "BACSERTA", strMetaReportCode,
                                   "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\");
 
            }
            this.Cursor = Cursors.Arrow;
        }
                    
        private void btnSearchEmployNr_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = true;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            txtSearchEmplyNr.Focus();
        }

        private void btnEmployName_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = true;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            txtSearchEmplName.Focus();
        }

        private void btnSearchGang_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = true;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            txtSearchGang.Focus();
        }

        private void txtSearchEmplyNr_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            searchEmplNr = txtSearchEmplyNr.Text.ToString();
            searchEmplName = "";
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod

        }

        private void txtSearchEmplName_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = txtSearchEmplName.Text.ToString();
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod

        }
                    
        private void txtSearchGang_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = "";
            searchEmplGang = txtSearchGang.Text.ToString();
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang, grdLabour); //Calls the metod
        }

        public void searchBonus(string nr, string name, string gang, DataGridView Grid)
        {
            //Sets the details passed to lower case
            nr = nr.ToLower();
            name = name.ToLower();
            gang = gang.ToLower();

            //Gets the length
            int nrLenght = nr.Length;
            int nameLenght = name.Length;
            int gangLenght = gang.Length;

            // Ensuring length are always 1 and not 0 as
            // "" can not be tested.
            if (nrLenght == 0)
            {
                nrLenght = 1;
            }
            if (nameLenght == 0)
            {
                nameLenght = 1;
            }
            if (gangLenght == 0)
            {
                gangLenght = 1;
            }

            //Iterate through all the rows in the grid
            for (int i = 0; i < Grid.Rows.Count - 1; i++)
            {
                //Gets the values of the grid in the different columns
                string nrColumn = Grid.Rows[i].Cells["Employee_No"].Value.ToString();  //Cells from grid count from left starting at 0
                string nameColumn = Grid.Rows[i].Cells[1].Value.ToString();
                string gangColumn = Grid.Rows[i].Cells["Gang"].Value.ToString();

                //Sets the values from grid to lowercase for testing
                nrColumn = nrColumn.ToLower();
                nameColumn = nameColumn.ToLower();
                gangColumn = gangColumn.ToLower();

                //Gets the same amount from the grid string as was entertered bty the user to 
                //ensure the string can be tested
                nrColumn = nrColumn.Substring(1, nrLenght);//Start at 1 to throw away the aphabetic nr
                nameColumn = nameColumn.Substring(0, nameLenght);
                gangColumn = gangColumn.Substring(0, gangLenght);

                //Compares the different strings
                if (nr == nrColumn) //Employee nr
                {
                    //Empty the string not used
                    nameColumn = "";
                    gangColumn = "";
                    Grid.ClearSelection(); // Clears all past selection
                    Grid.Rows[i].Selected = true; //Selects the current row
                    Grid.FirstDisplayedScrollingRowIndex = i; //Jumps automatically to the row
                    break; //breaks the loop
                }

                if (gang == gangColumn) //Gang
                {
                    nrColumn = "";
                    nameColumn = "";
                    Grid.ClearSelection();
                    Grid.Rows[i].Selected = true;
                    Grid.FirstDisplayedScrollingRowIndex = i;
                    break;
                }
            }
        }

        private void dataBonusShiftsFromClockedShifts_Click(object sender, EventArgs e)
        {
            InputBoxResult result = InputBox.Show("Import Shifts per Gang.  Gang Number: ", "Employees to import");

            if (result.ReturnCode == DialogResult.OK)
            {

                #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
                                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where section = '" +
                                txtSelectedSection.Text.Trim() + "' and Gang = '" + result.Text.Trim() + "'";

                BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

                if (BonusShifts.Rows.Count > 0)
                {
                    string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                    DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                    int intStartDay = Base.calcNoOfDays(CalendarFSH, Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));
                    int intEndDay = Base.calcNoOfDays(CalendarLSH, Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));
                    int intStopDay = 0;

                    //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
                    int intNoOfDays = Base.calcNoOfDays(Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()), Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));

                    if (intStartDay < 0)
                    {
                        //The calendarFSH falls outside the startdate of the sheet.
                        intStartDay = 0;
                    }
                    else
                    {

                    }

                    if (intEndDay < 0 && intEndDay < -44)
                    {
                        intStopDay = 0;
                    }
                    else
                    {
                        if (intEndDay < 0)
                        {
                            //the LSH of the measuring period falls within the spreadsheet
                            intStopDay = intNoOfDays + intEndDay;

                        }
                        else
                        {
                            //The LSH of the measuring period falls outside the spreadsheet
                            intStopDay = 44;
                        }


                        //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                        //were not imported.

                        #region count the shifts
                        //Count the the shifts

                        DialogResult result2 = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for gang " + result.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                        switch (result2)
                        {
                            case DialogResult.OK:
                               
                                strWhere = strWhere + " and gang = '" + result.Text.Trim() + "'";
                                extractAndCalcShifts(intStartDay, intStopDay);
                                
                                break;

                            case DialogResult.Cancel:
                                break;

                        }

                        #endregion

                #endregion

                    }
                }
            }
        }

        private void dataPrintFormulas_Click(object sender, EventArgs e)
        {
            DataTable dt = Base.dataPrintFormulasBonusShifts(Base.AnalysisConnectionString, Base.DBName, "Participants");
            if (dt.Rows.Count > 0)
            {
                printHTML(dt, "Formulas on PARTICIPANTS");
            }
            else
            {
                MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
            }
        }

        private void auditByTable_Click(object sender, EventArgs e)
        {
            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Audit", " where tablename = 'Ganglink'");
            string[] auditcolumns = new string[10];

            string test = audit.Rows[0]["PK"].ToString().Trim();
            int testlength = test.Length;

            for (int i = 0; i <= 9; i++)
            {
                int tstLength = test.IndexOf(">");
                if (tstLength != -1)
                {
                    auditcolumns[i] = test.Substring(0, tstLength).Replace("<", "").Trim();
                    test = test.Substring(test.IndexOf(">") + 1);
                }

            }





        }
                    
        private void btnEmplyeAudit_Click(object sender, EventArgs e)
        {


            #region extract the sheet name and FSH and LSH of the extract
            string FilePath = "C:\\iCalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\Data\\ADTeam_201004.xls";
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);

            #region remove invalid records
            // Delete records that does not conform to configurations
            //foreach (DataRow row in dt.Rows)
            //{
            //    if ((row["GANG NAME"].ToString().Substring(5, 1) == "A" || row["GANG NAME"].ToString().Substring(5, 1) == "B" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "C" || row["GANG NAME"].ToString().Substring(5, 1) == "D" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "E" || row["WAGE CODE"].ToString() == "245M003" ||
            //        row["WAGE CODE"].ToString() == "400M009" || row["WAGE CODE"].ToString() == "245M001" ||
            //        row["WAGE CODE"].ToString() == "246M004" || row["WAGE CODE"].ToString() == "400M009")
            //        && (row["GANG NAME"].ToString().Substring(0, 5) == txtSelectedSection.Text.Trim()))
            //    {
            //    }
            //    else
            //    {
            //        //row.Delete();
            //    }

            //}

            //dt.AcceptChanges();

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }

            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);
            noOFDay = intNoOfDays;

            if (intNoOfDays <= 44)
            {
                for (int j = intNoOfDays + 1; j <= 44; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;


            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "STOPE";
                if (row["GANG"].ToString().Length > 0)
                {
                    row["SECTION"] = row["GANG"].ToString().Substring(0, 5);
                }
                else
                {
                    row["SECTION"] = "XXX";
                }

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion

            //Write to the database
            // TB.saveCalculations2(dt, Base.DBConnectionString, strWhere, "CLOCKEDSHIFTS");

            // Application.DoEvents();

            // grdClocked.DataSource = dt;
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            //string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
            //                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where (section = '"
            //                + txtSelectedSection.Text.Trim() + "' or WAGE_DESCRIPTION = 'STOPER')";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";//jvdw

            // BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);//jvdw laai die hele clockedshift table

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;//jvdw
            sheetlhs = SheetLSH;//jvdw
            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                // DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

                #region Extract the ganglinking of the current section
                ////Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                ////before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                ////This is done in the methord extractGangLink

                //DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GANGLINK", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current ganglinking for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractGangLink();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractGangLink();
                //}

                //cboMinersGangNo.Items.Clear();
                //lstNames = TB.loadDistinctValuesFromColumn(Labour, "Gang");
                //if (lstNames.Count > 1)
                //{

                //    foreach (string s in lstNames)
                //    {
                //        if (cboMinersGangNo.Items.Contains(s))
                //        { }
                //        else
                //        {
                //            cboMinersGangNo.Items.Add(s.Trim());
                //        }
                //    }
                //}

                #endregion

                #region Extract the miners of the current section
                //Remember a previous section could have been imported and calculated.  Therefore a delete can not be done on the table
                //before checking.  If a calc has run on the table, the insert must be updated with the necessary calc columns.
                //This is done in the method extractMiners

                //temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MINERS", strWhere);

                //if (temp.Rows.Count > 0)
                //{
                //    result = MessageBox.Show("Do you want to REPLACE the current MINERS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //    switch (result)
                //    {
                //        case DialogResult.OK:
                //            extractMiners();
                //            break;

                //        case DialogResult.Cancel:
                //            break;

                //    }
                //}
                //else
                //{
                //    extractMiners();
                //}
                #endregion

                fillFixTable(fixShifts, sheetfhs, sheetlhs, intNoOfDays, intStartDay, intStopDay);
                this.Cursor = Cursors.Arrow;
                //}
            }

        }

        public void fillFixTable(DataTable clockedTable, DateTime SheetFSH, DateTime SheetLSH, int intNoOfDays, int DayStart, int DayEnd)//jvdw
        {
            //Calculate the shifts in the clockedshifts table and insert all in a fixed
            //table that cannot be changed by the user!

            string SQLTable = "IF OBJECT_ID(N'emplshiftfix', N'U')IS NOT NULL DROP TABLE EMPLSHIFTFIX create table EMPLSHIFTFIX (employeeno char(20),shiftsfix char(20)) truncate table EMPLSHIFTFIX";
            Base.VoidQuery(Base.DBConnectionString, SQLTable);

            #region Calculate the shifts per employee en output to bonusshifts

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -44)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 44;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                int intSubstringLength = 0;
                int intShiftsWorked = 0;
                int intAwopShifts = 0;
                int shiftsCheck = 0;
                StringBuilder sqlInsertFixShifts = new StringBuilder("BEGIN TRANSACTION; ");

                foreach (DataRow row in clockedTable.Rows)
                {
                    foreach (DataColumn column in clockedTable.Columns)
                    {
                        if ((column.ColumnName.Substring(0, 3) == "DAY"))
                        {

                            if (column.ColumnName.ToString().Length == 4)
                            {
                                intSubstringLength = 1;
                            }
                            else
                            {
                                intSubstringLength = 2;
                            }

                            if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                               Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                            {

                                if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" || row[column].ToString().Trim() == "q" || row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w")
                                {
                                    intShiftsWorked = intShiftsWorked + 1;
                                    shiftsCheck = 1;
                                }
                                else
                                {
                                    if (row[column].ToString().Trim() == "A")
                                    {
                                        intAwopShifts = intAwopShifts + 1;
                                    }
                                    else { }

                                }
                            }
                            else
                            {
                                row[column] = "*";
                            }
                        }
                        else
                        {
                            if (column.ColumnName == "BONUSTYPE")
                            {
                                row["BONUSTYPE"] = "TEAM";
                            }
                        }
                    }//foreach datacolumn

                    row["SHIFTS_WORKED"] = intShiftsWorked;

                    string emplNr = row["employee_no"].ToString();
                    workedShiftsFixedClockedShift = intShiftsWorked;
                    intShiftsWorked = 0;
                    intAwopShifts = 0;
                    if (shiftsCheck == 1)
                    {
                        sqlInsertFixShifts.Append("INSERT INTO EMPLSHIFTFIX VALUES ('" + emplNr.Trim() + "','" + workedShiftsFixedClockedShift.ToString().Trim() + "');");
                    }
                }

                sqlInsertFixShifts.Append(" COMMIT TRANSACTION");


                Base.VoidQuery(Base.DBConnectionString, sqlInsertFixShifts.ToString());

                //DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

            }
        }

                    
        private void hideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grdActiveSheet.Columns[columnnr].Visible = false;


        }

        private void btnSurveySummary_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("SurveySumSTM");
            this.Cursor = Cursors.Arrow;
        }

        private void calcTime_Tick(object sender, EventArgs e)
        {
            btnRefresh_Click("Method", null);
        }

        private bool createZipFolder(string path, string databasename)
        {
            path = Base.BackupPath.Replace(Base.BackupPath.Substring(0, 2), "C:") + "\\" + databasename + DateTime.Today.ToString("yyyyMMdd");
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void FastZipCompress(string pathDBBackup, string zipname)
        {
            FastZip fZip = new FastZip();

            fZip.CreateZip("C:\\icalc\\" + zipname + ".zip", pathDBBackup.Replace("xxx.bak", ""), false, ".bak$");
        }

        private void BackupDB(string connectionstring, string dbname, string dbPath)
        {
            Cursor.Current = Cursors.Arrow;
            bool check = false;
            check = Base.backupDatabase3(connectionstring, dbname, dbPath);

            //Copy the file to the C:\drive
            if (check == true)
            {
                //MessageBox.Show("Source = " + dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) + "\\ICALC", "X:") + 
                //                dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak", "Information", MessageBoxButtons.OK);

                Path = dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2), "C:") + dbname +
                       DateTime.Today.ToString("yyyyMMdd") + " \\\\";

                createZipFolder(Path, dbname);

                //MessageBox.Show("dest = " + Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak", "Information", MessageBoxButtons.OK);
                check = BusinessLanguage.copyBackupFile(dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) +
                        "\\ICALC", "Z:") + dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak",
                        Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak");

                if (check == true)
                {
                    string filename = dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak";
                    FastZipCompress(Path + "\\", dbname + DateTime.Today.ToString("yyyyMMdd"));
                    DialogResult checks = MessageBox.Show("Backup Done to : " + Path, "Information", MessageBoxButtons.YesNo);
                    //string pathes = "c:\\icalc\\" + dbname + DateTime.Today.ToString("yyyyMMdd") + ".zip";
                    string pathes = "c:\\icalc\\" + filename.Replace("xxx.bak", "").Trim() + ".zip";
                    if (checks == DialogResult.Yes)
                    {
                       
                    }
                }
                else
                {
                    MessageBox.Show("Copy unsuccessfull from : " + dbPath.Substring(0, 2) + "   Copy unsuccessfull to :" + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Backup unsuccessfull to : " + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
            }

            Cursor.Current = Cursors.Arrow;

        }
                    
        private void defaultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);
        }

        private void btnMetervs_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("MetersVSPayoutStope");
            this.Cursor = Cursors.Arrow;
        }

        private void btnManualSend_Click(object sender, EventArgs e)
        {
            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, "MANUAL");
            paysend.Show();
        }
                    
        private void extractAndCalcShiftsForRefresh(int DayStart, int DayEnd)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int shiftsCheck = 0;

            foreach (DataRow row in BonusShifts.Rows)
            {
                foreach (DataColumn column in BonusShifts.Columns)
                {
                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                        {
                            if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" ||
                                row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                                shiftsCheck = 1;
                            }
                            else
                            {
                                if (row[column].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else { }

                            }
                        }
                        else
                        {
                            row[column] = "*";
                        }
                    }
                    else
                    {
                        if (column.ColumnName == "BONUSTYPE")
                        {
                            row["BONUSTYPE"] = txtBonusType.Text.ToString();
                        }
                    }
                }//foreach datacolumn

                row["SHIFTS_WORKED"] = intShiftsWorked;
                row["AWOP_SHIFTS"] = intAwopShifts;
                intShiftsWorked = 0;
                intAwopShifts = 0;
            }


            if (importdone == 0)//jvdw
            {
                fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
                importdone = 1;
            }

            Application.DoEvents();
        }
                    
        public void updateShifts(DataTable BonusShifts)
        {

            foreach (DataRow row in BonusShifts.Rows)
            {
                IEnumerable<DataRow> query1 = from rec in BonusShifts.AsEnumerable()
                                              where rec.Field<string>("EMPLOYEE_NO").Trim() == row["EMPLOYEE_NO"].ToString().Trim()
                                              where rec.Field<string>("Gang").Trim() == row["GANG"].ToString().Trim()
                                              where rec.Field<string>("WAGECODE").Trim() == row["WAGECODE"].ToString().Trim()
                                              select rec;

                DataTable testTB = query1.CopyToDataTable<DataRow>();

                if (testTB.Rows.Count == 1)
                {
                    string update = "Update Bonusshifts set shifts_worked = '" + (Convert.ToInt32(testTB.Rows[0]["Shifts_Worked"].ToString().Trim()) +
                                                "', Awop_shifts  = '" + testTB.Rows[0]["Awop_Shifts"].ToString().Trim() +
                                                "' where employee_no = '" + testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim() + "' AND Gang = '" + testTB.Rows[0]["GANG"].ToString().Trim() + "'");
                    //Convert.ToInt32(BonusShifts.Rows[0]["Shifts_Worked"].ToString().Trim()
                    TB.InsertData(Base.DBConnectionString, update);
                }
                else
                {

                }
            }
        }

        public void insertEmployee(DataTable BonusShifts)
        {
            DataTable newMembers = new DataTable();
            DataTable bonusShiftCurrent = new DataTable();
            bonusShiftCurrent = TB.createDataTableWithAdapter(Base.DBConnectionString, "select * from bonusshifts where section = '" + txtSelectedSection.Text.ToString().Trim() + "'");
            DataTable bonusShiftCurrent2 = new DataTable();
            bonusShiftCurrent2 = TB.createDataTableWithAdapter(Base.DBConnectionString, "select * from bonusshifts where section = '" + txtSelectedSection.Text.ToString().Trim() + "'");

            foreach (DataRow row in BonusShifts.Rows)
            {
                IEnumerable<DataRow> query1 = from rec in BonusShifts.AsEnumerable()
                                              where rec.Field<string>("EMPLOYEE_NO").Trim() == row["EMPLOYEE_NO"].ToString().Trim()
                                              where rec.Field<string>("Gang").Trim() == row["GANG"].ToString().Trim()
                                              where rec.Field<string>("WAGECODE").Trim() == row["WAGECODE"].ToString().Trim()
                                              where rec.Field<string>("SECTION").Trim() == txtSelectedSection.Text.ToString().Trim()
                                              select rec;

                DataTable testTB = query1.CopyToDataTable<DataRow>();

                bool alreadyOn = false;
                string TEST = testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim();



                foreach (DataRow current in bonusShiftCurrent.Rows)
                {
                    string TEST2 = current["EMPLOYEE_NO"].ToString().Trim();
                    if (current["EMPLOYEE_NO"].ToString().Trim() == testTB.Rows[0]["EMPLOYEE_NO"].ToString().Trim() && current["GANG"].ToString().Trim() == testTB.Rows[0]["GANG"].ToString().Trim() && current["WAGECODE"].ToString().Trim() == testTB.Rows[0]["WAGECODE"].ToString().Trim())
                    {
                        alreadyOn = true;
                    }
                }

                if (alreadyOn == false)
                {
                    foreach (DataRow newMem in testTB.Rows)
                    {
                        DataRow fff = newMem;
                        bonusShiftCurrent2.Rows.Add(fff.ItemArray);
                    }

                    bonusShiftCurrent2.AcceptChanges();

                }
            }

            TB.saveCalculations2(bonusShiftCurrent2, Base.DBConnectionString, "where section ='" + txtSelectedSection.Text.ToString().Trim() + "'", "BONUSSHIFTS");
        }

        private void btnSectionSelect_Click(object sender, EventArgs e)
        {
            listBox2.SelectedIndex = 0;
            listBox2.Select();
            listBox2.Focus();
            listBox2.BackColor = Color.LightBlue;
        }

        private void brtnSetCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            btnBaseCalcsHeader_Click("me", e);
        }                 

        private void best2Worst_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.AppClass mm = new MetaReportRuntime.AppClass();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STP_BestToWorst");
            this.Cursor = Cursors.Arrow;

        }

        private void worst2Best_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.AppClass mm = new MetaReportRuntime.AppClass();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("STP_WorstToBest");
            this.Cursor = Cursors.Arrow;
        }

        private void MISAudits_Click(object sender, EventArgs e)
        {
            scrAudit Audit = new scrAudit();
            Audit.AuditLoad(Base.DBConnectionString, Base.BaseConnectionString, BusinessLanguage, TB, txtSelectedSection.Text.Trim());
            Audit.Show();
        }

        private void grdFactors_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                txtVarValue.Text = grdFactors["VarValue", e.RowIndex].Value.ToString().Trim();
                cboVarName.Text = grdFactors["VarName", e.RowIndex].Value.ToString().Trim();

                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = false;
                btnInsertRow.Enabled = false;
            }

            #region Trigger output
            //load the CURRENT values into dictionaries before the update 
            dictPrimaryKeyValues.Clear();
            dictGridValues.Clear();

            foreach (string s in lstPrimaryKeyColumns)
            {
                dictPrimaryKeyValues.Add(s, grdFactors[s, e.RowIndex].Value.ToString().Trim());
            }

            foreach (string s in lstTableColumns)
            {
                dictGridValues.Add(s, grdFactors[s, e.RowIndex].Value.ToString().Trim());
            }
            #endregion

            Cursor.Current = Cursors.Arrow;
        }

        private void cboColumnNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(newDataTable, cboColumnNames.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboColumnValues.Items.Add(s.Trim());
            }
        }

        private void cboColumnShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }
        }

        private void txtADTeamShifts_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPayShifts_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtAwops_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtMinersSafetyInd_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtSurname_TextChanged(object sender, EventArgs e)
        {

        }

        private void cboColumnValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            //unhide all the columns currently hidden.
            for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
            {
                grdActiveSheet.Columns[i].Visible = true;
            }

            //Extract the formulas
            AConn = Analysis.AnalysisConnection;
            AConn.Open();

            //color the calculations blue
            if (TB.TBName.Trim().Contains(BusinessLanguage.Period.Trim()))
            {
                DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(),
                                                                       TB.TBName.Substring(0, TB.TBName.Trim().IndexOf(BusinessLanguage.Period.Trim())),
                                                                       Base.AnalysisConnectionString);
                if (tempDataTable.Rows.Count > 0)
                {
                    foreach (DataRow dt in tempDataTable.Rows)
                    {
                        string strValue = dt["Calc_Name"].ToString().Trim();
                        int intValue = grdActiveSheet.Columns.Count - 1;

                        for (int i = intValue; i >= 3; --i)
                        {
                            string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                            if (strValue == strHeader)
                            {
                                for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                                {
                                    grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                                }
                            }
                        }
                    }
                }
            }
            else
            {

            }


            //Set boolean value to false to show that the listbox contains column names.
            blTablenames = false;

            listBox1.Items.Clear();
            listBox1.SelectionMode = SelectionMode.MultiSimple;
            foreach (string s in TB.ListOfSelectedTableColumns)
            {
                listBox1.Items.Add(s.Trim());
            }
        }

        private void grdParticipants_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if(grdParticipants.Columns.Count > 0 && grdParticipants.Columns.Count > 4)
                {
                    txtAutoDGang.Text = grdParticipants["GANG", e.RowIndex].Value.ToString().Trim();
                    txtAutoEmployee.Text = grdParticipants["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                    cboParticipantsCrew.Text = grdParticipants["CREW", e.RowIndex].Value.ToString().Trim();
                    txtAutoEmployee.Text = grdParticipants["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim(); 
                    txtParticipatantsAwopShifts.Text = grdParticipants["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                    txtParticipatantsShiftsWorked.Text = grdParticipants["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                    cboParticipantsCrewType.Text = grdParticipants["CREWtype", e.RowIndex].Value.ToString().Trim();

                    #region Trigger output
                    //load the CURRENT values into dictionaries before the update 
                    dictPrimaryKeyValues.Clear();
                    dictGridValues.Clear();

                    foreach (string s in lstPrimaryKeyColumns)
                    {
                        if (e.RowIndex < 0)
                        {
                        }
                        else
                        {
                            dictPrimaryKeyValues.Add(s, grdParticipants[s, e.RowIndex].Value.ToString().Trim());
                        }
                    }

                    foreach (string s in lstTableColumns)
                    {
                        if (e.RowIndex < 0)
                        {
                        }
                        else
                        {
                            dictGridValues.Add(s, grdParticipants[s, e.RowIndex].Value.ToString().Trim());
                        }
                    }
                    #endregion
                }
                else
                {
                    btnShowAll.Visible = true;
                    btnShowEmpl.Visible = true;
                    txtAutoDGang.Text = string.Empty;
                    txtAutoEmployee.Text = string.Empty;
                    cboParticipantsCrew.Text = string.Empty;
                    txtAutoEmployee.Text = string.Empty;
                    txtParticipatantsShiftsWorked.Text = string.Empty;
                    txtAutoEmployee.Text = grdParticipants["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                }
             
             }

           
        }

        private void grdParticipants_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdParticipants);
            }
        }

        private void grdGangNames_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdCrews);
            }
        }

        private void grdRates_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdRates);
            }
        }

        private void grdOffDays_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdOffDays);
            }
        }

        private void grdMineParameters_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdMineParameters);
            }
        }

        private List<string> extractDayShiftGangs(string Gang)
        {
            DataTable temp = new DataTable();
            List<string> lstTemp = new List<string>();


            if (SupportLink.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in SupportLink.AsEnumerable()
                                              where locks.Field<string>("GANG").TrimEnd() == Gang
                                              select locks;

                try
                {
                    temp = query1.CopyToDataTable<DataRow>();
                }
                catch
                {

                }
            }

            if (temp.Rows.Count > 0)
            {
                lstTemp = TB.loadDistinctValuesFromColumn(temp, "DAYGANG");
            }
            else
            {

            }

            return lstTemp;
        }

        private void createParticipants(string CrewName)
        {
            this.Cursor = Cursors.WaitCursor;

            if (Participants.Rows.Count > 0)
            {

                //Extract a new participants table from BONUSSHIFTS

                //The crew and crewttype will be replaced to the original crew names as from ADTEAM
                DataTable tempGangs = new DataTable();
                IEnumerable<DataRow> query1 = from locks in Crews.AsEnumerable()
                                              where locks.Field<string>("CREW").TrimEnd() == CrewName
                                              where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                              where locks.Field<string>("PERIOD").TrimEnd() == BusinessLanguage.Period.Trim()
                                              select locks;

                try
                {
                    tempGangs = query1.CopyToDataTable<DataRow>();

                    if (tempGangs.Rows.Count > 0)
                    {
                        //get the list of gangs to import and refresh from bonusshifts
                        lstNames = TB.loadDistinctValuesFromColumn(tempGangs, "GANG");
                        string strListOfGangs = string.Empty;
                        if (lstNames.Count > 0)
                        {
                            strListOfGangs = "Where GANG in ('";
                            foreach (string s in lstNames)
                            {
                                strListOfGangs = strListOfGangs.Trim() + s.Trim() + "','";

                            }

                            strListOfGangs = strListOfGangs.ToString().Trim().Substring(0, strListOfGangs.ToString().Trim().Length - 2) + ")";
                        }

                        DataTable temp = TB.extractParticipantsForBacSer8000(Base.DBConnectionString, txtSelectedSection.Text.Trim(), BusinessLanguage.BussUnit,
                                                                 BusinessLanguage.Period, "", strListOfGangs);

                        TB.saveCalculations2(temp, Base.DBConnectionString, strListOfGangs, "PARTICIPANTS");
                        evaluateParticipants();

                        MessageBox.Show("Participants were updated.", "Information", MessageBoxButtons.OK);



                    }
                    else
                    {
                        MessageBox.Show("Import not possible, because no gangs on Bonusshifts.  Please re-import bonusshifts.",
                                         "Message", MessageBoxButtons.OK);
                    }
                }
                catch { }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void cboParticipantsFilterCREWS_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboParticipantsFilterGANGS.Text = string.Empty;
            DataTable temp = new DataTable();
            evaluateParticipants();
            IEnumerable<DataRow> query1 = from locks in Participants.AsEnumerable()
                                          where locks.Field<string>("CREW").TrimEnd() == cboParticipantsFilterCREWS.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
            }
            catch
            {

            }

            grdParticipants.DataSource = temp;
            hideColumnsOfGrid("grdParticipants");

        }

        private void cboParticipantsFilterGANGS_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboParticipantsFilterCREWS.Text = string.Empty;
            DataTable temp = new DataTable();
            evaluateParticipants();
            IEnumerable<DataRow> query1 = from locks in Participants.AsEnumerable()
                                          where locks.Field<string>("Gang").TrimEnd() == cboParticipantsFilterGANGS.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
            }
            catch
            {

            }

            grdParticipants.DataSource = temp;
        }

        private void btnShowEmpl_Click(object sender, EventArgs e)
        {

            DataTable temp = new DataTable();
            IEnumerable<DataRow> query1 = from locks in Participants.AsEnumerable()
                                          where locks.Field<string>("Employee_no").TrimEnd() == txtAutoEmployee.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
            }
            catch
            {
            }

            grdParticipants.DataSource = temp;
        }

        

        private void grdMineParameters_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                
                txtTons_Actual.Text = grdMineParameters["TONS_Actual", e.RowIndex].Value.ToString().Trim();
                


                btnUpdate.Enabled = true;
                btnInsertRow.Enabled = false;
                btnDeleteRow.Enabled = false;

            }
            Cursor.Current = Cursors.Arrow;
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void btnResetListBos_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            extractDBTableNames(listBox1);

            this.Cursor = Cursors.Arrow;
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                //unhide first all the columns.
                for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
                {
                    grdActiveSheet.Columns[i].Visible = true;
                }

                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void btnOffdays_Click(object sender, EventArgs e)
        {
            openTab(tabOffdays);
        }

        private void btnPrintAttendanceShort_Click(object sender, EventArgs e)
        {
            btnAttendance_Click_1("Method", null);
        }

        private void btnUpdateCrews_Click(object sender, EventArgs e)
        {
            openTab(tabCrews);
        }

        private void btnUpdateParticipants_Click(object sender, EventArgs e)
        {
            openTab(tabParticipants);
        }

        private void btnUpdateMineparameters_Click(object sender, EventArgs e)
        {
            openTab(tabMineParameters);
        }

        private void cboCrewsFilterCREWS_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable temp = new DataTable();
            evaluateCrews();
            IEnumerable<DataRow> query1 = from locks in Crews.AsEnumerable()
                                          where locks.Field<string>("CREW").TrimEnd() == cboCrewsFilterCREWS.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
            }
            catch
            {

            }

            grdCrews.DataSource = temp;
            hideColumnsOfGrid("grdCrews");
        }

        private void cboCrewType_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboCrew.Items.Clear();
            cboCrew.Text = string.Empty;

            DataTable temp = new DataTable();

            IEnumerable<DataRow> query1 = from locks in Crews.AsEnumerable()
                                          where locks.Field<string>("CrewType").TrimEnd() == cboCrewType.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
                lstNames = TB.loadDistinctValuesFromColumn(temp, "CREW");  

                foreach (string s in lstNames)
                {
                    cboCrew.Items.Add(s.Trim()); 
                }
            }
            catch
            {

            }
        }

        private void cboParticipantsCrewType_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboParticipantsCrew.Items.Clear();
            cboParticipantsCrew.Text = string.Empty;

            DataTable temp = new DataTable();

            IEnumerable<DataRow> query1 = from locks in Crews.AsEnumerable()
                                          where locks.Field<string>("CrewType").TrimEnd() == cboParticipantsCrewType.Text.Trim()
                                          select locks;

            try
            {
                temp = query1.CopyToDataTable<DataRow>();
                lstNames = TB.loadDistinctValuesFromColumn(temp, "CREW");

                foreach (string s in lstNames)
                {
                    cboParticipantsCrew.Items.Add(s.Trim());
                }
            }
            catch
            {

            }
        }

        private void label68_Click(object sender, EventArgs e)
        {
            lstOffDayValue.SelectedItems.Clear();
        }

        private void btnImportParticipants_Click(object sender, EventArgs e)
        {
            createParticipants();
        }

        private void createParticipantsPerGang(string GangSelected)
        {
            this.Cursor = Cursors.WaitCursor;

            DataTable temp = TB.extractParticipantsForBacSer8000(Base.DBConnectionString, txtSelectedSection.Text.Trim(), BusinessLanguage.BussUnit,
                                                                 BusinessLanguage.Period, "",
                                                                 " where gang in ('" + GangSelected + "')");

            string strDelete = " where gang = '" + GangSelected + "' and period = '" + BusinessLanguage.Period.Trim() + "'";

            TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "PARTICIPANTS");
            evaluateParticipants();

            MessageBox.Show("Participants were updated.", "Information", MessageBoxButtons.OK);

            this.Cursor = Cursors.Arrow;
        }

        private void btnRefreshCrew_Click_1(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("The employees of the gang: " + cboParticipantsFilterGANGS.Text.Trim() + " will be refreshed from clockedshifts." +
                            Environment.NewLine + "Are you sure?", "Question", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    createParticipantsPerGang(cboParticipantsFilterGANGS.Text.Trim());
                    break;


                case DialogResult.No:
                    break;


            }
        }

        private void btnShowAll_Click(object sender, EventArgs e)
        {

        }

       

        }

    }




                   
