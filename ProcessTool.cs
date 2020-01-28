using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using DevExpress.XtraTab;
using DevExpress.XtraTab.Registrator;
using System.Threading;
using DevExpress.Spreadsheet;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraBars.Ribbon.ViewInfo;
using DevExpress.Utils;
using System.Data.OracleClient;
using DevExpress.Docs;
using DevExpress.Spreadsheet.Export;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using DevExpress.XtraEditors;
using MDLInputFileSetting;
using DevExpress.XtraReports.Parameters;
using DevExpress.Utils.CommonDialogs;
using DevExpress.XtraSpreadsheet;


namespace MDL
{
    public partial class ProcessTool : Form
    {
        #region declareVariables
        MessageBoxButtons buttons = MessageBoxButtons.OK;
        BusinessLayer blyBL = null;

        //Sudheer BGRF-2056
        public static DataSet ds = new DataSet();
        public static DataTable dtbActiveCategory = new DataTable("O_ACTIVE_CATEGORY");
        public static DataTable dtbPreSetChar = new DataTable("O_PRE_SET_CHAR");
        public static DataTable dtbProcGrpSet = new DataTable("O_PROC_GRP_SET");
        public static DataTable dtbExtCodeGrp = new DataTable("O_EXT_CODE_GRP");
        public static DataTable dtbActiveRetailer = new DataTable("O_ACTIVE_RETAILER");
        public static DataTable dtbDicLvlFlagSetting = new DataTable("O_DIC_LVL_FLAG_SETTING");
        public static DataTable dtbPCAsgnlvlFlagSetting = new DataTable("O_PC_ASGN_LVL_FLAG_SETTING");
        public static DataTable dtbModAsgnlvlFlagSetting = new DataTable("O_MOD_ASGN_LVL_FLAG_SETTING");
        public static DataTable dtbHeadingType = new DataTable("O_HEADING_TYPE");
        public static DataTable dtbHeadingPCSetting = new DataTable("O_HEADING_PC_SETTING");
        public static DataTable dtbHeadingModSetting = new DataTable("O_HEADING_MOD_SETTING");
        public static DataTable dtbRetailerDeptSupp = new DataTable("O_RETAILER_DEPT_SUPP");
        public static DataTable dtbUomMappingList = new DataTable("O_UOM_MAPPING_LIST");

        public static DataSet dsTemp = new DataSet();
        public static DataTable dtbTempActiveCategory = new DataTable("O_ACTIVE_CATEGORY");
        public static DataTable dtbTempPreSetChar = new DataTable("O_PRE_SET_CHAR");
        public static DataTable dtbTempProcGrpSet = new DataTable("O_PROC_GRP_SET");
        public static DataTable dtbTempExtCodeGrp = new DataTable("O_EXT_CODE_GRP");
        public static DataTable dtbTempActiveRetailer = new DataTable("O_ACTIVE_RETAILER");
        public static DataTable dtbTempDicLvlFlagSetting = new DataTable("O_DIC_LVL_FLAG_SETTING");
        public static DataTable dtbTempPCAsgnlvlFlagSetting = new DataTable("O_PC_ASGN_LVL_FLAG_SETTING");
        public static DataTable dtbTempModAsgnlvlFlagSetting = new DataTable("O_MOD_ASGN_LVL_FLAG_SETTING");
        public static DataTable dtbTempHeadingType = new DataTable("O_HEADING_TYPE");
        public static DataTable dtbTempHeadingPCSetting = new DataTable("O_HEADING_PC_SETTING");
        public static DataTable dtbTempHeadingModSetting = new DataTable("O_HEADING_MOD_SETTING");
        public static DataTable dtbTempRetailerDeptSupp = new DataTable("O_RETAILER_DEPT_SUPP");
        public static DataTable dtbTempUomMappingList = new DataTable("O_UOM_MAPPING_LIST");
        public static string errorRecord = string.Empty; //BGRF-2051
        public static bool cellDataValidation = false;//BGRF-2086
        public static bool errorCell = false;//BGRF-2086
        #endregion

        //Sudheer BGRF-1967
        public ProcessTool()
        {
            try
            {
                InitializeComponent();
                //  PaintStyleCollection.DefaultPaintStyles.Add(new MyRegistrator());
                //Read output file path from Config file.
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "MDL-Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Window Maximize
        //Form load event 
        private void ProcessTool_Load(object sender, EventArgs e)
        {
            Left = Top = 0;
            Width = Screen.PrimaryScreen.WorkingArea.Width;
            Height = Screen.PrimaryScreen.WorkingArea.Height;
            this.Activated += AfterLoading;
        }

        //Sudheer BGRF-2087
        private void AfterLoading(object sender, EventArgs e)
        {
            this.Activated -= AfterLoading;
            LoadSpreadSheet();
            ColumnWidth();
        }

        #endregion

        //Sudheer BGRF-2056
        //BGRF-2086 edited this method for cell value lenth and not allowed null validation
        public void CreateDt(int tableNumber)
        {
            try
            {
                int colCount = ds.Tables[tableNumber].Columns.Count;
                int totalRows = 1048576;
                Worksheet ws = spreadsheetControl1.Document.Worksheets[tableNumber];
                CellRange r = ws.Range.Parse("A1:" + ExcelColumnIndexToName(colCount - 1) + totalRows);

                //start BGRF-2086
                DataTable table = new DataTable();
                table = ds.Tables[tableNumber].Clone();
                //end BGRF-2086

                DataTableExporter exporter = ws.CreateDataTableExporter(r, table, true);
                exporter.Export();

                // start BGRF-2086
                // to trim all the dt column
                foreach (DataRow dr in table.Rows) // trim string data
                {
                    foreach (DataColumn dc in table.Columns)
                    {
                        if (dc.DataType == typeof(string))
                        {
                            object o = dr[dc];
                            if (!Convert.IsDBNull(o) && o != null)
                            {
                                dr[dc] = o.ToString().Trim();
                            }
                        }
                    }
                }
 
                Color colReset = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                IWorkbook workbook = spreadsheetControl1.Document;

                var columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;
                var rowIndex = spreadsheetControl1.ActiveCell.RowIndex;

                //end BGRF-2086

                #region O_ACTIVE_CATEGORY
                if (ds.Tables[tableNumber].TableName == "O_ACTIVE_CATEGORY")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool categoryValue = false;
                    bool service = false;
                    bool type = false;

                    bool CATEGORYATTRVALUELen = false;
                    bool SERVICELen = false;
                    bool TYPELen = false;

                    var worksheet = workbook.Worksheets["O_ACTIVE_CATEGORY"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["SERVICE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString()))//&& !categoryValue)
                            {
                                if (!categoryValue)
                                {
                                    error = true;
                                    categoryValue = true;
                                    errorMsg += "CATEGORYATTRVALUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 20)
                                {
                                    if (!CATEGORYATTRVALUELen)
                                    {
                                        error = true;
                                        CATEGORYATTRVALUELen = true;
                                        errorMsg += "CATEGORYATTRVALUE length is greater than Max lengh 20\n";
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["SERVICE"].ToString()))// && !service)
                            {
                                if (!service)
                                {
                                    error = true;
                                    service = true;
                                    errorMsg += "SERVICE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SERVICE"].ToString().Length > 20)
                                {
                                    if (!SERVICELen)
                                    {
                                        error = true;
                                        SERVICELen = true;
                                        errorMsg += "SERVICE length is greater than Max lengh 20\n";
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SERVICE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()))// && !type)
                            {
                                if (!type)
                                {
                                    error = true;
                                    type = true;
                                    errorMsg += "TYPE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["TYPE"].ToString().Length > 10)
                                {
                                    if (!TYPELen)
                                    {
                                        error = true;
                                        TYPELen = true;
                                        errorMsg += "TYPE length is greater than Max lengh 10\n";
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["TYPE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }
                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_ACTIVE_CATEGORY\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempActiveCategory.Clear();
                    dtbTempActiveCategory.Merge(table);
                }
                #endregion
                #region O_PRE_SET_CHAR
                if (ds.Tables[tableNumber].TableName == "O_PRE_SET_CHAR")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool type = false;
                    bool attrNo = false;
                    bool required = false;

                    bool typeLen = false;
                    bool attrNoLen = false;
                    bool requiredLen = false;

                    var worksheet = workbook.Worksheets["O_PRE_SET_CHAR"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["REQUIRED"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()))
                            {
                                if (!type)
                                {
                                    error = true;
                                    type = true;
                                    errorMsg += "TYPE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["TYPE"].ToString().Length > 30)
                                {
                                    if (!typeLen)
                                    {
                                        error = true;
                                        typeLen = true;
                                        errorMsg += "TYPE length is greater than Max lengh 30\n";
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["TYPE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()))
                            {
                                if (!attrNo)
                                {
                                    error = true;
                                    attrNo = true;
                                    errorMsg += "ATTRNO can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ATTRNO"].ToString().Length > 15)
                                {
                                    if (!attrNoLen)
                                    {
                                        error = true;
                                        attrNoLen = true;
                                        errorMsg += "ATTRNO length is greater than Max lengh 15\n";
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ATTRNO"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["REQUIRED"].ToString()))
                            {
                                if (!required)
                                {
                                    error = true;
                                    required = true;
                                    errorMsg += "REQUIRED can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["REQUIRED"].ToString().Length > 1)
                                {
                                    if (!requiredLen)
                                    {
                                        error = true;
                                        requiredLen = true;
                                        errorMsg += "REQUIRED length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["REQUIRED"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_PRE_SET_CHAR\n";
                        //errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempPreSetChar.Clear();
                    dtbTempPreSetChar.Merge(table);
                }
                #endregion
                #region O_PROC_GRP_SET
                if (ds.Tables[tableNumber].TableName == "O_PROC_GRP_SET")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool PROC_GROUP_SET_ID = false;
                    bool PROC_GROUP_SET_NAME = false;
                    bool FOLLOW_EXT_RULE = false;
                    bool EXCP_SENT_TO_XCD_BRWSR = false;
                    bool EXCP_SENT_TO_UNCDBLE_ITM = false;

                    bool PROC_GROUP_SET_IDLen = false;
                    bool PROC_GROUP_SET_NAMELen = false;
                    bool FOLLOW_EXT_RULELen = false;
                    bool EXCP_SENT_TO_XCD_BRWSRLen = false;
                    bool EXCP_SENT_TO_UNCDBLE_ITMLen = false;
                    bool EXCP_SEN_TO_SUPER_GTP_ITMLen = false;

                    var worksheet = workbook.Worksheets["O_PROC_GRP_SET"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_ID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_NAME"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["FOLLOW_EXT_RULE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["EXCP_SENT_TO_XCD_BRWSR"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["EXCP_SENT_TO_UNCDBLE_ITM"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count - 1; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString()))
                            {
                                if (table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString().Length > 1)
                                {
                                    if (!EXCP_SEN_TO_SUPER_GTP_ITMLen)
                                    {
                                        error = true;
                                        EXCP_SEN_TO_SUPER_GTP_ITMLen = true;
                                        errorMsg += "EXCP_SEN_TO_SUPER_GTP_ITM length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_ID"].ToString()))
                            {
                                if (!PROC_GROUP_SET_ID)
                                {
                                    error = true;
                                    PROC_GROUP_SET_ID = true;
                                    errorMsg += "PROC_GROUP_SET_ID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["PROC_GROUP_SET_ID"].ToString().Length > 20)
                                {
                                    if (!PROC_GROUP_SET_IDLen)
                                    {
                                        error = true;
                                        PROC_GROUP_SET_IDLen = true;
                                        errorMsg += "PROC_GROUP_SET_ID length is greater than Max lengh 20\n";
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["PROC_GROUP_SET_ID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_NAME"].ToString()))
                            {
                                if (!PROC_GROUP_SET_NAME)
                                {
                                    error = true;
                                    PROC_GROUP_SET_NAME = true;
                                    errorMsg += "PROC_GROUP_SET_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["PROC_GROUP_SET_NAME"].ToString().Length > 20)
                                {
                                    if (!PROC_GROUP_SET_NAMELen)
                                    {
                                        error = true;
                                        PROC_GROUP_SET_NAMELen = true;
                                        errorMsg += "PROC_GROUP_SET_NAME length is greater than Max lengh 20\n";
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["PROC_GROUP_SET_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["FOLLOW_EXT_RULE"].ToString()))
                            {
                                if (!FOLLOW_EXT_RULE)
                                {
                                    error = true;
                                    FOLLOW_EXT_RULE = true;
                                    errorMsg += "FOLLOW_EXT_RULE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FOLLOW_EXT_RULE"].ToString().Length > 1)
                                {
                                    if (!FOLLOW_EXT_RULELen)
                                    {
                                        error = true;
                                        FOLLOW_EXT_RULELen = true;
                                        errorMsg += "FOLLOW_EXT_RULE length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["FOLLOW_EXT_RULE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["EXCP_SENT_TO_XCD_BRWSR"].ToString()))
                            {
                                if (!EXCP_SENT_TO_XCD_BRWSR)
                                {
                                    error = true;
                                    EXCP_SENT_TO_XCD_BRWSR = true;
                                    errorMsg += "EXCP_SENT_TO_XCD_BRWSR can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["EXCP_SENT_TO_XCD_BRWSR"].ToString().Length > 1)
                                {
                                    if (!EXCP_SENT_TO_XCD_BRWSRLen)
                                    {
                                        error = true;
                                        EXCP_SENT_TO_XCD_BRWSRLen = true;
                                        errorMsg += "EXCP_SENT_TO_XCD_BRWSR length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EXCP_SENT_TO_XCD_BRWSR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["EXCP_SENT_TO_UNCDBLE_ITM"].ToString()))
                            {
                                if (!EXCP_SENT_TO_UNCDBLE_ITM)
                                {
                                    error = true;
                                    EXCP_SENT_TO_UNCDBLE_ITM = true;
                                    errorMsg += "EXCP_SENT_TO_UNCDBLE_ITM can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["EXCP_SENT_TO_UNCDBLE_ITM"].ToString().Length > 1)
                                {
                                    if (!EXCP_SENT_TO_UNCDBLE_ITMLen)
                                    {
                                        error = true;
                                        EXCP_SENT_TO_UNCDBLE_ITMLen = true;
                                        errorMsg += "EXCP_SENT_TO_UNCDBLE_ITM length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EXCP_SENT_TO_UNCDBLE_ITM"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString()))
                            {
                                if (table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString().Length > 1)
                                {
                                    if (!EXCP_SEN_TO_SUPER_GTP_ITMLen)
                                    {
                                        error = true;
                                        EXCP_SEN_TO_SUPER_GTP_ITMLen = true;
                                        errorMsg += "EXCP_SEN_TO_SUPER_GTP_ITM length is greater than Max lengh 1\n";
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EXCP_SEN_TO_SUPER_GTP_ITM"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }
                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_PROC_GRP_SET\n";
                        //errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempProcGrpSet.Clear();
                    dtbTempProcGrpSet.Merge(table);
                }
                #endregion
                #region O_EXT_CODE_GRP
                if (ds.Tables[tableNumber].TableName == "O_EXT_CODE_GRP")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool SHORTNAME = false;
                    bool EXTERNAL_CODE_GROUP_NAME = false;
                    bool SHORTNAMELen = false;
                    bool EXTERNAL_CODE_GROUP_NAMELen = false;

                    var worksheet = workbook.Worksheets["O_EXT_CODE_GRP"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["SHORTNAME"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["EXTERNAL_CODE_GROUP_NAME"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["SHORTNAME"].ToString()))
                            {
                                if (!SHORTNAME)
                                {
                                    error = true;
                                    SHORTNAME = true;
                                    errorMsg += "SHORTNAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SHORTNAME"].ToString().Length > 10)
                                {
                                    if (!SHORTNAMELen)
                                    {
                                        error = true;
                                        SHORTNAMELen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SHORTNAME"].ToString().Length > 10)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["EXTERNAL_CODE_GROUP_NAME"].ToString()))
                            {
                                if (!EXTERNAL_CODE_GROUP_NAME)
                                {
                                    error = true;
                                    EXTERNAL_CODE_GROUP_NAME = true;
                                    errorMsg += "EXTERNAL_CODE_GROUP_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["EXTERNAL_CODE_GROUP_NAME"].ToString().Length > 30)
                                {
                                    if (!EXTERNAL_CODE_GROUP_NAMELen)
                                    {
                                        error = true;
                                        EXTERNAL_CODE_GROUP_NAMELen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EXTERNAL_CODE_GROUP_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_EXT_CODE_GRP\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempExtCodeGrp.Clear();
                    dtbTempExtCodeGrp.Merge(table);
                }
                #endregion
                #region O_ACTIVE_RETAILER
                if (ds.Tables[tableNumber].TableName == "O_ACTIVE_RETAILER")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool SOURCEID = false;
                    bool PROC_GROUP_SET_NAME = false;
                    bool TYPE = false;
                    bool EAN = false;
                    bool UPC = false;
                    bool LAC = false;
                    bool CIP = false;

                    bool SOURCEIDLen = false;
                    bool PROC_GROUP_SET_NAMELen = false;
                    bool TYPELen = false;
                    bool EANLen = false;
                    bool UPCLen = false;
                    bool LACLen = false;
                    bool CIPLen = false;


                    var worksheet = workbook.Worksheets["O_ACTIVE_RETAILER"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["SOURCEID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_NAME"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["EAN"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["UPC"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["LAC"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["CIP"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["SOURCEID"].ToString()))
                            {
                                if (!SOURCEID)
                                {
                                    error = true;
                                    SOURCEID = true;
                                    errorMsg += "SOURCEID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SOURCEID"].ToString().Length > 10)
                                {
                                    if (!SOURCEIDLen)
                                    {
                                        error = true;
                                        SOURCEIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SOURCEID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["PROC_GROUP_SET_NAME"].ToString()))
                            {
                                if (!PROC_GROUP_SET_NAME)
                                {
                                    error = true;
                                    PROC_GROUP_SET_NAME = true;
                                    errorMsg += "PROC_GROUP_SET_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["PROC_GROUP_SET_NAME"].ToString().Length > 20)
                                {
                                    if (!PROC_GROUP_SET_NAMELen)
                                    {
                                        error = true;
                                        PROC_GROUP_SET_NAMELen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["PROC_GROUP_SET_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()))
                            {
                                if (!TYPE)
                                {
                                    error = true;
                                    TYPE = true;
                                    errorMsg += "TYPE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["TYPE"].ToString().Length > 30)
                                {
                                    if (!TYPELen)
                                    {
                                        error = true;
                                        TYPELen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["TYPE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["EAN"].ToString()))
                            {
                                if (!EAN)
                                {
                                    error = true;
                                    EAN = true;
                                    errorMsg += "EAN can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["EAN"].ToString().Length > 1)
                                {
                                    if (!EANLen)
                                    {
                                        error = true;
                                        EANLen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["EAN"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["UPC"].ToString()))
                            {
                                if (!UPC)
                                {
                                    error = true;
                                    UPC = true;
                                    errorMsg += "UPC can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["UPC"].ToString().Length > 1)
                                {
                                    if (!UPCLen)
                                    {
                                        error = true;
                                        UPCLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["UPC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["LAC"].ToString()))
                            {
                                if (!LAC)
                                {
                                    error = true;
                                    LAC = true;
                                    errorMsg += "LAC can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["LAC"].ToString().Length > 1)
                                {
                                    if (!LACLen)
                                    {
                                        error = true;
                                        LACLen = true;
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["LAC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["CIP"].ToString()))
                            {
                                if (!CIP)
                                {
                                    error = true;
                                    CIP = true;
                                    errorMsg += "CIP can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CIP"].ToString().Length > 30)
                                {
                                    if (!CIPLen)
                                    {
                                        error = true;
                                        CIPLen = true;
                                    }
                                    worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CIP"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_ACTIVE_RETAILER\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempActiveRetailer.Clear();
                    dtbTempActiveRetailer.Merge(table);
                    // dtbTempActiveRetailer = dtbActiveRetailer.Clone();

                }
                #endregion
                #region O_DIC_LVL_FLAG_SETTING
                if (ds.Tables[tableNumber].TableName == "O_DIC_LVL_FLAG_SETTING")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool ATTRNO = false;
                    bool MAP_TO_OGRDS_CHR_VAL = false;
                    bool LONGDESC = false;
                    bool ALTER30MAX = false;
                    bool ALTER5MAX = false;
                    bool CATEGORY_FLAG = false;
                    bool CHAR_TYPE = false;
                    bool NUMERIC_FLAG = false;
                    bool FIXED_ITEM_VALUE = false;
                    bool COPY_ITEM = false;
                    bool MULTI_VALUE = false;
                    bool ABBREVIATE_VALUE = false;
                    bool FIXED_VALUE_LIST = false;
                    bool TRANSLATION_IND = false;
                    bool CHR_VAL_DESCRIPTION_ONLY = false;
                    bool LOCAL = false;
                    bool SORT_ORDER = false;

                    //start BGRF-2086
                    bool ATTRNOLen = false;
                    bool MAP_TO_OGRDS_CHR_VALLen = false;
                    bool LONGDESCLen = false;
                    bool ALTER30MAXLen = false;
                    bool ALTER5MAXLen = false;
                    bool CATEGORY_FLAGLen = false;
                    bool CHAR_TYPELen = false;
                    bool NUMERIC_FLAGLen = false;
                    bool FIXED_ITEM_VALUELen = false;
                    bool COPY_ITEMLen = false;
                    bool MULTI_VALUELen = false;
                    bool ABBREVIATE_VALUELen = false;
                    bool FIXED_VALUE_LISTLen = false;
                    bool TRANSLATION_INDLen = false;
                    bool CHR_VAL_DESCRIPTION_ONLYLen = false;
                    bool LOCALLen = false;
                    bool SORT_ORDERLen = false;
                    bool OGRDSIDLen = false;
                    bool OTHERDESCLen = false;
                    bool ALTER30MAX_OTHERLen = false;
                    bool CUSTOMER_FLAGLen = false;
                    //end BGRF 2086

                    var worksheet = workbook.Worksheets["O_DIC_LVL_FLAG_SETTING"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["MAP_TO_OGRDS_CHR_VAL"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["LONGDESC"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ALTER30MAX"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["ALTER5MAX"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["CATEGORY_FLAG"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["CHAR_TYPE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["NUMERIC_FLAG"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["FIXED_ITEM_VALUE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["COPY_ITEM"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["MULTI_VALUE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ABBREVIATE_VALUE"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["FIXED_VALUE_LIST"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["TRANSLATION_IND"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["CHR_VAL_DESCRIPTION_ONLY"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["LOCAL"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["SORT_ORDER"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["OGRDSID"].ToString()))
                            {
                                if (table.Rows[i]["OGRDSID"].ToString().Length > 10)
                                {
                                    if (!OGRDSIDLen)
                                    {
                                        error = true;
                                        OGRDSIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OGRDSID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["OTHERDESC"].ToString()))
                            {
                                if (table.Rows[i]["OTHERDESC"].ToString().Length > 70)
                                {
                                    if (!OTHERDESCLen)
                                    {
                                        error = true;
                                        OTHERDESCLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OTHERDESC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["ALTER30MAX_OTHER"].ToString()))
                            {
                                if (table.Rows[i]["ALTER30MAX_OTHER"].ToString().Length > 70)
                                {
                                    if (!ALTER30MAX_OTHERLen)
                                    {
                                        error = true;
                                        ALTER30MAX_OTHERLen = true;
                                    }
                                    worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ALTER30MAX_OTHER"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["CUSTOMER_FLAG"].ToString()))
                            {
                                if (table.Rows[i]["CUSTOMER_FLAG"].ToString().Length > 1)
                                {
                                    if (!CUSTOMER_FLAGLen)
                                    {
                                        error = true;
                                        CUSTOMER_FLAGLen = true;
                                    }
                                    worksheet.Cells[i + 1, 11].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CUSTOMER_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 11].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()))
                            {
                                if (!ATTRNO)
                                {
                                    error = true;
                                    ATTRNO = true;
                                    errorMsg += "ATTRNO can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ATTRNO"].ToString().Length > 10)
                                {
                                    if (!ATTRNOLen)
                                    {
                                        error = true;
                                        ATTRNOLen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ATTRNO"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["MAP_TO_OGRDS_CHR_VAL"].ToString()))
                            {
                                if (!MAP_TO_OGRDS_CHR_VAL)
                                {
                                    error = true;
                                    MAP_TO_OGRDS_CHR_VAL = true;
                                    errorMsg += "MAP_TO_OGRDS_CHR_VAL can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MAP_TO_OGRDS_CHR_VAL"].ToString().Length > 1)
                                {
                                    if (!MAP_TO_OGRDS_CHR_VALLen)
                                    {
                                        error = true;
                                        MAP_TO_OGRDS_CHR_VALLen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["MAP_TO_OGRDS_CHR_VAL"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["LONGDESC"].ToString()))
                            {
                                if (!LONGDESC)
                                {
                                    error = true;
                                    LONGDESC = true;
                                    errorMsg += "LONGDESC can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["LONGDESC"].ToString().Length > 70)
                                {
                                    if (!LONGDESCLen)
                                    {
                                        error = true;
                                        LONGDESCLen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["LONGDESC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["ALTER30MAX"].ToString()))
                            {
                                if (!ALTER30MAX)
                                {
                                    error = true;
                                    ALTER30MAX = true;
                                    errorMsg += "ALTER30MAX can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ALTER30MAX"].ToString().Length > 70)
                                {
                                    if (!ALTER30MAXLen)
                                    {
                                        error = true;
                                        ALTER30MAXLen = true;
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ALTER30MAX"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }


                            if (string.IsNullOrEmpty(table.Rows[i]["ALTER5MAX"].ToString()))
                            {
                                if (!ALTER5MAX)
                                {
                                    error = true;
                                    ALTER5MAX = true;
                                    errorMsg += "ALTER5MAX can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ALTER5MAX"].ToString().Length > 5)
                                {
                                    if (!ALTER5MAXLen)
                                    {
                                        error = true;
                                        ALTER5MAXLen = true;
                                    }
                                    worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ALTER5MAX"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 7].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["CATEGORY_FLAG"].ToString()))
                            {
                                if (!CATEGORY_FLAG)
                                {
                                    error = true;
                                    CATEGORY_FLAG = true;
                                    errorMsg += "CATEGORY_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 8].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CATEGORY_FLAG"].ToString().Length > 1)
                                {
                                    if (!CATEGORY_FLAGLen)
                                    {
                                        error = true;
                                        CATEGORY_FLAGLen = true;
                                    }
                                    worksheet.Cells[i + 1, 8].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CATEGORY_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 8].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["CHAR_TYPE"].ToString()))
                            {
                                if (!CHAR_TYPE)
                                {
                                    error = true;
                                    CHAR_TYPE = true;
                                    errorMsg += "CHAR_TYPE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CHAR_TYPE"].ToString().Length > 20)
                                {
                                    if (!CHAR_TYPELen)
                                    {
                                        error = true;
                                        CHAR_TYPELen = true;
                                    }
                                    worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CHAR_TYPE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 9].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["NUMERIC_FLAG"].ToString()))
                            {
                                if (!NUMERIC_FLAG)
                                {
                                    error = true;
                                    NUMERIC_FLAG = true;
                                    errorMsg += "NUMERIC_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 10].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["NUMERIC_FLAG"].ToString().Length > 1)
                                {
                                    if (!NUMERIC_FLAGLen)
                                    {
                                        error = true;
                                        NUMERIC_FLAGLen = true;
                                    }
                                    worksheet.Cells[i + 1, 10].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["NUMERIC_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 10].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["FIXED_ITEM_VALUE"].ToString()))
                            {
                                if (!FIXED_ITEM_VALUE)
                                {
                                    error = true;
                                    FIXED_ITEM_VALUE = true;
                                    errorMsg += "FIXED_ITEM_VALUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 12].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FIXED_ITEM_VALUE"].ToString().Length > 1)
                                {
                                    if (!FIXED_ITEM_VALUELen)
                                    {
                                        error = true;
                                        FIXED_ITEM_VALUELen = true;
                                    }
                                    worksheet.Cells[i + 1, 12].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["FIXED_ITEM_VALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 12].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["COPY_ITEM"].ToString()))
                            {
                                if (!COPY_ITEM)
                                {
                                    error = true;
                                    COPY_ITEM = true;
                                    errorMsg += "COPY_ITEM can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 13].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["COPY_ITEM"].ToString().Length > 1)
                                {
                                    if (!COPY_ITEMLen)
                                    {
                                        error = true;
                                        COPY_ITEMLen = true;
                                    }
                                    worksheet.Cells[i + 1, 13].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["COPY_ITEM"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 13].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["MULTI_VALUE"].ToString()))
                            {
                                if (!MULTI_VALUE)
                                {
                                    error = true;
                                    MULTI_VALUE = true;
                                    errorMsg += "MULTI_VALUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 14].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MULTI_VALUE"].ToString().Length > 1)
                                {
                                    if (!MULTI_VALUELen)
                                    {
                                        error = true;
                                        MULTI_VALUELen = true;
                                    }
                                    worksheet.Cells[i + 1, 14].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["MULTI_VALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 14].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["ABBREVIATE_VALUE"].ToString()))
                            {
                                if (!ABBREVIATE_VALUE)
                                {
                                    error = true;
                                    ABBREVIATE_VALUE = true;
                                     errorMsg += "ABBREVIATE_VALUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 15].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ABBREVIATE_VALUE"].ToString().Length > 1)
                                {
                                    if (!ABBREVIATE_VALUELen)
                                    {
                                        error = true;
                                        ABBREVIATE_VALUELen = true;
                                    }
                                    worksheet.Cells[i + 1, 15].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ABBREVIATE_VALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 15].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["FIXED_VALUE_LIST"].ToString()))
                            {
                                if (!FIXED_VALUE_LIST)
                                {
                                    error = true;
                                    FIXED_VALUE_LIST = true;
                                     errorMsg += "FIXED_VALUE_LIST can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 16].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FIXED_VALUE_LIST"].ToString().Length > 1)
                                {
                                    if (!FIXED_VALUE_LISTLen)
                                    {
                                        error = true;
                                        FIXED_VALUE_LISTLen = true;
                                    }
                                    worksheet.Cells[i + 1, 16].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["FIXED_VALUE_LIST"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 16].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["TRANSLATION_IND"].ToString()))
                            {
                                if (!TRANSLATION_IND)
                                {
                                    error = true;
                                    TRANSLATION_IND = true;
                                    errorMsg += "TRANSLATION_IND can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 17].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["TRANSLATION_IND"].ToString().Length > 1)
                                {
                                    if (!TRANSLATION_INDLen)
                                    {
                                        error = true;
                                        TRANSLATION_INDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 17].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["TRANSLATION_IND"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 17].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["CHR_VAL_DESCRIPTION_ONLY"].ToString()))
                            {
                                if (!CHR_VAL_DESCRIPTION_ONLY)
                                {
                                    error = true;
                                    CHR_VAL_DESCRIPTION_ONLY = true;
                                     errorMsg += "CHR_VAL_DESCRIPTION_ONLY can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 18].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CHR_VAL_DESCRIPTION_ONLY"].ToString().Length > 1)
                                {
                                    if (!CHR_VAL_DESCRIPTION_ONLYLen)
                                    {
                                        error = true;
                                        CHR_VAL_DESCRIPTION_ONLYLen = true;
                                    }
                                    worksheet.Cells[i + 1, 18].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CHR_VAL_DESCRIPTION_ONLY"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 18].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["LOCAL"].ToString()))
                            {
                                if (!LOCAL)
                                {
                                    error = true;
                                    LOCAL = true;
                                     errorMsg += "LOCAL can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 19].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["LOCAL"].ToString().Length > 1)
                                {
                                    if (!LOCALLen)
                                    {
                                        error = true;
                                        LOCALLen = true;
                                    }
                                    worksheet.Cells[i + 1, 19].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["LOCAL"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 19].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["SORT_ORDER"].ToString()))
                            {
                                if (!SORT_ORDER)
                                {
                                    error = true;
                                    SORT_ORDER = true;
                                      errorMsg += "SORT_ORDER can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 20].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SORT_ORDER"].ToString().Length > 10)
                                {
                                    if (!SORT_ORDERLen)
                                    {
                                        error = true;
                                        SORT_ORDERLen = true;
                                    }
                                    worksheet.Cells[i + 1, 20].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SORT_ORDER"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 20].FillColor = Color.Empty; //BGRF-2086
                                }
                            }


                            if (!string.IsNullOrEmpty(table.Rows[i]["OGRDSID"].ToString()))
                            {
                                if (table.Rows[i]["OGRDSID"].ToString().Length > 10)
                                {
                                    if (!OGRDSIDLen)
                                    {
                                        error = true;
                                        OGRDSIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OGRDSID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["OTHERDESC"].ToString()))
                            {
                                if (table.Rows[i]["OTHERDESC"].ToString().Length > 70)
                                {
                                    if (!OTHERDESCLen)
                                    {
                                        error = true;
                                        OTHERDESCLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OTHERDESC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["ALTER30MAX_OTHER"].ToString()))
                            {
                                if (table.Rows[i]["ALTER30MAX_OTHER"].ToString().Length > 70)
                                {
                                    if (!ALTER30MAX_OTHERLen)
                                    {
                                        error = true;
                                        ALTER30MAX_OTHERLen = true;
                                    }
                                    worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ALTER30MAX_OTHER"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["CUSTOMER_FLAG"].ToString()))
                            {
                                if (table.Rows[i]["CUSTOMER_FLAG"].ToString().Length > 1)
                                {
                                    if (!CUSTOMER_FLAGLen)
                                    {
                                        error = true;
                                        CUSTOMER_FLAGLen = true;
                                    }
                                    worksheet.Cells[i + 1, 11].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CUSTOMER_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 11].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_DIC_LVL_FLAG_SETTING\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempDicLvlFlagSetting.Clear();
                    dtbTempDicLvlFlagSetting.Merge(table);
                }
                #endregion
                #region O_PC_ASGN_LVL_FLAG_SETTING
                if (ds.Tables[tableNumber].TableName == "O_PC_ASGN_LVL_FLAG_SETTING")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool CATEGORYATTRVALUE = false;
                    bool ATTRNO = false;
                    bool FIXED_VALUE_LIST = false;
                    bool COPY_ITEM_VAL = false;
                    bool FIELD_COLL_FLAG = false;
                    bool MANDATORY_FLAG = false;
                    bool ALIGNED = false;
                    bool UNIQUE = false;
                    bool LOCAL = false;
                    bool SORT_ORDER = false;

                    bool CATEGORYATTRVALUELen = false;
                    bool ATTRNOLen = false;
                    bool FIXED_VALUE_LISTLen = false;
                    bool COPY_ITEM_VALLen = false;
                    bool FIELD_COLL_FLAGLen = false;
                    bool MANDATORY_FLAGLen = false;
                    bool ALIGNEDLen = false;
                    bool UNIQUELen = false;
                    bool LOCALLen = false;
                    bool SORT_ORDERLen = false;

                    var worksheet = workbook.Worksheets["O_PC_ASGN_LVL_FLAG_SETTING"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["FIXED_VALUE_LIST"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["COPY_ITEM_VAL"].ToString())
                             && string.IsNullOrEmpty(table.Rows[i]["FIELD_COLL_FLAG"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["MANDATORY_FLAG"].ToString())
                             && string.IsNullOrEmpty(table.Rows[i]["ALIGNED"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["UNIQUE"].ToString())
                             && string.IsNullOrEmpty(table.Rows[i]["LOCAL"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["SORT_ORDER"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString()))
                            {
                                if (!CATEGORYATTRVALUE)
                                {
                                    error = true;
                                    CATEGORYATTRVALUE = true;
                                    errorMsg += "CATEGORYATTRVALUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 20)
                                {
                                    if (!CATEGORYATTRVALUELen)
                                    {
                                        error = true;
                                        CATEGORYATTRVALUELen = true;
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()))
                            {
                                if (!ATTRNO)
                                {
                                    error = true;
                                    ATTRNO = true;
                                    errorMsg += "ATTRNO can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ATTRNO"].ToString().Length > 10)
                                {
                                    if (!ATTRNOLen)
                                    {
                                        error = true;
                                        ATTRNOLen = true;
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ATTRNO"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["FIXED_VALUE_LIST"].ToString()))
                            {
                                if (!FIXED_VALUE_LIST)
                                {
                                    error = true;
                                    FIXED_VALUE_LIST = true;
                                    errorMsg += "FIXED_VALUE_LIST can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FIXED_VALUE_LIST"].ToString().Length > 1)
                                {
                                    if (!FIXED_VALUE_LISTLen)
                                    {
                                        error = true;
                                        FIXED_VALUE_LISTLen = true;
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["FIXED_VALUE_LIST"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["COPY_ITEM_VAL"].ToString()))
                            {
                                if (!COPY_ITEM_VAL)
                                {
                                    error = true;
                                    COPY_ITEM_VAL = true;
                                    errorMsg += "COPY_ITEM_VAL can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["COPY_ITEM_VAL"].ToString().Length > 1)
                                {
                                    if (!COPY_ITEM_VALLen)
                                    {
                                        error = true;
                                        COPY_ITEM_VALLen = true;
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["COPY_ITEM_VAL"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["FIELD_COLL_FLAG"].ToString()))
                            {
                                if (!FIELD_COLL_FLAG)
                                {
                                    error = true;
                                    FIELD_COLL_FLAG = true;
                                    errorMsg += "FIELD_COLL_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FIELD_COLL_FLAG"].ToString().Length > 1)
                                {
                                    if (!FIELD_COLL_FLAGLen)
                                    {
                                        error = true;
                                        FIELD_COLL_FLAGLen = true;
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["FIELD_COLL_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["MANDATORY_FLAG"].ToString()))
                            {
                                if (!MANDATORY_FLAG)
                                {
                                    error = true;
                                    MANDATORY_FLAG = true;
                                    errorMsg += "MANDATORY_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MANDATORY_FLAG"].ToString().Length > 1)
                                {
                                    if (!MANDATORY_FLAGLen)
                                    {
                                        error = true;
                                        MANDATORY_FLAGLen = true;
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["MANDATORY_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["ALIGNED"].ToString()))
                            {
                                if (!ALIGNED)
                                {
                                    error = true;
                                    ALIGNED = true;
                                    errorMsg += "ALIGNED can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ALIGNED"].ToString().Length > 1)
                                {
                                    if (!ALIGNEDLen)
                                    {
                                        error = true;
                                        ALIGNEDLen = true;
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ALIGNED"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["UNIQUE"].ToString()))
                            {
                                if (!UNIQUE)
                                {
                                    error = true;
                                    UNIQUE = true;
                                    errorMsg += "UNIQUE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["UNIQUE"].ToString().Length > 1)
                                {
                                    if (!UNIQUELen)
                                    {
                                        error = true;
                                        UNIQUELen = true;
                                        worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["UNIQUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 7].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["LOCAL"].ToString()))
                            {
                                if (!LOCAL)
                                {
                                    error = true;
                                    LOCAL = true;
                                    errorMsg += "LOCAL can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 8].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["LOCAL"].ToString().Length > 1)
                                {
                                    if (!LOCALLen)
                                    {
                                        error = true;
                                        LOCALLen = true;
                                        worksheet.Cells[i + 1, 8].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["LOCAL"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 8].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["SORT_ORDER"].ToString()))
                            {
                                if (!SORT_ORDER)
                                {
                                    error = true;
                                    SORT_ORDER = true;
                                    errorMsg += "SORT_ORDER can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SORT_ORDER"].ToString().Length > 10)
                                {
                                    if (!SORT_ORDERLen)
                                    {
                                        error = true;
                                        SORT_ORDERLen = true;
                                        worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["SORT_ORDER"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 9].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }
                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_PC_ASGN_LVL_FLAG_SETTING\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempPCAsgnlvlFlagSetting.Clear();
                    dtbTempPCAsgnlvlFlagSetting.Merge(table);
                }
                #endregion
                #region O_MOD_ASGN_LVL_FLAG_SETTING
                if (ds.Tables[tableNumber].TableName == "O_MOD_ASGN_LVL_FLAG_SETTING")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool MODULEID = false;
                    bool ATTRNO = false;
                    bool FIELD_COLL_FLAG = false;
                    bool MANDATORY_FLAG = false;

                    bool MODULEIDLen = false;
                    bool ATTRNOLen = false;
                    bool FIELD_COLL_FLAGLen = false;
                    bool MANDATORY_FLAGLen = false;

                    var worksheet = workbook.Worksheets["O_MOD_ASGN_LVL_FLAG_SETTING"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["MODULEID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["FIELD_COLL_FLAG"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["MANDATORY_FLAG"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["MODULEID"].ToString()))
                            {
                                if (!MODULEID)
                                {
                                    error = true;
                                    MODULEID = true;
                                    errorMsg += "MODULEID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MODULEID"].ToString().Length > 10)
                                {
                                    if (!MODULEIDLen)
                                    {
                                        error = true;
                                        MODULEIDLen = true;
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["MODULEID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()))
                            {
                                if (!ATTRNO)
                                {
                                    error = true;
                                    ATTRNO = true;
                                    errorMsg += "ATTRNO can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ATTRNO"].ToString().Length > 10)
                                {
                                    if (!ATTRNOLen)
                                    {
                                        error = true;
                                        ATTRNOLen = true;
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ATTRNO"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["FIELD_COLL_FLAG"].ToString()))
                            {
                                if (!FIELD_COLL_FLAG)
                                {
                                    error = true;
                                    FIELD_COLL_FLAG = true;
                                    errorMsg += "FIELD_COLL_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["FIELD_COLL_FLAG"].ToString().Length > 1)
                                {
                                    if (!FIELD_COLL_FLAGLen)
                                    {
                                        error = true;
                                        FIELD_COLL_FLAGLen = true;
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["FIELD_COLL_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["MANDATORY_FLAG"].ToString()))
                            {
                                if (!MANDATORY_FLAG)
                                {
                                    error = true;
                                    MANDATORY_FLAG = true;
                                    errorMsg += "MANDATORY_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MANDATORY_FLAG"].ToString().Length > 1)
                                {
                                    if (!MANDATORY_FLAGLen)
                                    {
                                        error = true;
                                        MANDATORY_FLAGLen = true;
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["MANDATORY_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_MOD_ASGN_LVL_FLAG_SETTING\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempModAsgnlvlFlagSetting.Clear();
                    dtbTempModAsgnlvlFlagSetting.Merge(table);
                }
                #endregion
                #region O_HEADING_TYPE
                if (ds.Tables[tableNumber].TableName == "O_HEADING_TYPE")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool HEADING_ID = false;
                    bool HEADING_NAME = false;
                    bool MAXIMUM_LENGTH = false;
                    bool ONLINE_HDNG_FLAG = false;
                    bool OUTPUT_FLAG = false;
                    bool ACTIVE_FLAG = false;
                    bool ALTERNATE_HEADING_FLAG = false;
                    bool ASGN_LVL = false;
                    bool LOCALLANG = false;

                    bool HEADING_IDLen = false;
                    bool HEADING_NAMELen = false;
                    bool MAXIMUM_LENGTHLen = false;
                    bool ONLINE_HDNG_FLAGLen = false;
                    bool OUTPUT_FLAGLen = false;
                    bool ACTIVE_FLAGLen = false;
                    bool ALTERNATE_HEADING_FLAGLen = false;
                    bool ASGN_LVLLen = false;
                    bool LOCALLANGLen = false;
                    bool HDNG_TYPLen = false;

                    var worksheet = workbook.Worksheets["O_HEADING_TYPE"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["HEADING_ID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["MAXIMUM_LENGTH"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["ONLINE_HDNG_FLAG"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["OUTPUT_FLAG"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ACTIVE_FLAG"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["ALTERNATE_HEADING_FLAG"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ASGN_LVL"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["LOCALLANG"].ToString())
                            )
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["HDNG_TYP"].ToString()))
                            {
                                if (table.Rows[i]["HDNG_TYP"].ToString().Length > 30)
                                {
                                    if (!HDNG_TYPLen)
                                    {
                                        error = true;
                                        HDNG_TYPLen = true;
                                    }
                                    worksheet.Cells[i + 1, 8].FillColor = Color.Red;
                                }
                                else
                                {
                                    if (table.Rows[i]["HDNG_TYP"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 8].FillColor = Color.Empty;
                                }
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["HEADING_ID"].ToString().Trim()))
                            {
                                if (!HEADING_ID)
                                {
                                    error = true;
                                    HEADING_ID = true;
                                    errorMsg += "HEADING_ID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["HEADING_ID"].ToString().Length > 30)
                                {
                                    if (!HEADING_IDLen)
                                    {
                                        error = true;
                                        HEADING_IDLen = true;
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["HEADING_ID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString()))
                            {
                                if (!HEADING_NAME)
                                {
                                    error = true;
                                    HEADING_NAME = true;
                                    errorMsg += "HEADING_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["HEADING_NAME"].ToString().Length > 50)
                                {
                                    if (!HEADING_NAMELen)
                                    {
                                        error = true;
                                        HEADING_NAMELen = true;
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["HEADING_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["MAXIMUM_LENGTH"].ToString()))
                            {
                                if (!MAXIMUM_LENGTH)
                                {
                                    error = true;
                                    MAXIMUM_LENGTH = true;
                                    errorMsg += "MAXIMUM_LENGTH can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MAXIMUM_LENGTH"].ToString().Length > 10)
                                {
                                    if (!MAXIMUM_LENGTHLen)
                                    {
                                        error = true;
                                        MAXIMUM_LENGTHLen = true;
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["MAXIMUM_LENGTH"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ONLINE_HDNG_FLAG"].ToString()))
                            {
                                if (!ONLINE_HDNG_FLAG)
                                {
                                    error = true;
                                    ONLINE_HDNG_FLAG = true;
                                    errorMsg += "ONLINE_HDNG_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ONLINE_HDNG_FLAG"].ToString().Length > 1)
                                {
                                    if (!ONLINE_HDNG_FLAGLen)
                                    {
                                        error = true;
                                        ONLINE_HDNG_FLAGLen = true;
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ONLINE_HDNG_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["OUTPUT_FLAG"].ToString()))
                            {
                                if (!OUTPUT_FLAG)
                                {
                                    error = true;
                                    OUTPUT_FLAG = true;
                                    errorMsg += "OUTPUT_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["OUTPUT_FLAG"].ToString().Length > 1)
                                {
                                    if (!OUTPUT_FLAGLen)
                                    {
                                        error = true;
                                        OUTPUT_FLAGLen = true;
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["OUTPUT_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ACTIVE_FLAG"].ToString()))
                            {
                                if (!ACTIVE_FLAG)
                                {
                                    error = true;
                                    ACTIVE_FLAG = true;
                                    errorMsg += "ACTIVE_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ACTIVE_FLAG"].ToString().Length > 1)
                                {
                                    if (!ACTIVE_FLAGLen)
                                    {
                                        error = true;
                                        ACTIVE_FLAGLen = true;
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ACTIVE_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ALTERNATE_HEADING_FLAG"].ToString()))
                            {
                                if (!ALTERNATE_HEADING_FLAG)
                                {
                                    error = true;
                                    ALTERNATE_HEADING_FLAG = true;
                                    errorMsg += "ALTERNATE_HEADING_FLAG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ALTERNATE_HEADING_FLAG"].ToString().Length > 1)
                                {
                                    if (!ALTERNATE_HEADING_FLAGLen)
                                    {
                                        error = true;
                                        ALTERNATE_HEADING_FLAGLen = true;
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ALTERNATE_HEADING_FLAG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 6].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ASGN_LVL"].ToString()))
                            {
                                if (!ASGN_LVL)
                                {
                                    error = true;
                                    ASGN_LVL = true;
                                    errorMsg += "ASGN_LVL can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ASGN_LVL"].ToString().Length > 1)
                                {
                                    if (!ASGN_LVLLen)
                                    {
                                        error = true;
                                        ASGN_LVLLen = true;
                                        worksheet.Cells[i + 1, 7].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["ASGN_LVL"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 7].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["LOCALLANG"].ToString()))
                            {
                                if (!LOCALLANG)
                                {
                                    error = true;
                                    LOCALLANG = true;
                                    errorMsg += "LOCALLANG can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["LOCALLANG"].ToString().Length > 1)
                                {
                                    if (!LOCALLANGLen)
                                    {
                                        error = true;
                                        LOCALLANGLen = true;
                                        worksheet.Cells[i + 1, 9].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["LOCALLANG"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 9].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (!string.IsNullOrEmpty(table.Rows[i]["HDNG_TYP"].ToString()))
                            {
                                if (table.Rows[i]["HDNG_TYP"].ToString().Length > 30)
                                {
                                    if (!HDNG_TYPLen)
                                    {
                                        error = true;
                                        HDNG_TYPLen = true;
                                    }
                                    worksheet.Cells[i + 1, 8].FillColor = Color.Red;
                                }
                                else
                                {
                                    if (table.Rows[i]["HDNG_TYP"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 8].FillColor = Color.Empty;
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_HEADING_TYPE\n";
                        //errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempHeadingType.Clear();
                    dtbTempHeadingType.Merge(table);
                }
                #endregion
                #region O_HEADING_PC_SETTING
                if (ds.Tables[tableNumber].TableName == "O_HEADING_PC_SETTING")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool HEADING_NAME = false;
                    bool CATEGORYATTRVALUE = false;
                    bool CHAR_ID = false;
                    bool SEQUENCE = false;
                    bool DROP_PRIORITY = false;

                    bool HEADING_NAMELen = false;
                    bool CATEGORYATTRVALUELen = false;
                    bool CHAR_IDLen = false;
                    bool SEQUENCELen = false;
                    bool DROP_PRIORITYLen = false;
                    bool CONCATENATORLen = false;

                    var worksheet = workbook.Worksheets["O_HEADING_PC_SETTING"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["CHAR_ID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["SEQUENCE"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["DROP_PRIORITY"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["CONCATENATOR"].ToString()))
                            {
                                if (table.Rows[i]["CONCATENATOR"].ToString().Length > 10)
                                {
                                    if (!CONCATENATORLen)
                                    {
                                        error = true;
                                        CONCATENATORLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CONCATENATOR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString().Trim()))
                            {
                                if (!HEADING_NAME)
                                {
                                    error = true;
                                    HEADING_NAME = true;
                                    errorMsg += "HEADING_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["HEADING_NAME"].ToString().Length > 50)
                                {
                                    if (!HEADING_NAMELen)
                                    {
                                        error = true;
                                        HEADING_NAMELen = true;
                                        errorMsg += "HEADING_NAME can not be Empty\n";
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["HEADING_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["CATEGORYATTRVALUE"].ToString()))
                            {
                                if (!CATEGORYATTRVALUE)
                                {
                                    error = true;
                                    CATEGORYATTRVALUE = true;
                                    errorMsg += "CATEGORYATTRVALUE can not be Empty\n";
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                            }
                            else
                            {
                                if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 20)
                                {
                                    if (!CATEGORYATTRVALUELen)
                                    {
                                        error = true;
                                        CATEGORYATTRVALUELen = true;
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["CHAR_ID"].ToString()))
                            {
                                if (!CHAR_ID)
                                {
                                    error = true;
                                    CHAR_ID = true;
                                    errorMsg += "CHAR_ID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CHAR_ID"].ToString().Length > 20)
                                {
                                    if (!CHAR_IDLen)
                                    {
                                        error = true;
                                        CHAR_IDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CHAR_ID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["SEQUENCE"].ToString()))
                            {
                                if (!SEQUENCE)
                                {
                                    error = true;
                                    SEQUENCE = true;
                                    errorMsg += "SEQUENCE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SEQUENCE"].ToString().Length > 5)
                                {
                                    if (!SEQUENCELen)
                                    {
                                        error = true;
                                        SEQUENCELen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SEQUENCE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["DROP_PRIORITY"].ToString()))
                            {
                                if (!DROP_PRIORITY)
                                {
                                    error = true;
                                    DROP_PRIORITY = true;
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                            }
                            else
                            {
                                if (table.Rows[i]["DROP_PRIORITY"].ToString().Length > 5)
                                {
                                    if (!DROP_PRIORITYLen)
                                    {
                                        error = true;
                                        DROP_PRIORITYLen = true;
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["DROP_PRIORITY"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["CONCATENATOR"].ToString()))
                            {
                                if (table.Rows[i]["CONCATENATOR"].ToString().Length > 10)
                                {
                                    if (!CONCATENATORLen)
                                    {
                                        error = true;
                                        CONCATENATORLen = true;
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                    }
                                }
                                else
                                {
                                    if (table.Rows[i]["CONCATENATOR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_HEADING_PC_SETTING\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempHeadingPCSetting.Clear();
                    dtbTempHeadingPCSetting.Merge(table);
                }
                #endregion
                #region O_HEADING_MOD_SETTING
                if (ds.Tables[tableNumber].TableName == "O_HEADING_MOD_SETTING")
                {
                    //start BGRF-2051
                    bool error = false;

                    bool HEADING_NAME = false;
                    bool MODULEID = false;
                    bool CHAR_ID = false;
                    bool SEQUENCE = false;
                    bool DROP_PRIORITY = false;

                    bool HEADING_NAMELen = false;
                    bool MODULEIDLen = false;
                    bool CHAR_IDLen = false;
                    bool SEQUENCELen = false;
                    bool DROP_PRIORITYLen = false;
                    bool CONCATENATORLen = false;


                    var worksheet = workbook.Worksheets["O_HEADING_MOD_SETTING"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {

                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["MODULEID"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["CHAR_ID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["SEQUENCE"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["DROP_PRIORITY"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                            if (!string.IsNullOrEmpty(table.Rows[i]["CONCATENATOR"].ToString()))
                            {
                                if (table.Rows[i]["CONCATENATOR"].ToString().Length > 10)
                                {
                                    if (!CONCATENATORLen)
                                    {
                                        error = true;
                                        CONCATENATORLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CONCATENATOR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["HEADING_NAME"].ToString()))
                            {
                                if (!HEADING_NAME)
                                {
                                    error = true;
                                    HEADING_NAME = true;
                                    errorMsg += "HEADING_NAME can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["HEADING_NAME"].ToString().Length > 50)
                                {
                                    if (!HEADING_NAMELen)
                                    {
                                        error = true;
                                        HEADING_NAMELen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["HEADING_NAME"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["MODULEID"].ToString()))
                            {
                                if (!MODULEID)
                                {
                                    error = true;
                                    MODULEID = true;
                                    errorMsg += "MODULEID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["MODULEID"].ToString().Length > 10)
                                {
                                    if (!MODULEIDLen)
                                    {
                                        error = true;
                                        MODULEIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["MODULEID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["CHAR_ID"].ToString()))
                            {
                                if (!CHAR_ID)
                                {
                                    error = true;
                                    CHAR_ID = true;
                                    errorMsg += "CHAR_ID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["CHAR_ID"].ToString().Length > 20)
                                {
                                    if (!CHAR_IDLen)
                                    {
                                        error = true;
                                        CHAR_IDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CHAR_ID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["SEQUENCE"].ToString()))
                            {
                                if (!SEQUENCE)
                                {
                                    error = true;
                                    SEQUENCE = true;
                                    errorMsg += "SEQUENCE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SEQUENCE"].ToString().Length > 5)
                                {
                                    if (!SEQUENCELen)
                                    {
                                        error = true;
                                        SEQUENCELen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SEQUENCE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["DROP_PRIORITY"].ToString()))
                            {
                                if (!DROP_PRIORITY)
                                {
                                    error = true;
                                    DROP_PRIORITY = true;
                                    errorMsg += "DROP_PRIORITY can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["DROP_PRIORITY"].ToString().Length > 5)
                                {
                                    if (!DROP_PRIORITYLen)
                                    {
                                        error = true;
                                        DROP_PRIORITYLen = true;
                                    }
                                    worksheet.Cells[i + 1, 5].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["DROP_PRIORITY"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 5].FillColor = Color.Empty; //BGRF-2086
                                }
                            }

                            if (!string.IsNullOrEmpty(table.Rows[i]["CONCATENATOR"].ToString()))
                            {
                                if (table.Rows[i]["CONCATENATOR"].ToString().Length > 10)
                                {
                                    if (!CONCATENATORLen)
                                    {
                                        error = true;
                                        CONCATENATORLen = true;
                                    }
                                    worksheet.Cells[i + 1, 4].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["CONCATENATOR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 4].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_HEADING_MOD_SETTING\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempHeadingModSetting.Clear();
                    dtbTempHeadingModSetting.Merge(table);
                }
                #endregion
                #region O_RETAILER_DEPT_SUPP
                if (ds.Tables[tableNumber].TableName == "O_RETAILER_DEPT_SUPP")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool SOURCEID = false;
                    bool ATTRNO = false;
                    bool TYPE = false;
                    bool SEQ = false;

                    bool SOURCEIDLen = false;
                    bool ATTRNOLen = false;
                    bool TYPELen = false;
                    bool SEQLen = false;

                    var worksheet = workbook.Worksheets["O_RETAILER_DEPT_SUPP"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["SOURCEID"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["SEQ"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["SOURCEID"].ToString()))
                            {
                                if (!SOURCEID)
                                {
                                    error = true;
                                    SOURCEID = true;
                                    errorMsg += "SOURCEID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SOURCEID"].ToString().Length > 15)
                                {
                                    if (!SOURCEIDLen)
                                    {
                                        error = true;
                                        SOURCEIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SOURCEID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["ATTRNO"].ToString()))
                            {
                                if (!ATTRNO)
                                {
                                    error = true;
                                    ATTRNO = true;
                                    errorMsg += "ATTRNO can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["ATTRNO"].ToString().Length > 15)
                                {
                                    if (!ATTRNOLen)
                                    {
                                        error = true;
                                        ATTRNOLen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["ATTRNO"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["TYPE"].ToString()))
                            {
                                if (!TYPE)
                                {
                                    error = true;
                                    TYPE = true;
                                    errorMsg += "TYPE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["TYPE"].ToString().Length > 30)
                                {
                                    if (!TYPELen)
                                    {
                                        error = true;
                                        TYPELen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["TYPE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["SEQ"].ToString()))
                            {
                                if (!SEQ)
                                {
                                    error = true;
                                    SEQ = true;
                                    errorMsg += "SEQ can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["SEQ"].ToString().Length > 5)
                                {
                                    if (!SEQLen)
                                    {
                                        error = true;
                                        SEQLen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["SEQ"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_RETAILER_DEPT_SUPP\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempRetailerDeptSupp.Clear();
                    dtbTempRetailerDeptSupp.Merge(table);
                }
                #endregion
                #region O_UOM_MAPPING_LIST
                if (ds.Tables[tableNumber].TableName == "O_UOM_MAPPING_LIST")
                {
                    //start BGRF-2051
                    bool error = false;
                    bool RF_UOMCODE = false;
                    bool OGRDS_UOMID = false;
                    bool OGRDS_UOMDESC = false;
                    bool PREFERED_UNIT_DSCR = false;

                    bool RF_UOMCODELen = false;
                    bool OGRDS_UOMIDLen = false;
                    bool OGRDS_UOMDESCLen = false;
                    bool PREFERED_UNIT_DSCRLen = false;

                    var worksheet = workbook.Worksheets["O_UOM_MAPPING_LIST"]; //BGRF-2086

                    string errorMsg = string.Empty;
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //BGRF-2086 : when whole row is null removing red color 
                        if (string.IsNullOrEmpty(table.Rows[i]["RF_UOMCODE"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["OGRDS_UOMID"].ToString())
                            && string.IsNullOrEmpty(table.Rows[i]["OGRDS_UOMDESC"].ToString()) && string.IsNullOrEmpty(table.Rows[i]["PREFERED_UNIT_DSCR"].ToString()))
                        {
                            for (int col = 0; col < table.Columns.Count; col++)
                            {
                                worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(table.Rows[i]["RF_UOMCODE"].ToString()))
                            {
                                if (!RF_UOMCODE)
                                {
                                    error = true;
                                    RF_UOMCODE = true;
                                    errorMsg += "RF_UOMCODE can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["RF_UOMCODE"].ToString().Length > 20)
                                {
                                    if (!RF_UOMCODELen)
                                    {
                                        error = true;
                                        RF_UOMCODELen = true;
                                    }
                                    worksheet.Cells[i + 1, 0].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["RF_UOMCODE"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 0].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["OGRDS_UOMID"].ToString()))
                            {
                                if (!OGRDS_UOMID)
                                {
                                    error = true;
                                    OGRDS_UOMID = true;
                                    errorMsg += "OGRDS_UOMID can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["OGRDS_UOMID"].ToString().Length > 5)
                                {
                                    if (!OGRDS_UOMIDLen)
                                    {
                                        error = true;
                                        OGRDS_UOMIDLen = true;
                                    }
                                    worksheet.Cells[i + 1, 1].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OGRDS_UOMID"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 1].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["OGRDS_UOMDESC"].ToString()))
                            {
                                if (!OGRDS_UOMDESC)
                                {
                                    error = true;
                                    OGRDS_UOMDESC = true;
                                    errorMsg += "OGRDS_UOMDESC can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["OGRDS_UOMDESC"].ToString().Length > 20)
                                {
                                    if (!OGRDS_UOMDESCLen)
                                    {
                                        error = true;
                                        OGRDS_UOMDESCLen = true;
                                    }
                                    worksheet.Cells[i + 1, 2].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["OGRDS_UOMDESC"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 2].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                            if (string.IsNullOrEmpty(table.Rows[i]["PREFERED_UNIT_DSCR"].ToString()))
                            {
                                if (!PREFERED_UNIT_DSCR)
                                {
                                    error = true;
                                    PREFERED_UNIT_DSCR = true;
                                    errorMsg += "PREFERED_UNIT_DSCR can not be Empty\n";
                                }
                                worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                if (table.Rows[i]["PREFERED_UNIT_DSCR"].ToString().Length > 30)
                                {
                                    if (!PREFERED_UNIT_DSCRLen)
                                    {
                                        error = true;
                                        PREFERED_UNIT_DSCRLen = true;
                                    }
                                    worksheet.Cells[i + 1, 3].FillColor = Color.Red; //BGRF-2086
                                }
                                else
                                {
                                    if (table.Rows[i]["PREFERED_UNIT_DSCR"].ToString().Length > 0)
                                        worksheet.Cells[i + 1, 3].FillColor = Color.Empty; //BGRF-2086
                                }
                            }
                        }
                    }

                    if (error)
                    {
                        errorRecord += "\n";
                        errorRecord += "Error tab : O_UOM_MAPPING_LIST\n";
                        // errorRecord += errorMsg; //BGRF-2086 for showing only error tab name in error msg
                    }
                    //end BGRF-2051

                    RemoveEmptyRows(table); // start BGRF-2086   
                    dtbTempUomMappingList.Clear();
                    dtbTempUomMappingList.Merge(table);
                }
                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        //Sudheer BGRF-2056
        private string ExcelColumnIndexToName(int Index)
        {
            string range = "";
            if (Index < 0) return range;
            for (int i = 1; Index + i > 0; i = 0)
            {
                range = ((char)(65 + Index % 26)).ToString() + range;
                Index /= 26;
            }
            if (range.Length > 1) range = ((char)((int)range[0] - 1)).ToString() + range.Substring(1);
            return range;
        }
        //Sudheer BGRF-2056
        public void LoadSpreadSheet()
        {
            saveButton.Enabled = false;
            saveButton.Refresh();

            //Start BGRF-2051
            dtbActiveCategory.Clear();
            dtbPreSetChar.Clear();
            dtbProcGrpSet.Clear();
            dtbExtCodeGrp.Clear();
            dtbActiveRetailer.Clear();
            dtbDicLvlFlagSetting.Clear();
            dtbPCAsgnlvlFlagSetting.Clear();
            dtbModAsgnlvlFlagSetting.Clear();
            dtbHeadingType.Clear();
            dtbHeadingPCSetting.Clear();
            dtbHeadingModSetting.Clear();
            dtbRetailerDeptSupp.Clear();
            dtbUomMappingList.Clear();

            SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
            try
            {
                blyBL = new BusinessLayer();
                ds.Clear();
                dsTemp.Clear();
                blyBL.MDLInputSettings_Load();

                if (ds.Tables.Count < 13)
                {
                    ds.Tables.Add(dtbActiveCategory);
                    ds.Tables.Add(dtbPreSetChar);
                    ds.Tables.Add(dtbProcGrpSet);
                    ds.Tables.Add(dtbExtCodeGrp);
                    ds.Tables.Add(dtbActiveRetailer);
                    ds.Tables.Add(dtbDicLvlFlagSetting);
                    ds.Tables.Add(dtbPCAsgnlvlFlagSetting);
                    ds.Tables.Add(dtbModAsgnlvlFlagSetting);
                    ds.Tables.Add(dtbHeadingType);
                    ds.Tables.Add(dtbHeadingPCSetting);
                    ds.Tables.Add(dtbHeadingModSetting);
                    ds.Tables.Add(dtbRetailerDeptSupp);
                    ds.Tables.Add(dtbUomMappingList);


                    //Sudheer BGRF-2087
                    dsTemp.Tables.Add(dtbTempActiveCategory);
                    dsTemp.Tables.Add(dtbTempPreSetChar);
                    dsTemp.Tables.Add(dtbTempProcGrpSet);
                    dsTemp.Tables.Add(dtbTempExtCodeGrp);
                    dsTemp.Tables.Add(dtbTempActiveRetailer);
                    dsTemp.Tables.Add(dtbTempDicLvlFlagSetting);
                    dsTemp.Tables.Add(dtbTempPCAsgnlvlFlagSetting);
                    dsTemp.Tables.Add(dtbTempModAsgnlvlFlagSetting);
                    dsTemp.Tables.Add(dtbTempHeadingType);
                    dsTemp.Tables.Add(dtbTempHeadingPCSetting);
                    dsTemp.Tables.Add(dtbTempHeadingModSetting);
                    dsTemp.Tables.Add(dtbTempRetailerDeptSupp);
                    dsTemp.Tables.Add(dtbTempUomMappingList);

                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        if (i == 0)
                        {
                            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
                            spreadsheetControl1.Document.Worksheets[0].Name = ds.Tables[i].TableName;
                            worksheet.Import(ds.Tables[0], true, 0, 0);
                        }
                        else
                        {
                            Worksheet worksheet = spreadsheetControl1.Document.Worksheets.Add(ds.Tables[i].TableName);
                            worksheet.Import(ds.Tables[i], true, 0, 0);
                        }

                    }
                }

                else
                {
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        Worksheet worksheet = spreadsheetControl1.Document.Worksheets[i];
                        worksheet.Import(ds.Tables[i], true, 0, 0);
                    }
                }
                saveButton.Enabled = true;
                saveButton.Refresh();
                SplashScreenManager.CloseForm(false);
            }
            catch (Exception ex)
            {
                saveButton.Enabled = true;
                saveButton.Refresh();
                SplashScreenManager.CloseForm(false);
                throw ex;
            }

        }
        private void ProcessTool_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);//BGRF-2086
        }
        //Sudheer BGRF-2056
        private void saveButton_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            blyBL = new BusinessLayer();
            errorRecord = string.Empty;
                try
                {
                    SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                    SplashScreenManager.Default.SetWaitFormCaption("Save data");
                    //Disable controls while saving.
                    btnExport.Enabled = false;
                    saveButton.Enabled = false;
                    btnHelp.Enabled = false;
                    bool blnChanges = false; ;
                    //Save Spreadsheet data into Datatable.
                    for (int i = 0; i < ds.Tables.Count; i++)
                    {
                        CreateDt(i);
                    }

                    if (string.IsNullOrEmpty(errorRecord)) //BGRF-2051
                    {
                        //Compare data before saving data.
                        for (int i = 0; i < ds.Tables.Count; i++)
                        {
                            RemoveColor(i); //BGRF-2086
                            if (ds.Tables[i].Rows.Count != dsTemp.Tables[i].Rows.Count | ds.Tables[i].AsEnumerable().Union(dsTemp.Tables[i].AsEnumerable(), DataRowComparer.Default).Count() != ds.Tables[i].Rows.Count)
                            {
                                //Save Data
                                blyBL.MDLInputSettings_PopulateData(dsTemp.Tables[i].TableName, dsTemp.Tables[i]);
                                blnChanges = true;
                            }
                           
                        }
                       
                        if (blnChanges)
                        {
                            LoadSpreadSheet();
                        }

                        SplashScreenManager.ShowForm(this, typeof(WaitForm1), true, true, false);
                        SplashScreenManager.Default.SetWaitFormCaption("Save data - Success");
                        SplashScreenManager.CloseForm(false);
                        btnExport.Enabled = true;
                        saveButton.Enabled = true;
                        btnHelp.Enabled = true;
                        MessageBox.Show("Load MDL Input Settings File process - Success", "Info", buttons, MessageBoxIcon.Information);
                    }
                    else
                    {
                        SplashScreenManager.CloseForm(false);
                        btnExport.Enabled = true;
                        saveButton.Enabled = true;
                       btnHelp.Enabled = true;
                       HighlightErrorSheet();  //BGRF-2086  
                        MessageBox.Show(errorRecord, "Error", buttons, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    SplashScreenManager.CloseForm(false);
                    btnExport.Enabled = true;
                    saveButton.Enabled = true;
                    btnHelp.Enabled = true;
                    MessageBox.Show("Load MDL Input Settings File process - Failed", "Error", buttons, MessageBoxIcon.Error);
                }
        }

        //Sudheer BGRF-2056

        #region BGRF-2051 ValiDation for spreadsheet
        private void spreadsheetControl1_CellValueChanged(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellEventArgs e)
        {
            CellDataValidation();//BGRF-2051
            CellSpecialCharValidation();
            NumericCell();//BGRF-2086
        }
        private void spreadsheetControl1_KeyDown(object sender, KeyEventArgs e)
        {
            CellDataValidation();//BGRF-2051
            IWorkbook workbook = spreadsheetControl1.Document;
            var worksheet = workbook.Worksheets.ActiveWorksheet;

            var columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;

            if (spreadsheetControl1.ActiveCell.Value.Type.ToString() == "Text" || spreadsheetControl1.ActiveCell.Value.Type.ToString() == "None")
            {
                var columnValue = spreadsheetControl1.ActiveCell.Value.TextValue;

                if (e.KeyData.ToString() == "Tab" && string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
                else if (e.KeyData.ToString() == "Return" && string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
                else if (e.KeyData.ToString() == "Delete" && !string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
            }
            if (spreadsheetControl1.ActiveCell.Value.Type.ToString() == "Numeric")
            {
                var columnValue = Convert.ToString(spreadsheetControl1.ActiveCell.Value.IsNumeric);
                if (e.KeyData.ToString() == "Tab" && string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
                else if (e.KeyData.ToString() == "Return" && string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
                else if (e.KeyData.ToString() == "Delete" && !string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
            }
        }
        private void CellDataValidation()
        {
            if (!cellDataValidation)
            {
                cellDataValidation = true;
                // O_ACTIVE_CATEGORY
                IWorkbook workbook = spreadsheetControl1.Document;
                var worksheet1 = workbook.Worksheets["O_ACTIVE_CATEGORY"];
                DataValidation validation_1A = worksheet1.DataValidations.Add(worksheet1["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.LessThanOrEqual, 20);
                DataValidation validation_1B = worksheet1.DataValidations.Add(worksheet1["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.LessThanOrEqual, 20);
                DataValidation validation_1C = worksheet1.DataValidations.Add(worksheet1["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 10);

                validation_1A.ErrorTitle = "Wrong CATEGORYATTRVALUE";
                validation_1A.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_1A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_1A.ShowErrorMessage = true;

                validation_1B.ErrorTitle = "Wrong SERVICE";
                validation_1B.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_1B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_1B.ShowErrorMessage = true;

                validation_1C.ErrorTitle = "Wrong TYPE";
                validation_1C.ErrorMessage = "The value you entered is not valid. Use 1-10 Alphabet/Numeric characters value.";
                validation_1C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_1C.ShowErrorMessage = true;

                //worksheet1["$A:$XFD"].Protection.Locked = false;
                //worksheet1["A1:C1"].Protection.Locked = true;
                //worksheet1.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);

                // O_PRE_SET_CHAR
                var worksheet2 = workbook.Worksheets["O_PRE_SET_CHAR"];
                DataValidation validation_2A = worksheet2.DataValidations.Add(worksheet2["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);
                DataValidation validation_2B = worksheet2.DataValidations.Add(worksheet2["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=15)");
                DataValidation validation_2C = worksheet2.DataValidations.Add(worksheet2["C2:C1048576"], DataValidationType.List, "Y, N");

                validation_2A.ErrorTitle = "Wrong TYPE";
                validation_2A.ErrorMessage = "The value you entered is not valid. Use 1-15 Alphabet/Numeric characters value.";
                validation_2A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_2A.ShowErrorMessage = true;

                validation_2B.ErrorTitle = "Wrong ATTRNO";
                validation_2B.ErrorMessage = "The value you entered is not valid. Use 1-15 digit Numeric value.";
                validation_2B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_2B.ShowErrorMessage = true;

                validation_2C.ErrorTitle = "Wrong REQUIRED";
                validation_2C.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_2C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_2C.ShowErrorMessage = true;

                //worksheet2["$A:$XFD"].Protection.Locked = false; //
                //worksheet2["A1:C1"].Protection.Locked = true;
                //worksheet2.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);

                //O_PROC_GRP_SET
                var worksheet3 = workbook.Worksheets["O_PROC_GRP_SET"];
                DataValidation validation_3A = worksheet3.DataValidations.Add(worksheet3["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_3B = worksheet3.DataValidations.Add(worksheet3["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_3C = worksheet3.DataValidations.Add(worksheet3["C2:C1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_3D = worksheet3.DataValidations.Add(worksheet3["D2:D1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_3E = worksheet3.DataValidations.Add(worksheet3["E2:E1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_3F = worksheet3.DataValidations.Add(worksheet3["F2:F1048576"], DataValidationType.List, "Y, N");

                validation_3A.ErrorTitle = "Wrong PROC_GROUP_SET_ID";
                validation_3A.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_3A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3A.ShowErrorMessage = true;

                validation_3B.ErrorTitle = "Wrong PROC_GROUP_SET_NAME";
                validation_3B.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_3B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3B.ShowErrorMessage = true;

                validation_3C.ErrorTitle = "Wrong FOLLOW_EXT_RULE";
                validation_3C.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_3C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3C.ShowErrorMessage = true;

                validation_3D.ErrorTitle = "Wrong EXCP_SENT_TO_XCD_BRWSR";
                validation_3D.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_3D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3D.ShowErrorMessage = true;

                validation_3E.ErrorTitle = "Wrong EXCP_SENT_TO_UNCDBLE_ITM";
                validation_3E.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_3E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3E.ShowErrorMessage = true;

                //null
                validation_3F.ErrorTitle = "Wrong EXCP_SEN_TO_SUPER_GTP_ITM";
                validation_3F.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_3F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_3F.ShowErrorMessage = true;

                //worksheet3["$A:$XFD"].Protection.Locked = false;
                //worksheet3["A1:F1"].Protection.Locked = true;
                //worksheet3.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);

                //O_EXT_CODE_GRP
                var worksheet4 = workbook.Worksheets["O_EXT_CODE_GRP"];
                DataValidation validation_4A = worksheet4.DataValidations.Add(worksheet4["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 10);
                DataValidation validation_4B = worksheet4.DataValidations.Add(worksheet4["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);

                validation_4A.ErrorTitle = "Wrong SHORTNAME";
                validation_4A.ErrorMessage = "The value you entered is not valid. Use 1-10 Alphabet/Numeric characters value.";
                validation_4A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_4A.ShowErrorMessage = true;

                validation_4B.ErrorTitle = "Wrong EXTERNAL_CODE_GROUP_NAME";
                validation_4B.ErrorMessage = "The value you entered is not valid. Use 1-30 Alphabet/Numeric characters value.";
                validation_4B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_4B.ShowErrorMessage = true;

                //worksheet4["$A:$XFD"].Protection.Locked = false;
                //worksheet4["A1:B1"].Protection.Locked = true;
                //worksheet4.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);

                //O_ACTIVE_RETAILER
                var worksheet5 = workbook.Worksheets["O_ACTIVE_RETAILER"];
                DataValidation validation_5A = worksheet5.DataValidations.Add(worksheet5["A2:A1048576"], DataValidationType.Custom, "=AND(ISNUMBER(A2:A1048576),LEN(A2:A1048576)<=10)");
                DataValidation validation_5B = worksheet5.DataValidations.Add(worksheet5["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_5C = worksheet5.DataValidations.Add(worksheet5["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);
                DataValidation validation_5D = worksheet5.DataValidations.Add(worksheet5["D2:D1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_5E = worksheet5.DataValidations.Add(worksheet5["E2:E1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_5F = worksheet5.DataValidations.Add(worksheet5["F2:F1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_5G = worksheet5.DataValidations.Add(worksheet5["G2:G1048576"], DataValidationType.List, "Y, N");

                validation_5A.ErrorTitle = "Wrong SOURCEID";
                validation_5A.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_5A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5A.ShowErrorMessage = true;

                validation_5B.ErrorTitle = "Wrong PROC_GROUP_SET_NAME";
                validation_5B.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_5B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5B.ShowErrorMessage = true;

                validation_5C.ErrorTitle = "Wrong TYPE";
                validation_5C.ErrorMessage = "The value you entered is not valid. Use 1-30 Alphabet/Numeric characters value.";
                validation_5C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5C.ShowErrorMessage = true;

                validation_5D.ErrorTitle = "Wrong ENN";
                validation_5D.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_5D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5D.ShowErrorMessage = true;

                validation_5E.ErrorTitle = "Wrong UPC";
                validation_5E.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_5E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5E.ShowErrorMessage = true;

                validation_5F.ErrorTitle = "Wrong LAC";
                validation_5F.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_5F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5F.ShowErrorMessage = true;

                validation_5G.ErrorTitle = "Wrong CIP";
                validation_5G.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_5G.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_5G.ShowErrorMessage = true;

                //worksheet5.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet5["$A:$XFD"].Protection.Locked = false;
                //worksheet5["A1:G1"].Protection.Locked = true;

                //O_DIC_LVL_FLAG_SETTING
                var worksheet6 = workbook.Worksheets["O_DIC_LVL_FLAG_SETTING"];
                DataValidation validation_6A = worksheet6.DataValidations.Add(worksheet6["A2:A1048576"], DataValidationType.Custom, "=AND(ISNUMBER(A2:A1048576),LEN(A2:A1048576)<=10)");
                DataValidation validation_6B = worksheet6.DataValidations.Add(worksheet6["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=10)");
                DataValidation validation_6C = worksheet6.DataValidations.Add(worksheet6["C2:C1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6D = worksheet6.DataValidations.Add(worksheet6["D2:D1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 70);
                DataValidation validation_6E = worksheet6.DataValidations.Add(worksheet6["E2:E1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 0, 70);
                DataValidation validation_6F = worksheet6.DataValidations.Add(worksheet6["F2:F1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 70);
                DataValidation validation_6G = worksheet6.DataValidations.Add(worksheet6["G2:G1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 0, 70);
                DataValidation validation_6H = worksheet6.DataValidations.Add(worksheet6["H2:H1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 5);
                DataValidation validation_6I = worksheet6.DataValidations.Add(worksheet6["I2:I1048576"], DataValidationType.TextLength, DataValidationOperator.Equal, 1); //P
                DataValidation validation_6J = worksheet6.DataValidations.Add(worksheet6["J2:J1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_6K = worksheet6.DataValidations.Add(worksheet6["K2:K1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6L = worksheet6.DataValidations.Add(worksheet6["L2:L1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6M = worksheet6.DataValidations.Add(worksheet6["M2:M1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6N = worksheet6.DataValidations.Add(worksheet6["N2:N1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6O = worksheet6.DataValidations.Add(worksheet6["O2:O1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6P = worksheet6.DataValidations.Add(worksheet6["P2:P1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6Q = worksheet6.DataValidations.Add(worksheet6["Q2:Q1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6R = worksheet6.DataValidations.Add(worksheet6["R2:R1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6S = worksheet6.DataValidations.Add(worksheet6["S2:S1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6T = worksheet6.DataValidations.Add(worksheet6["T2:T1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_6U = worksheet6.DataValidations.Add(worksheet6["U2:U1048576"], DataValidationType.Custom, "=AND(ISNUMBER(U2:U1048576),LEN(U2:U1048576)<=10)");

                validation_6A.ErrorTitle = "Wrong ATTRNO";
                validation_6A.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_6A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6A.ShowErrorMessage = true;

                //null
                validation_6B.ErrorTitle = "Wrong OGRDSID";
                validation_6B.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_6B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6B.ShowErrorMessage = true;

                validation_6C.ErrorTitle = "Wrong MAP_TO_OGRDS_CHR_VAL";
                validation_6C.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6C.ShowErrorMessage = true;

                validation_6D.ErrorTitle = "Wrong LONGDESC";
                validation_6D.ErrorMessage = "The value you entered is not valid.  Use 1-70 Alphabet/Numeric characters value.";
                validation_6D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6D.ShowErrorMessage = true;

                //null
                validation_6E.ErrorTitle = "Wrong OTHERDESC";
                validation_6E.ErrorMessage = "The value you entered is not valid. Use 1-70 Alphabet/Numeric characters value.";
                validation_6E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6E.ShowErrorMessage = true;

                validation_6F.ErrorTitle = "Wrong ALTER30MAX";
                validation_6F.ErrorMessage = "The value you entered is not valid. Use 1-70 Alphabet/Numeric characters value.";
                validation_6F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6F.ShowErrorMessage = true;

                //null
                validation_6G.ErrorTitle = "Wrong ALTER30MAX_OTHER";
                validation_6G.ErrorMessage = "The value you entered is not valid. Use 1-70 Alphabet/Numeric characters value.";
                validation_6G.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6G.ShowErrorMessage = true;

                validation_6H.ErrorTitle = "Wrong ALTER5MAX";
                validation_6H.ErrorMessage = "The value you entered is not valid.Use 1-5 Alphabet/Numeric characters value.";
                validation_6H.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6H.ShowErrorMessage = true;

                validation_6I.ErrorTitle = "Wrong CATEGORY_FLAG";
                validation_6I.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6I.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6I.ShowErrorMessage = true;

                validation_6J.ErrorTitle = "Wrong CHAR_TYPE";
                validation_6J.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_6J.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6J.ShowErrorMessage = true;

                validation_6K.ErrorTitle = "Wrong NUMERIC_FLAG";
                validation_6K.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6K.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6K.ShowErrorMessage = true;

                //null
                validation_6L.ErrorTitle = "Wrong CUSTOMER_FLAG";
                validation_6L.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6L.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6L.ShowErrorMessage = true;

                validation_6M.ErrorTitle = "Wrong FIXED_ITEM_VALUE";
                validation_6M.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6M.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6M.ShowErrorMessage = true;

                validation_6N.ErrorTitle = "Wrong COPY_ITEM";
                validation_6N.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6N.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6N.ShowErrorMessage = true;

                validation_6O.ErrorTitle = "Wrong MULTI_VALUE";
                validation_6O.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6O.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6O.ShowErrorMessage = true;

                validation_6P.ErrorTitle = "Wrong ABBREVIATE_VALUE";
                validation_6P.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6P.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6P.ShowErrorMessage = true;

                validation_6Q.ErrorTitle = "Wrong FIXED_VALUE_LIST";
                validation_6Q.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6Q.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6Q.ShowErrorMessage = true;

                validation_6R.ErrorTitle = "Wrong TRANSLATION_IND";
                validation_6R.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6R.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6R.ShowErrorMessage = true;

                validation_6S.ErrorTitle = "Wrong CHR_VAL_DESCRIPTION_ONLY";
                validation_6S.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6S.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6S.ShowErrorMessage = true;

                validation_6T.ErrorTitle = "Wrong LOCAL";
                validation_6T.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_6T.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6T.ShowErrorMessage = true;

                validation_6U.ErrorTitle = "Wrong SORT_ORDER";
                validation_6U.ErrorMessage = "The value you entered is not valid.  Use 1-10 digit Numeric value.";
                validation_6U.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_6U.ShowErrorMessage = true;

                //worksheet6.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet6["$A:$XFD"].Protection.Locked = false;
                //worksheet6["A1:U1"].Protection.Locked = true;

                //O_PC_ASGN_LVL_FLAG_SETTING
                var worksheet7 = workbook.Worksheets["O_PC_ASGN_LVL_FLAG_SETTING"];
                DataValidation validation_7A = worksheet7.DataValidations.Add(worksheet7["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_7B = worksheet7.DataValidations.Add(worksheet7["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=10)");
                DataValidation validation_7C = worksheet7.DataValidations.Add(worksheet7["C2:C1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7D = worksheet7.DataValidations.Add(worksheet7["D2:D1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7E = worksheet7.DataValidations.Add(worksheet7["E2:E1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7F = worksheet7.DataValidations.Add(worksheet7["F2:F1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7G = worksheet7.DataValidations.Add(worksheet7["G2:G1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7H = worksheet7.DataValidations.Add(worksheet7["H2:H1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7I = worksheet7.DataValidations.Add(worksheet7["I2:I1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_7J = worksheet7.DataValidations.Add(worksheet7["J2:J1048576"], DataValidationType.Custom, "=AND(ISNUMBER(J2:J1048576),LEN(J2:J1048576)<=10)");

                validation_7A.ErrorTitle = "Wrong CATEGORYATTRVALUE";
                validation_7A.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_7A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7A.ShowErrorMessage = true;

                validation_7B.ErrorTitle = "Wrong ATTRNO";
                validation_7B.ErrorMessage = "The value you entered is not valid.Use 1-10 digit Numeric value.";
                validation_7B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7B.ShowErrorMessage = true;

                validation_7C.ErrorTitle = "Wrong FIXED_VALUE_LIST";
                validation_7C.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7C.ShowErrorMessage = true;

                validation_7D.ErrorTitle = "Wrong COPY_ITEM_VAL";
                validation_7D.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7D.ShowErrorMessage = true;

                validation_7E.ErrorTitle = "Wrong FIELD_COLL_FLAG";
                validation_7E.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7E.ShowErrorMessage = true;

                validation_7F.ErrorTitle = "Wrong MANDATORY_FLAG";
                validation_7F.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7F.ShowErrorMessage = true;

                validation_7G.ErrorTitle = "Wrong ALIGNED";
                validation_7G.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7G.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7G.ShowErrorMessage = true;

                validation_7H.ErrorTitle = "Wrong UNIQUE";
                validation_7H.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7H.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7H.ShowErrorMessage = true;

                validation_7I.ErrorTitle = "Wrong LOCAL";
                validation_7I.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_7I.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7I.ShowErrorMessage = true;

                validation_7J.ErrorTitle = "Wrong SORT_ORDER";
                validation_7J.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_7J.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_7J.ShowErrorMessage = true;

                //worksheet7.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet7["$A:$XFD"].Protection.Locked = false;
                //worksheet7["A1:J1"].Protection.Locked = true;

                //O_MOD_ASGN_LVL_FLAG_SETTING
                var worksheet8 = workbook.Worksheets["O_MOD_ASGN_LVL_FLAG_SETTING"];
                DataValidation validation_8A = worksheet8.DataValidations.Add(worksheet8["A2:A1048576"], DataValidationType.Custom, "=AND(ISNUMBER(A2:A1048576),LEN(A2:A1048576)<=10)");
                DataValidation validation_8B = worksheet8.DataValidations.Add(worksheet8["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=10)");
                DataValidation validation_8C = worksheet8.DataValidations.Add(worksheet8["C2:C1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_8D = worksheet8.DataValidations.Add(worksheet8["D2:D1048576"], DataValidationType.List, "Y, N");

                validation_8A.ErrorTitle = "Wrong MODULEID";
                validation_8A.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_8A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_8A.ShowErrorMessage = true;

                validation_8B.ErrorTitle = "Wrong ATTRNO";
                validation_8B.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_8B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_8B.ShowErrorMessage = true;

                validation_8C.ErrorTitle = "Wrong FIELD_COLL_FLAG";
                validation_8C.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_8C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_8C.ShowErrorMessage = true;

                validation_8D.ErrorTitle = "Wrong MANDATORY_FLAG";
                validation_8D.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_8D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_8D.ShowErrorMessage = true;

                //worksheet8.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet8["$A:$XFD"].Protection.Locked = false;
                //worksheet8["A1:D1"].Protection.Locked = true;

                //O_HEADING_TYPE
                var worksheet9 = workbook.Worksheets["O_HEADING_TYPE"];
                DataValidation validation_9A = worksheet9.DataValidations.Add(worksheet9["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);
                DataValidation validation_9B = worksheet9.DataValidations.Add(worksheet9["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 50);
                DataValidation validation_9C = worksheet9.DataValidations.Add(worksheet9["C2:C1048576"], DataValidationType.Custom, "=AND(ISNUMBER(C2:C1048576),LEN(C2:C1048576)<=10)");
                DataValidation validation_9D = worksheet9.DataValidations.Add(worksheet9["D2:D1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_9E = worksheet9.DataValidations.Add(worksheet9["E2:E1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_9F = worksheet9.DataValidations.Add(worksheet9["F2:F1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_9G = worksheet9.DataValidations.Add(worksheet9["G2:G1048576"], DataValidationType.List, "Y, N");
                DataValidation validation_9H = worksheet9.DataValidations.Add(worksheet9["H2:H1048576"], DataValidationType.TextLength, DataValidationOperator.Equal, 1); //M L
                DataValidation validation_9I = worksheet9.DataValidations.Add(worksheet9["I2:I1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 0, 30);
                DataValidation validation_9J = worksheet9.DataValidations.Add(worksheet9["J2:J1048576"], DataValidationType.List, "Y, N");

                validation_9A.ErrorTitle = "Wrong HEADING_ID";
                validation_9A.ErrorMessage = "The value you entered is not valid. Use 1-30 Alphabet/Numeric characters value.";
                validation_9A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9A.ShowErrorMessage = true;

                validation_9B.ErrorTitle = "Wrong HEADING_NAME";
                validation_9B.ErrorMessage = "The value you entered is not valid. Use 1-50 Alphabet/Numeric characters value.";
                validation_9B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9B.ShowErrorMessage = true;

                validation_9C.ErrorTitle = "Wrong MAXIMUM_LENGTH";
                validation_9C.ErrorMessage = "The value you entered is not valid. Use 1-10 digit Numeric value.";
                validation_9C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9C.ShowErrorMessage = true;

                validation_9D.ErrorTitle = "Wrong ONLINE_HDNG_FLAG";
                validation_9D.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9D.ShowErrorMessage = true;

                validation_9E.ErrorTitle = "Wrong OUTPUT_FLAG";
                validation_9E.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9E.ShowErrorMessage = true;

                validation_9F.ErrorTitle = "Wrong ACTIVE_FLAG";
                validation_9F.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9F.ShowErrorMessage = true;

                validation_9G.ErrorTitle = "Wrong ALTERNATE_HEADING_FLAG";
                validation_9G.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9G.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9G.ShowErrorMessage = true;

                validation_9H.ErrorTitle = "Wrong ASGN_LVL";
                validation_9H.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9H.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9H.ShowErrorMessage = true;

                //null
                validation_9I.ErrorTitle = "Wrong HDNG_TYP";
                validation_9I.ErrorMessage = "The value you entered is not valid. Use upto 20 Alphabet/Numeric characters value.";
                validation_9I.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9I.ShowErrorMessage = true;

                validation_9J.ErrorTitle = "Wrong LOCALLANG";
                validation_9J.ErrorMessage = "The value you entered is not valid. Use single character Alphabet value.";
                validation_9J.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_9J.ShowErrorMessage = true;

                //worksheet9.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet9["$A:$XFD"].Protection.Locked = false;
                //worksheet9["A1:J1"].Protection.Locked = true;

                //O_HEADING_PC_SETTING
                var worksheet10 = workbook.Worksheets["O_HEADING_PC_SETTING"];
                DataValidation validation_10A = worksheet10.DataValidations.Add(worksheet10["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 50);
                DataValidation validation_10B = worksheet10.DataValidations.Add(worksheet10["B2:B1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_10C = worksheet10.DataValidations.Add(worksheet10["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_10D = worksheet10.DataValidations.Add(worksheet10["D2:D1048576"], DataValidationType.Custom, "=AND(ISNUMBER(D2:D1048576),LEN(D2:D1048576)<=5)");
                DataValidation validation_10E = worksheet10.DataValidations.Add(worksheet10["E2:E1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 0, 10);
                DataValidation validation_10F = worksheet10.DataValidations.Add(worksheet10["F2:F1048576"], DataValidationType.Custom, "=AND(ISNUMBER(F2:F1048576),LEN(F2:F1048576)<=5)");

                validation_10A.ErrorTitle = "Wrong HEADING_NAME";
                validation_10A.ErrorMessage = "The value you entered is not valid. Use 1-50 Alphabet/Numeric characters value.";
                validation_10A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10A.ShowErrorMessage = true;

                validation_10B.ErrorTitle = "Wrong CATEGORYATTRVALUE";
                validation_10B.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_10B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10B.ShowErrorMessage = true;

                validation_10C.ErrorTitle = "Wrong CHAR_ID";
                validation_10C.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_10C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10C.ShowErrorMessage = true;

                validation_10D.ErrorTitle = "Wrong SEQUENCE";
                validation_10D.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_10D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10D.ShowErrorMessage = true;

                //null
                validation_10E.ErrorTitle = "Wrong CONCATENATOR";
                validation_10E.ErrorMessage = "The value you entered is not valid. Use upto 10 Alphabet/Numeric characters value.";
                validation_10E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10E.ShowErrorMessage = true;

                validation_10F.ErrorTitle = "Wrong DROP_PRIORITY";
                validation_10F.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_10F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_10F.ShowErrorMessage = true;

                //worksheet10.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet10["$A:$XFD"].Protection.Locked = false;
                //worksheet10["A1:F1"].Protection.Locked = true;

                //O_HEADING_MOD_SETTING
                var worksheet11 = workbook.Worksheets["O_HEADING_MOD_SETTING"];
                DataValidation validation_11A = worksheet11.DataValidations.Add(worksheet11["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 50);
                DataValidation validation_11B = worksheet11.DataValidations.Add(worksheet11["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=10)");
                DataValidation validation_11C = worksheet11.DataValidations.Add(worksheet11["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_11D = worksheet11.DataValidations.Add(worksheet11["D2:D1048576"], DataValidationType.Custom, "=AND(ISNUMBER(D2:D1048576),LEN(D2:D1048576)<=5)");
                DataValidation validation_11E = worksheet11.DataValidations.Add(worksheet11["E2:E1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 0, 10);
                DataValidation validation_11F = worksheet11.DataValidations.Add(worksheet11["F2:F1048576"], DataValidationType.Custom, "=AND(ISNUMBER(F2:F1048576),LEN(F2:F1048576)<=5)");

                validation_11A.ErrorTitle = "Wrong HEADING_NAME";
                validation_11A.ErrorMessage = "The value you entered is not valid. Use 1-50 Alphabet/Numeric characters value.";
                validation_11A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11A.ShowErrorMessage = true;

                validation_11B.ErrorTitle = "Wrong MODULEID";
                validation_11B.ErrorMessage = "The value you entered is not valid. Use 1-10 Alphabet/Numeric characters value.";
                validation_11B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11B.ShowErrorMessage = true;

                validation_11C.ErrorTitle = "Wrong CHAR_ID";
                validation_11C.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_11C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11C.ShowErrorMessage = true;

                validation_11D.ErrorTitle = "Wrong SEQUENCE";
                validation_11D.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_11D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11D.ShowErrorMessage = true;

                //null
                validation_11E.ErrorTitle = "Wrong CONCATENATOR";
                validation_11E.ErrorMessage = "The value you entered is not valid. Use upto 10 Alphabet/Numeric characters value.";
                validation_11E.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11E.ShowErrorMessage = true;

                validation_11F.ErrorTitle = "Wrong DROP_PRIORITY";
                validation_11F.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_11F.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_11F.ShowErrorMessage = true;

                //worksheet11.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet11["$A:$XFD"].Protection.Locked = false;
                //worksheet11["A1:F1"].Protection.Locked = true;

                var worksheet12 = workbook.Worksheets["O_RETAILER_DEPT_SUPP"];
                DataValidation validation_12A = worksheet12.DataValidations.Add(worksheet12["A2:A1048576"], DataValidationType.Custom, "=AND(ISNUMBER(A2:A1048576),LEN(A2:A1048576)<=15)");
                DataValidation validation_12B = worksheet12.DataValidations.Add(worksheet12["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=15)");
                DataValidation validation_12C = worksheet12.DataValidations.Add(worksheet12["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);
                DataValidation validation_12D = worksheet12.DataValidations.Add(worksheet12["D2:D1048576"], DataValidationType.Custom, "=AND(ISNUMBER(D2:D1048576),LEN(D2:D1048576)<=5)");

                validation_12A.ErrorTitle = "Wrong SOURCEID";
                validation_12A.ErrorMessage = "The value you entered is not valid. Use 1-15 digit Numeric value.";
                validation_12A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_12A.ShowErrorMessage = true;

                validation_12B.ErrorTitle = "Wrong ATTRNO";
                validation_12B.ErrorMessage = "The value you entered is not valid.  Use 1-15 digit Numeric value.";
                validation_12B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_12B.ShowErrorMessage = true;

                validation_12C.ErrorTitle = "Wrong TYPE";
                validation_12C.ErrorMessage = "The value you entered is not valid. Use 1-30 Alphabet/Numeric characters value.";
                validation_12C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_12C.ShowErrorMessage = true;

                validation_12D.ErrorTitle = "Wrong SEQ";
                validation_12D.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_12D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_12D.ShowErrorMessage = true;

                //worksheet12.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet12["$A:$XFD"].Protection.Locked = false;
                //worksheet12["A1:D1"].Protection.Locked = true;

                //O_UOM_MAPPING_LIST
                var worksheet13 = workbook.Worksheets["O_UOM_MAPPING_LIST"];
                DataValidation validation_13A = worksheet13.DataValidations.Add(worksheet13["A2:A1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20); ;
                DataValidation validation_13B = worksheet13.DataValidations.Add(worksheet13["B2:B1048576"], DataValidationType.Custom, "=AND(ISNUMBER(B2:B1048576),LEN(B2:B1048576)<=5)");
                DataValidation validation_13C = worksheet13.DataValidations.Add(worksheet13["C2:C1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 20);
                DataValidation validation_13D = worksheet13.DataValidations.Add(worksheet13["D2:D1048576"], DataValidationType.TextLength, DataValidationOperator.Between, 1, 30);

                validation_13A.ErrorTitle = "Wrong RF_UOMCODE";
                validation_13A.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_13A.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_13A.ShowErrorMessage = true;

                validation_13B.ErrorTitle = "Wrong OGRDS_UOMID";
                validation_13B.ErrorMessage = "The value you entered is not valid. Use 1-5 digit Numeric value.";
                validation_13B.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_13B.ShowErrorMessage = true;

                validation_13C.ErrorTitle = "Wrong OGRDS_UOMDESC";
                validation_13C.ErrorMessage = "The value you entered is not valid. Use 1-20 Alphabet/Numeric characters value.";
                validation_13C.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_13C.ShowErrorMessage = true;

                validation_13D.ErrorTitle = "Wrong PREFERED_UNIT_DSCR";
                validation_13D.ErrorMessage = "The value you entered is not valid. Use 1-30 Alphabet/Numeric characters value.";
                validation_13D.ErrorStyle = DataValidationErrorStyle.Stop;
                validation_13D.ShowErrorMessage = true;

                //worksheet13.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
                //worksheet13["$A:$XFD"].Protection.Locked = false;
                //worksheet13["A1:D1"].Protection.Locked = true;
                GC.Collect();
            }
        }
        private void CellSpecialCharValidation()
        {
            RemoveSingleCellColor(); //BGRF-2086
            IWorkbook workbook = spreadsheetControl1.Document;
            var worksheet = workbook.Worksheets.ActiveWorksheet;

            var columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;
            var columnValue = spreadsheetControl1.ActiveCell.Value.TextValue;

            if (columnValue != null && !string.IsNullOrEmpty(columnValue.Trim()))
            {
                if (!Regex.IsMatch(columnValue, "^[a-zA-Z][a-zA-Z0-9-_ ]*$"))
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (char c in columnValue)
                    {
                        if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '-' || c == '_' || c == ' ')
                        {
                            sb.Append(c);
                        }
                    }
                    // spreadsheetControl1.ActiveCell.Value = sb.ToString(); //BGRF-2086
                    // MessageBox.Show("Value can not contain Special Characters", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); //BGRF-2086
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(columnValue))
                {
                    CellValueNotNull(worksheet.Name, columnIndex);
                }
            }


            //var worksheet1 = workbook.Worksheets["O_ACTIVE_CATEGORY"];
            //worksheet1["$a:$xfd"].Protection.Locked = false;
            //worksheet1["a1:c1"].Protection.Locked = true;
            //worksheet1.Protect("password", DevExpress.Spreadsheet.WorksheetProtectionPermissions.Default);
        }
        protected bool IsCellValid(DataValidation dataValidation, Cell cell)
        {
            switch (dataValidation.ValidationType)
            {
                case DataValidationType.TextLength:
                    {
                        if ((dataValidation.Operator == DataValidationOperator.LessThanOrEqual) && !cell.Value.IsEmpty)
                        {
                            if (cell.Value.IsBoolean)
                                return IsTextIsValid(cell.Value.BooleanValue.ToString(), dataValidation.Criteria.NumericValue);
                            if (cell.Value.IsDateTime)
                                return IsTextIsValid(cell.Value.DateTimeValue.ToString(), dataValidation.Criteria.NumericValue);
                            if (cell.Value.IsError)
                                return IsTextIsValid(cell.Value.ErrorValue.ToString(), dataValidation.Criteria.NumericValue);
                            if (cell.Value.IsNumeric)
                                return IsTextIsValid(cell.Value.NumericValue.ToString(), dataValidation.Criteria.NumericValue);
                            if (cell.Value.IsText)
                                return IsTextIsValid(cell.Value.TextValue, dataValidation.Criteria.NumericValue);
                        }
                    }
                    break;
                case DataValidationType.AnyValue:
                    {
                        if (dataValidation.Operator == DataValidationOperator.NotEqual)
                            return !cell.Value.IsEmpty;
                    }
                    break;
            }
            return false;
        }
        protected bool IsTextIsValid(string text, double maxStringLength)
        {
            return (text.Length <= maxStringLength) ? true : false;
        }
        private void CellValueNotNull(string worksheet, int columnIndex)
        {
            if (worksheet == "O_ACTIVE_CATEGORY")
            {
                if (columnIndex == 0) { MessageBox.Show("CATEGORYATTRVALUE can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("SERVICE can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("TYPE can not be Empty,Use 1-10 Alphabet/Numeric characters value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            if (worksheet == "O_PRE_SET_CHAR")
            {
                if (columnIndex == 0) { MessageBox.Show("TYPE can not be Empty, Use 1-15 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("ATTRNO can not be Empty, Use 1-15 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("REQUIRED can not be Empty,Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            if (worksheet == "O_PROC_GRP_SET")
            {
                if (columnIndex == 0) { MessageBox.Show("PROC_GROUP_SET_ID can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("PROC_GROUP_SET_NAME can not be Empty, Use 1-20 Alphabet/Numeric characters value., ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("FOLLOW_EXT_RULE can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("EXCP_SENT_TO_XCD_BRWSR can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 4) { MessageBox.Show("EXCP_SENT_TO_UNCDBLE_ITM can not be Empty, Use single character Alphabet value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            if (worksheet == "O_EXT_CODE_GRP")
            {
                if (columnIndex == 0) { MessageBox.Show("SHORTNAME can not be Empty, Use 1-10 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("EXTERNAL_CODE_GROUP_NAME can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            if (worksheet == "O_ACTIVE_RETAILER")
            {
                if (columnIndex == 0) { MessageBox.Show("SOURCEID can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("PROC_GROUP_SET_NAME can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("TYPE can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("EAN can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 4) { MessageBox.Show("UPC can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("LAC can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 6) { MessageBox.Show("CIP can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_DIC_LVL_FLAG_SETTING")
            {
                if (columnIndex == 0) { MessageBox.Show("ATTRNO can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 1) { MessageBox.Show("OGRDSID can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("MAP_TO_OGRDS_CHR_VAL can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("LONGDESC can not be Empty, Use 1-70 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 4) { MessageBox.Show("OTHERDESC can not be Empty, Use 1-70 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("ALTER30MAX can not be Empty, Use 1-70 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 6) { MessageBox.Show("ALTER30MAX_OTHER can not be Empty, Use 1-70 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 7) { MessageBox.Show("ALTER5MAX can not be Empty, Use 1-5 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 8) { MessageBox.Show("CATEGORY_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 9) { MessageBox.Show("CHAR_TYPE can not be Empty,  Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 10) { MessageBox.Show("NUMERIC_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 11) { MessageBox.Show("CUSTOMER_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 12) { MessageBox.Show("FIXED_ITEM_VALUE can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 13) { MessageBox.Show("COPY_ITEM can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 14) { MessageBox.Show("MULTI_VALUE can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 15) { MessageBox.Show("ABBREVIATE_VALUE can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 16) { MessageBox.Show("FIXED_VALUE_LIST can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 17) { MessageBox.Show("TRANSLATION_IND can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 18) { MessageBox.Show("CHR_VAL_DESCRIPTION_ONLY can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 19) { MessageBox.Show("LOCAL can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 20) { MessageBox.Show("SORT_ORDER can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            if (worksheet == "O_PC_ASGN_LVL_FLAG_SETTING")
            {
                if (columnIndex == 0) { MessageBox.Show("CATEGORYATTRVALUE can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("ATTRNO can not be Empty, Use 1-10 digit Numeric value..", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("FIXED_VALUE_LIST can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("COPY_ITEM_VAL can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 4) { MessageBox.Show("FIELD_COLL_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("MANDATORY_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 6) { MessageBox.Show("ALIGNED can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 7) { MessageBox.Show("UNIQUE can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 8) { MessageBox.Show("LOCAL can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 9) { MessageBox.Show("SORT_ORDER can not be Empty,  Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_MOD_ASGN_LVL_FLAG_SETTING")
            {
                if (columnIndex == 0) { MessageBox.Show("MODULEID can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("ATTRNO can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("FIELD_COLL_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("MANDATORY_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_HEADING_TYPE")
            {
                if (columnIndex == 0) { MessageBox.Show("HEADING_ID can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("HEADING_NAME can not be Empty, Use 1-50 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("MAXIMUM_LENGTH can not be Empty, Use 1-10 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("ONLINE_HDNG_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 4) { MessageBox.Show("OUTPUT_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("ACTIVE_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 6) { MessageBox.Show("ALTERNATE_HEADING_FLAG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 7) { MessageBox.Show("ASGN_LVL can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                // if (columnIndex == 8) { MessageBox.Show("HDNG_TYP can not be Empty, Use upto 20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 9) { MessageBox.Show("LOCALLANG can not be Empty, Use single character Alphabet value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_HEADING_PC_SETTING")
            {
                if (columnIndex == 0) { MessageBox.Show("HEADING_NAME can not be Empty, Use 1-50 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("CATEGORYATTRVALUE can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("CHAR_ID can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("SEQUENCE can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 4) { MessageBox.Show("CONCATENATOR can not be Empty, Use upto 10 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("DROP_PRIORITY can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_HEADING_MOD_SETTING")
            {
                if (columnIndex == 0) { MessageBox.Show("HEADING_NAME can not be Empty, Use 1-50 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("MODULEID can not be Empty, Use 1-10 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("CHAR_ID can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("SEQUENCE can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                //if (columnIndex == 4) { MessageBox.Show("CONCATENATOR can not be Empty, Use upto 10 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 5) { MessageBox.Show("DROP_PRIORITY can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_RETAILER_DEPT_SUPP")
            {
                if (columnIndex == 0) { MessageBox.Show("SOURCEID can not be Empty, Use 1-15 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("ATTRNO can not be Empty, Use 1-15 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("TYPE can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("SEQ can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            if (worksheet == "O_UOM_MAPPING_LIST")
            {
                if (columnIndex == 0) { MessageBox.Show("RF_UOMCODE can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 1) { MessageBox.Show("OGRDS_UOMID can not be Empty, Use 1-5 digit Numeric value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 2) { MessageBox.Show("OGRDS_UOMDESC can not be Empty, Use 1-20 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (columnIndex == 3) { MessageBox.Show("PREFERED_UNIT_DSCR can not be Empty, Use 1-30 Alphabet/Numeric characters value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }
        private void ErrorLengthCell(string tableName, DataTable dt)
        {
            string errorLengthMsg = string.Empty;
            ProcessTool pt = new ProcessTool();
            IWorkbook workbook = spreadsheetControl1.Document;
            var worksheet = workbook.Worksheets[tableName];

            if (tableName == "O_ACTIVE_CATEGORY")
            {
                bool error = false;
                bool CATEGORYATTRVALUE = false;
                bool SERVICE = false;
                bool TYPE = false;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(dt.Rows[i]["CATEGORYATTRVALUE"].ToString()))
                        {
                            if (dt.Rows[i]["CATEGORYATTRVALUE"].ToString().Length > 20)
                            {
                                if (!CATEGORYATTRVALUE)
                                {
                                    error = true;
                                    CATEGORYATTRVALUE = true;
                                    errorLengthMsg += "CATEGORYATTRVALUE length is greater than Max lengh 20\n";
                                }
                                worksheet.Cells[i + 1, j].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j].FillColor = Color.Empty; //BGRF-2086
                            }
                        }
                        else if (!string.IsNullOrEmpty(dt.Rows[i]["SERVICE"].ToString()))
                        {
                            if (dt.Rows[i]["SERVICE"].ToString().Length > 20)
                            {
                                if (!SERVICE)
                                {
                                    error = true;
                                    SERVICE = true;
                                    errorLengthMsg += "SERVICE length is greater than Max lengh 20\n";
                                }
                                worksheet.Cells[i + 1, j].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j].FillColor = Color.Empty; //BGRF-2086
                            }
                        }
                        else if (!string.IsNullOrEmpty(dt.Rows[i]["TYPE"].ToString()))
                        {
                            if (dt.Rows[i]["TYPE"].ToString().Length > 10)
                            {
                                if (!TYPE)
                                {
                                    error = true;
                                    TYPE = true;
                                    errorLengthMsg += "TYPE length is greater than Max lengh 10\n";
                                }
                                worksheet.Cells[i + 1, j].FillColor = Color.Red; //BGRF-2086
                            }
                            else
                            {
                                worksheet.Cells[i + 1, j].FillColor = Color.Empty; //BGRF-2086
                            }
                        }
                    }
                }
                if (error)
                {
                    errorRecord += "\n";
                    errorRecord += "Error tab : O_ACTIVE_CATEGORY\n";
                    errorRecord += errorLengthMsg;
                }
            }


        }
        private void RemoveEmptyRows(DataTable source)
        {
            for (int i = source.Rows.Count; i >= 1; i--)
            {
                bool deleteRow = true;
                DataRow currentRow = source.Rows[i - 1];
                foreach (var colValue in currentRow.ItemArray)
                {
                    if (!string.IsNullOrEmpty(colValue.ToString().Trim()))
                    {
                        deleteRow = false;
                        break;
                    }
                }
                if (deleteRow)
                {
                    source.Rows[i - 1].Delete();
                }
            }
            source.AcceptChanges();
        }

        // start BGRF-2086
        private void spreadsheetControl1_CellBeginEdit(object sender, DevExpress.XtraSpreadsheet.SpreadsheetCellCancelEventArgs e)
        {
            CellDataValidation();//BGRF-2051
            if (e.RowIndex == 0)
            {
                e.Cancel = true;
                MessageBox.Show("Headers are non-editable", "Error", buttons, MessageBoxIcon.Error);
                //spreadsheetControl1.Options.Behavior.Paste = DocumentCapability.Disabled;
            }
        }
        private void spreadsheetControl1_CellEndEdit(object sender, SpreadsheetCellValidatingEventArgs e)
        {
            if (string.IsNullOrEmpty(e.EditorText.Trim()))
            {
                errorCell = false;
                e.Cancel = true;
            }
            NumericCell();//BGRF-2086
        }
        private void RemoveSingleCellColor()//int table, int rowIndex, int columnIndex)
        {
            Color colReset = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
            IWorkbook workbook = spreadsheetControl1.Document;
            var worksheet = workbook.Worksheets.ActiveWorksheet;
            var rowIndex = spreadsheetControl1.ActiveCell.RowIndex;
            var columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;
            worksheet.Cells[rowIndex, columnIndex].FillColor = Color.Empty;
        }
        private void ColumnWidth()
        {
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                IWorkbook workbook = spreadsheetControl1.Document;
                var worksheet = workbook.Worksheets[ds.Tables[i].TableName.ToString()];
                worksheet.DefaultColumnWidthInCharacters = 20;
            }
        }
        private void RemoveColor(int table)
        {
            for (int i = 0; i < ds.Tables[table].Rows.Count; i++)
            {
                try
                {
                    Color colReset = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
                    IWorkbook workbook = spreadsheetControl1.Document;
                    var worksheet = workbook.Worksheets[ds.Tables[table].TableName];
                    for (int col = 0; col < ds.Tables[table].Columns.Count; col++)
                    {
                        worksheet.Cells[i + 1, col].FillColor = Color.Empty;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        private void NumericCell()
        {
            //start BGRF-2086 : added for numeric cell
            IWorkbook workbook = spreadsheetControl1.Document;
            var worksheet = workbook.Worksheets.ActiveWorksheet;
            var rowIndex = spreadsheetControl1.ActiveCell.RowIndex + 1;
            var columnIndex = spreadsheetControl1.ActiveCell.ColumnIndex;
            string cellName = spreadsheetControl1.ActiveCell.ToString();
            cellName = cellName.Substring(5, 1);
            string cell = string.Format("{0}{1}", cellName, rowIndex);

            if (spreadsheetControl1.ActiveWorksheet[cell].Value.IsNumeric)
            {
                spreadsheetControl1.ActiveWorksheet[cell].NumberFormat = "##";
                spreadsheetControl1.ActiveWorksheet[cell].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
            }
            //end BGRF-2086
        }
        private void HighlightErrorSheet()
        {
            bool activeTab = false;
            IWorkbook workbook = spreadsheetControl1.Document;
            if (activeTab == false && errorRecord.Contains("O_ACTIVE_CATEGORY"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_ACTIVE_CATEGORY"];
            }
            if (activeTab == false && errorRecord.Contains("O_PRE_SET_CHAR"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_PRE_SET_CHAR"];
            }
            if (activeTab == false && errorRecord.Contains("O_PROC_GRP_SET"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_PROC_GRP_SET"];
            }
            if (activeTab == false && errorRecord.Contains("O_EXT_CODE_GRP"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_EXT_CODE_GRP"];
            }
            if (activeTab == false && errorRecord.Contains("O_ACTIVE_RETAILER"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_ACTIVE_RETAILER"];
            }
            if (activeTab == false && errorRecord.Contains("O_DIC_LVL_FLAG_SETTING"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_DIC_LVL_FLAG_SETTING"];
            }
            if (activeTab == false && errorRecord.Contains("O_PC_ASGN_LVL_FLAG_SETTING"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_PC_ASGN_LVL_FLAG_SETTING"];
            }
            if (activeTab == false && errorRecord.Contains("O_MOD_ASGN_LVL_FLAG_SETTING"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_MOD_ASGN_LVL_FLAG_SETTING"];
            }
            if (activeTab == false && errorRecord.Contains("O_HEADING_TYPE"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_HEADING_TYPE"];
            }
            if (activeTab == false && errorRecord.Contains("O_HEADING_PC_SETTING"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_HEADING_PC_SETTING"];
            }
            if (activeTab == false && errorRecord.Contains("O_HEADING_MOD_SETTING"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_HEADING_MOD_SETTING"];
            }
            if (activeTab == false && errorRecord.Contains("O_RETAILER_DEPT_SUPP"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_RETAILER_DEPT_SUPP"];
            }
            if (activeTab == false && errorRecord.Contains("O_UOM_MAPPING_LIST"))
            {
                activeTab = true;
                workbook.Worksheets.ActiveWorksheet = workbook.Worksheets["O_UOM_MAPPING_LIST"];
            }
        }
        // end BGRF-2086
        #endregion

        private void btnHelp_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            HelpForm d = new HelpForm();
            d.Show();
        }
        //Sudheer - BGRF-2087
        private void btnExport_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.InitialDirectory = @"C:\";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.CheckPathExists = true;
                saveFileDialog.Title = "Export";
                saveFileDialog.Filter = "Data File (*.xlsx)|*.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        spreadsheetControl1.SaveDocument(stream, DocumentFormat.Xlsx);
                    }
                    MessageBox.Show("Export File Successfully", "Info", buttons, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export File Failed", "Error", buttons, MessageBoxIcon.Error);
            }
        }
        //Added for BGRF-2086


    }
}
