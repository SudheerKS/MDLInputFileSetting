using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;
using System.IO;
using System.Data;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Xml;

namespace MDL
{
    #region Step1:GroupforCodetypeComputation
    //start added for codetype Computation --SIVA BGRF-1795
    class Group
    {
        public int GroupNo;
        public string TableName;
    }
    //end added for codetype Computation --SIVA BGRF-1795
    #endregion Step1:GroupforCodetypeComputation

    class BusinessLayer
    {
        private static string _strConn = "";
        private string Sqlquery = string.Empty;
        public BusinessLayer()
        {
            _strConn = "Data Source=" + Common.strDBName +
                      ";Persist Security Info=True;User ID=" + Common.strUserName +
                      ";Password=" + Common.strPassword;
        }

        //BGRF-2087 Sudheer
        #region 3. Load MDL Input Setting File
        public void MDLInputSettings_Load()
        {
            ProcessTool.dtbActiveCategory.Clear();
            DataLayer objDataLayer1 = new DataLayer(_strConn);
            OracleCommand ocdSelect1 = new OracleCommand();
            StringBuilder sql1 = new StringBuilder();
            sql1.Append("SELECT * FROM  O_ACTIVE_CATEGORY order by categoryattrvalue");
            ocdSelect1.CommandText = sql1.ToString();
            ocdSelect1.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbActiveCategory.Merge(objDataLayer1.GetResultDT(ocdSelect1));


            DataLayer objDataLayer2 = new DataLayer(_strConn);
            OracleCommand ocdSelect2 = new OracleCommand();
            StringBuilder sql2 = new StringBuilder();
            //sql2.Append("SELECT * FROM O_PRE_SET_CHAR");
            sql2.Append("select TYPE ,TO_CHAR(ATTRNO) ATTRNO,REQUIRED from O_PRE_SET_CHAR order by type");
            ocdSelect2.CommandText = sql2.ToString();
            ocdSelect2.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbPreSetChar.Merge(objDataLayer2.GetResultDT(ocdSelect2));

            DataLayer objDataLayer3 = new DataLayer(_strConn);
            OracleCommand ocdSelect3 = new OracleCommand();
            StringBuilder sql3 = new StringBuilder();
            sql3.Append("SELECT * FROM O_PROC_GRP_SET  order by proc_group_set_id");
            ocdSelect3.CommandText = sql3.ToString();
            ocdSelect3.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbProcGrpSet.Merge(objDataLayer3.GetResultDT(ocdSelect3));

            DataLayer objDataLayer4 = new DataLayer(_strConn);
            OracleCommand ocdSelect4 = new OracleCommand();
            StringBuilder sql4 = new StringBuilder();
            sql4.Append("SELECT * FROM O_EXT_CODE_GRP  order by shortname");
            ocdSelect4.CommandText = sql4.ToString();
            ocdSelect4.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbExtCodeGrp.Merge(objDataLayer4.GetResultDT(ocdSelect4));

            DataLayer objDataLayer5 = new DataLayer(_strConn);
            OracleCommand ocdSelect5 = new OracleCommand();
            StringBuilder sql5 = new StringBuilder();
            //sql5.Append("SELECT * FROM O_ACTIVE_RETAILER");
            sql5.Append("select TO_CHAR(SOURCEID) SOURCEID,PROC_GROUP_SET_NAME,TYPE,EAN, UPC,LAC,CIP   from   O_ACTIVE_RETAILER order by O_ACTIVE_RETAILER.SOURCEID");

            ocdSelect5.CommandText = sql5.ToString();
            ocdSelect5.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbActiveRetailer.Merge(objDataLayer5.GetResultDT(ocdSelect5));

            DataLayer objDataLayer6 = new DataLayer(_strConn);
            OracleCommand ocdSelect6 = new OracleCommand();
            StringBuilder sql6 = new StringBuilder();
            //sql6.Append("SELECT * FROM O_DIC_LVL_FLAG_SETTING");

            sql6.Append("select  TO_CHAR(ATTRNO) ATTRNO,TO_CHAR(OGRDSID) OGRDSID,MAP_TO_OGRDS_CHR_VAL,LONGDESC,OTHERDESC,ALTER30MAX,ALTER30MAX_OTHER,ALTER5MAX,CATEGORY_FLAG ," +
                         @"CHAR_TYPE ,NUMERIC_FLAG,CUSTOMER_FLAG,FIXED_ITEM_VALUE,COPY_ITEM,MULTI_VALUE,ABBREVIATE_VALUE,FIXED_VALUE_LIST," +
                         @"TRANSLATION_IND,CHR_VAL_DESCRIPTION_ONLY,LOCAL,TO_CHAR(SORT_ORDER) SORT_ORDER from O_DIC_LVL_FLAG_SETTING order by O_DIC_LVL_FLAG_SETTING.ATTRNO");
            ocdSelect6.CommandText = sql6.ToString();
            ocdSelect6.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbDicLvlFlagSetting.Merge(objDataLayer6.GetResultDT(ocdSelect6));

            DataLayer objDataLayer7 = new DataLayer(_strConn);
            OracleCommand ocdSelect7 = new OracleCommand();
            StringBuilder sql7 = new StringBuilder();
           // sql7.Append("SELECT * FROM O_PC_ASGN_LVL_FLAG_SETTING");
            sql7.Append("select CATEGORYATTRVALUE,TO_CHAR(ATTRNO) ATTRNO,FIXED_VALUE_LIST,COPY_ITEM_VAL,FIELD_COLL_FLAG,MANDATORY_FLAG,ALIGNED, \"UNIQUE\" ,LOCAL,TO_CHAR(SORT_ORDER) SORT_ORDER  from O_PC_ASGN_LVL_FLAG_SETTING order by ATTRNO");

            ocdSelect7.CommandText = sql7.ToString();
            ocdSelect7.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbPCAsgnlvlFlagSetting.Merge(objDataLayer7.GetResultDT(ocdSelect7));

            DataLayer objDataLayer8 = new DataLayer(_strConn);
            OracleCommand ocdSelect8 = new OracleCommand();
            StringBuilder sql8 = new StringBuilder();
            //sql8.Append("SELECT * FROM O_MOD_ASGN_LVL_FLAG_SETTING");
            sql8.Append("select TO_CHAR(MODULEID) MODULEID,TO_CHAR(ATTRNO) ATTRNO,FIELD_COLL_FLAG,MANDATORY_FLAG  from O_MOD_ASGN_LVL_FLAG_SETTING order by O_MOD_ASGN_LVL_FLAG_SETTING.ATTRNO");

            ocdSelect8.CommandText = sql8.ToString();
            ocdSelect8.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbModAsgnlvlFlagSetting.Merge(objDataLayer8.GetResultDT(ocdSelect8));

            DataLayer objDataLayer9 = new DataLayer(_strConn);
            OracleCommand ocdSelect9 = new OracleCommand();
            StringBuilder sql9 = new StringBuilder();
           // sql9.Append("SELECT * FROM O_HEADING_TYPE");
            sql9.Append("select HEADING_ID,HEADING_NAME,TO_CHAR(MAXIMUM_LENGTH) MAXIMUM_LENGTH,ONLINE_HDNG_FLAG,OUTPUT_FLAG,ACTIVE_FLAG,ALTERNATE_HEADING_FLAG ,ASGN_LVL,HDNG_TYP,LOCALLANG from O_HEADING_TYPE order by HEADING_ID");

            ocdSelect9.CommandText = sql9.ToString();
            ocdSelect9.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbHeadingType.Merge(objDataLayer9.GetResultDT(ocdSelect9));

            DataLayer objDataLayer10 = new DataLayer(_strConn);
            OracleCommand ocdSelect10 = new OracleCommand();
            StringBuilder sql10 = new StringBuilder();
            //sql10.Append("SELECT * FROM O_HEADING_PC_SETTING");
            sql10.Append("select HEADING_NAME,CATEGORYATTRVALUE,CHAR_ID,TO_CHAR(SEQUENCE) SEQUENCE,CONCATENATOR,TO_CHAR(DROP_PRIORITY) DROP_PRIORITY from O_HEADING_PC_SETTING order by HEADING_NAME");

            ocdSelect10.CommandText = sql10.ToString();
            ocdSelect10.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbHeadingPCSetting.Merge(objDataLayer10.GetResultDT(ocdSelect10));

            DataLayer objDataLayer11 = new DataLayer(_strConn);
            OracleCommand ocdSelect11 = new OracleCommand();
            StringBuilder sql11 = new StringBuilder();
            //sql11.Append("SELECT * FROM O_HEADING_MOD_SETTING");
            sql11.Append("select HEADING_NAME,TO_CHAR(MODULEID) MODULEID,CHAR_ID,TO_CHAR(SEQUENCE) SEQUENCE,CONCATENATOR,TO_CHAR(DROP_PRIORITY) DROP_PRIORITY  from O_HEADING_MOD_SETTING order by HEADING_NAME");

            ocdSelect11.CommandText = sql11.ToString();
            ocdSelect11.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbHeadingModSetting.Merge(objDataLayer11.GetResultDT(ocdSelect11));

            DataLayer objDataLayer12 = new DataLayer(_strConn);
            OracleCommand ocdSelect12 = new OracleCommand();
            StringBuilder sql12 = new StringBuilder();
            //sql12.Append("SELECT * FROM O_RETAILER_DEPT_SUPP");
            sql12.Append("select TO_CHAR(SOURCEID) SOURCEID,TO_CHAR(ATTRNO) ATTRNO,TYPE,TO_CHAR(SEQ) SEQ from  O_RETAILER_DEPT_SUPP order by O_RETAILER_DEPT_SUPP.SOURCEID");

            ocdSelect12.CommandText = sql12.ToString();
            ocdSelect12.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbRetailerDeptSupp.Merge(objDataLayer12.GetResultDT(ocdSelect12));

            DataLayer objDataLayer13 = new DataLayer(_strConn);
            OracleCommand ocdSelect13 = new OracleCommand();
            StringBuilder sql13 = new StringBuilder();
            //sql13.Append("SELECT * FROM O_UOM_MAPPING_LIST");
            sql13.Append("select RF_UOMCODE,TO_CHAR(OGRDS_UOMID) OGRDS_UOMID, OGRDS_UOMDESC,PREFERED_UNIT_DSCR from O_UOM_MAPPING_LIST order by RF_UOMCODE");

            ocdSelect13.CommandText = sql13.ToString();
            ocdSelect13.CommandType = System.Data.CommandType.Text;
            ProcessTool.dtbUomMappingList.Merge(objDataLayer13.GetResultDT(ocdSelect13));
        }
        public void MDLInputSettings_PopulateData(string tableName, DataTable dt)
        {
            #region 1. ACTIVE_CATEGORY
            if (tableName == "O_ACTIVE_CATEGORY")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_ACTIVE_CATEGORY;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_ACTIVE_CATEGORY(CATEGORYATTRVALUE, SERVICE, TYPE) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString()) + ";" ;
                }
                    Sqlquery += "END;";
                    objDataLayer.TableOperation_BySQL(Sqlquery, true);
                        if (objDataLayer != null)
                            objDataLayer.CheckTransaction(true);
                   
             }
            #endregion

            #region 2. PRE_SET_CHAR
            if (tableName == "O_PRE_SET_CHAR")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_PRE_SET_CHAR;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_PRE_SET_CHAR(TYPE,ATTRNO,REQUIRED) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString()) + ";";
                }

                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
                
            }
            #endregion

            #region 3. PROC_GRP_SET
            if (tableName == "O_PROC_GRP_SET")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_PROC_GRP_SET;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    
                        Sqlquery += "INSERT INTO O_PROC_GRP_SET(PROC_GROUP_SET_ID,PROC_GROUP_SET_NAME,FOLLOW_EXT_RULE,EXCP_SENT_TO_XCD_BRWSR,EXCP_SENT_TO_UNCDBLE_ITM,EXCP_SEN_TO_SUPER_GTP_ITM) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}','{3}','{4}','{5}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 

            }
            #endregion

            #region 4. EXT_CODE_GRP
            if (tableName == "O_EXT_CODE_GRP")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_EXT_CODE_GRP;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_EXT_CODE_GRP(SHORTNAME,EXTERNAL_CODE_GROUP_NAME) Values(";
                        Sqlquery += string.Format("'{0}','{1}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 

            }
            #endregion

            #region 5. ACTIVE_RETAILER
            if (tableName == "O_ACTIVE_RETAILER")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_ACTIVE_RETAILER;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_ACTIVE_RETAILER(SOURCEID,PROC_GROUP_SET_NAME,TYPE,EAN,UPC,LAC,CIP) Values(";
                        Sqlquery += string.Format("{0},'{1}','{2}','{3}','{4}','{5}','{6}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString(), dt.Rows[i][6].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 6. DIC_LVL_FLAG_SETTING
            if (tableName == "O_DIC_LVL_FLAG_SETTING")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_DIC_LVL_FLAG_SETTING;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_DIC_LVL_FLAG_SETTING(ATTRNO,OGRDSID,MAP_TO_OGRDS_CHR_VAL,LONGDESC ,OTHERDESC,ALTER30MAX,ALTER30MAX_OTHER,ALTER5MAX,CATEGORY_FLAG,CHAR_TYPE,NUMERIC_FLAG,CUSTOMER_FLAG,";
                        Sqlquery += " FIXED_ITEM_VALUE,COPY_ITEM,MULTI_VALUE,ABBREVIATE_VALUE,FIXED_VALUE_LIST,TRANSLATION_IND, CHR_VAL_DESCRIPTION_ONLY,LOCAL,SORT_ORDER) Values(";
                        Sqlquery += string.Format("{0},{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}',{20})",
                            dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString() == "" ? "NULL" : dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString(),
                                   dt.Rows[i][6].ToString(), dt.Rows[i][7].ToString(), dt.Rows[i][8].ToString(), dt.Rows[i][9].ToString(), dt.Rows[i][10].ToString(), dt.Rows[i][11].ToString(),
                                   dt.Rows[i][12].ToString(), dt.Rows[i][13].ToString(), dt.Rows[i][14].ToString(), dt.Rows[i][15].ToString(), dt.Rows[i][16].ToString(), dt.Rows[i][17].ToString(),
                                   dt.Rows[i][18].ToString(), dt.Rows[i][19].ToString(), dt.Rows[i][20].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 7. PC_ASGN_LVL_FLAG_SETTING
            if (tableName == "O_PC_ASGN_LVL_FLAG_SETTING")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_PC_ASGN_LVL_FLAG_SETTING;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_PC_ASGN_LVL_FLAG_SETTING(CATEGORYATTRVALUE,ATTRNO,FIXED_VALUE_LIST,COPY_ITEM_VAL,FIELD_COLL_FLAG ,MANDATORY_FLAG,ALIGNED, " + "\"" + "UNIQUE" + "\"" + " , LOCAL,SORT_ORDER) Values(";
                        Sqlquery += string.Format("'{0}',{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}',{9})", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString(), dt.Rows[i][6].ToString(), dt.Rows[i][7].ToString(), dt.Rows[i][8].ToString(), dt.Rows[i][9].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 8. MOD_ASGN_LVL_FLAG_SETTING
            if (tableName == "O_MOD_ASGN_LVL_FLAG_SETTING")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_MOD_ASGN_LVL_FLAG_SETTING;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += " INSERT INTO O_MOD_ASGN_LVL_FLAG_SETTING(MODULEID,ATTRNO,FIELD_COLL_FLAG,MANDATORY_FLAG) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}','{3}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 9. HEADING_TYPE
            if (tableName == "O_HEADING_TYPE")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_HEADING_TYPE;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_HEADING_TYPE(HEADING_ID,HEADING_NAME,MAXIMUM_LENGTH,ONLINE_HDNG_FLAG,OUTPUT_FLAG,ACTIVE_FLAG,ALTERNATE_HEADING_FLAG,ASGN_LVL,HDNG_TYP,LOCALLANG) Values(";
                        Sqlquery += string.Format("'{0}','{1}',{2},'{3}','{4}','{5}','{6}','{7}','{8}','{9}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString(), dt.Rows[i][6].ToString(), dt.Rows[i][7].ToString(), dt.Rows[i][8].ToString(), dt.Rows[i][9].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 

            }
            #endregion

            #region 10. HEADING_PC_SETTING
            if (tableName == "O_HEADING_PC_SETTING")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_HEADING_PC_SETTING;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_HEADING_PC_SETTING(HEADING_NAME,CATEGORYATTRVALUE,CHAR_ID,SEQUENCE,CONCATENATOR,DROP_PRIORITY) Values("; //BGRF-1961
                        Sqlquery += string.Format("'{0}','{1}','{2}',{3},'{4}',{5})", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 11. HEADING_MOD_SETTING
            if (tableName == "O_HEADING_MOD_SETTING")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_HEADING_MOD_SETTING;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_HEADING_MOD_SETTING(HEADING_NAME,MODULEID,CHAR_ID,SEQUENCE,CONCATENATOR,DROP_PRIORITY) Values("; //BGRF-1961
                        Sqlquery += string.Format("'{0}',{1},'{2}',{3},'{4}',{5})", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), dt.Rows[i][4].ToString(), dt.Rows[i][5].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 12. RETAILER_DEPT_SUPP
            if (tableName == "O_RETAILER_DEPT_SUPP")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_RETAILER_DEPT_SUPP;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_RETAILER_DEPT_SUPP(SOURCEID,ATTRNO,TYPE,SEQ) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}',{3})", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

            #region 13. UOM_MAPPING_LIST
            if (tableName == "O_UOM_MAPPING_LIST")
            {
                DataLayer objDataLayer = new DataLayer(_strConn, true);
                Sqlquery = "BEGIN ";
                Sqlquery += "DELETE FROM O_UOM_MAPPING_LIST;";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                        Sqlquery += "INSERT INTO O_UOM_MAPPING_LIST(RF_UOMCODE,OGRDS_UOMID,OGRDS_UOMDESC,PREFERED_UNIT_DSCR) Values(";
                        Sqlquery += string.Format("'{0}','{1}','{2}','{3}')", dt.Rows[i][0].ToString(), dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString()) + ";";
                }
                Sqlquery += "END;";
                objDataLayer.TableOperation_BySQL(Sqlquery, true);
                if (objDataLayer != null)
                    objDataLayer.CheckTransaction(true); 
            }
            #endregion

        }
        #endregion

        //start change for BGRF-1795
        //public void PopulateCommonData()
        public DataTable PopulateCommonData()
        {
            try
            {
                DataLayer objDataLayer = new DataLayer(_strConn);
                OracleCommand ocdPopulateCommonData = new OracleCommand();
                //start -commented above line and added for countrycode and locallanguagecode from country table BGRF-1795
                //ocdPopulateCommonData.CommandText = "SELECT IMDB_COUNTRYCODE, LOCALLANGUAGE_CODE FROM rrs.country WHERE HOST='"+Common.strDBName.ToUpper().ToString()+"' AND USERID='"+Common.strUserName+"' AND PASSWORD='"+Common.strPassword+"'";
                ocdPopulateCommonData.CommandText = "SELECT IMDB_COUNTRYCODE, LOCALLANGUAGE_CODE FROM RRS.COUNTRY WHERE ROWNUM <= 1";
                DataTable dtbResult = objDataLayer.GetResultDT(ocdPopulateCommonData, false);
                return dtbResult;
                //end -added for countrycode and locallanguagecode from country table BGRF-1795
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //end change for BGRF-1795
    }

    public sealed class Common
    {
        public static string strDBName { get; set; }

        public static string strUserName { get; set; }

        public static string strPassword { get; set; }

        public static string CountryCode { get; set; }

        public static string LocalLangCode { get; set; }
    }
}
