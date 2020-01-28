using MDL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MDLInputFileSetting
{
    public partial class Login : Form
    {

        #region declareVariables
        DialogResult result;
        MessageBoxButtons buttons = MessageBoxButtons.OK;
        BusinessLayer blyBL = null;
        private static string strConn = null;
        public static string strConnectionString = "";
        public static string strLogFilePath = "";

        #endregion 
        public Login()
        {
            InitializeComponent();

            txtDatabase.Text = "rfma01dv";
            txtUserId.Text = "rrs";
            txtPassword.Text = "rrs";
        }

        //start login validation --SIVA BGRF-1734
        private void btnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                DataLayer objDataLayer = null;

                if (((txtUserId.Text.ToString() == string.Empty) || (txtUserId.Text.ToString() == null)) ||
                    ((txtPassword.Text.ToString() == string.Empty) || (txtPassword.Text.ToString() == null)) ||
                    ((txtDatabase.Text.ToString() == string.Empty) || (txtDatabase.Text.ToString() == null)))
                {
                    result = MessageBox.Show("Please fill in login details and try again.", "Error", buttons, MessageBoxIcon.Error);
                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }
                }
                else
                {
                    try
                    {
                        if (((txtUserId.Text.ToString() != string.Empty) || (txtUserId.Text.ToString() != null)) &&
                           ((txtPassword.Text.ToString() != string.Empty) || (txtPassword.Text.ToString() != null)) &&
                           ((txtDatabase.Text.ToString() != string.Empty) || (txtDatabase.Text.ToString() != null)))
                        {
                            Common.strUserName = txtUserId.Text.ToString();
                            Common.strPassword = txtPassword.Text.ToString();
                            Common.strDBName = txtDatabase.Text.ToString();

                            strConn = "Data Source=" + Common.strDBName + ";Persist Security Info=True;User ID=" + Common.strUserName + ";Password=" + Common.strPassword;
                            strConnectionString = Common.strUserName.ToUpper().Trim() + "/" + Common.strPassword.ToUpper().Trim() + "@" + Common.strDBName.ToUpper().Trim();
                            objDataLayer = new DataLayer(strConn);

                            #region populateCommonData
                            try
                            {
                                blyBL = new BusinessLayer();
                                DataTable dtcountry = blyBL.PopulateCommonData();
                                if (string.IsNullOrEmpty(dtcountry.Rows[0]["LOCALLANGUAGE_CODE"].ToString()) ||
                                    string.IsNullOrWhiteSpace(dtcountry.Rows[0]["LOCALLANGUAGE_CODE"].ToString()))
                                {
                                    result = MessageBox.Show("LOCALLANGUAGECODE not set properly.", "Error", buttons, MessageBoxIcon.Error);
                                    if (result == System.Windows.Forms.DialogResult.OK)
                                        return;
                                }
                                else if (string.IsNullOrEmpty(dtcountry.Rows[0]["IMDB_COUNTRYCODE"].ToString()) ||
                                    string.IsNullOrWhiteSpace(dtcountry.Rows[0]["IMDB_COUNTRYCODE"].ToString()))
                                {
                                    result = MessageBox.Show("COUNTRYCODE not set properly.", "Error", buttons, MessageBoxIcon.Error);
                                    if (result == System.Windows.Forms.DialogResult.OK)
                                        return;
                                }
                                else
                                {
                                    Common.CountryCode = dtcountry.Rows[0]["IMDB_COUNTRYCODE"].ToString();
                                    Common.LocalLangCode = dtcountry.Rows[0]["LOCALLANGUAGE_CODE"].ToString();
                                    //start BGRF-1961
                                    Common.strDBName = txtDatabase.Text;
                                    Common.strUserName = txtUserId.Text;
                                    Common.strPassword = txtPassword.Text;
                                    this.Hide();
                                    ProcessTool form = new ProcessTool();
                                    form.Show();
                                    //end BGRF-1961
                                }
                            }
                            catch (Exception ex)
                            {
                                result = MessageBox.Show(ex.Message.ToString(), "Error", buttons, MessageBoxIcon.Error);
                                if (result == System.Windows.Forms.DialogResult.OK)
                                    return;
                            }
                            #endregion populateCommonData
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("connect identifier specified"))
                        {
                            result = MessageBox.Show("Invalid Database.", "Error", buttons, MessageBoxIcon.Error);
                            if (result == System.Windows.Forms.DialogResult.OK)
                            {
                                txtDatabase.Text = "";
                                txtDatabase.Focus();
                                return;
                            }
                        }
                        else if (ex.Message.Contains("username/password"))
                        {
                            result = MessageBox.Show("Invalid UserID/Password.", "Error", buttons, MessageBoxIcon.Error);
                            if (result == System.Windows.Forms.DialogResult.OK)
                            {
                                txtUserId.Text = "";
                                txtPassword.Text = "";
                                txtUserId.Focus();
                                return;
                            }
                        }
                        else
                        {
                            result = MessageBox.Show(ex.Message, "Error", buttons, MessageBoxIcon.Error);
                            if (result == System.Windows.Forms.DialogResult.OK)
                            {
                                txtUserId.Text = "";
                                txtPassword.Text = "";
                                txtDatabase.Text = "";
                                txtUserId.Focus();
                                return;
                            }
                        }
                    }
                }
            }

            catch
            {
            }
            finally
            {
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //end login validation --SIVA BGRF-1734
    }
}
