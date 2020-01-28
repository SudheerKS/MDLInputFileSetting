namespace MDLInputFileSetting
{
    partial class Login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.txtDatabase = new DevExpress.XtraEditors.TextEdit();
            this.lblDatabase = new DevExpress.XtraEditors.LabelControl();
            this.txtUserId = new DevExpress.XtraEditors.TextEdit();
            this.lblUserId = new DevExpress.XtraEditors.LabelControl();
            this.txtPassword = new DevExpress.XtraEditors.TextEdit();
            this.lblPassword = new DevExpress.XtraEditors.LabelControl();
            this.btnLogin = new DevExpress.XtraEditors.SimpleButton();
            this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.defaultLookAndFeel1 = new DevExpress.LookAndFeel.DefaultLookAndFeel(this.components);
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.txtDatabase.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUserId.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPassword.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // txtDatabase
            // 
            this.txtDatabase.Location = new System.Drawing.Point(147, 127);
            this.txtDatabase.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtDatabase.Name = "txtDatabase";
            this.txtDatabase.Size = new System.Drawing.Size(243, 22);
            this.txtDatabase.TabIndex = 6;
            // 
            // lblDatabase
            // 
            this.lblDatabase.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblDatabase.Location = new System.Drawing.Point(45, 130);
            this.lblDatabase.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lblDatabase.Name = "lblDatabase";
            this.lblDatabase.Size = new System.Drawing.Size(71, 20);
            this.lblDatabase.TabIndex = 5;
            this.lblDatabase.Text = "Database";
            // 
            // txtUserId
            // 
            this.txtUserId.Location = new System.Drawing.Point(147, 55);
            this.txtUserId.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtUserId.Name = "txtUserId";
            this.txtUserId.Size = new System.Drawing.Size(243, 22);
            this.txtUserId.TabIndex = 8;
            // 
            // lblUserId
            // 
            this.lblUserId.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblUserId.Location = new System.Drawing.Point(45, 57);
            this.lblUserId.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lblUserId.Name = "lblUserId";
            this.lblUserId.Size = new System.Drawing.Size(61, 20);
            this.lblUserId.TabIndex = 7;
            this.lblUserId.Text = "User ID";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(147, 91);
            this.txtPassword.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Properties.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(243, 22);
            this.txtPassword.TabIndex = 10;
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblPassword.Location = new System.Drawing.Point(45, 95);
            this.lblPassword.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(71, 20);
            this.lblPassword.TabIndex = 9;
            this.lblPassword.Text = "Password";
            // 
            // btnLogin
            // 
            this.btnLogin.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.btnLogin.Appearance.Options.UseFont = true;
            this.btnLogin.Location = new System.Drawing.Point(467, 17);
            this.btnLogin.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(100, 28);
            this.btnLogin.TabIndex = 11;
            this.btnLogin.Text = "OK";
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Appearance.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold);
            this.btnCancel.Appearance.Options.UseFont = true;
            this.btnCancel.Location = new System.Drawing.Point(467, 50);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 28);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // defaultLookAndFeel1
            // 
            this.defaultLookAndFeel1.LookAndFeel.SkinName = "MySkin_Money Twins1";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(39, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(340, 20);
            this.label1.TabIndex = 13;
            this.label1.Text = "Enter your user ID and password.";
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(581, 180);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnLogin);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.lblPassword);
            this.Controls.Add(this.txtUserId);
            this.Controls.Add(this.lblUserId);
            this.Controls.Add(this.txtDatabase);
            this.Controls.Add(this.lblDatabase);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            ((System.ComponentModel.ISupportInitialize)(this.txtDatabase.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUserId.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtPassword.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.TextEdit txtDatabase;
        private DevExpress.XtraEditors.LabelControl lblDatabase;
        private DevExpress.XtraEditors.TextEdit txtUserId;
        private DevExpress.XtraEditors.LabelControl lblUserId;
        private DevExpress.XtraEditors.TextEdit txtPassword;
        private DevExpress.XtraEditors.LabelControl lblPassword;
        private DevExpress.XtraEditors.SimpleButton btnLogin;
        private DevExpress.XtraEditors.SimpleButton btnCancel;
        private DevExpress.LookAndFeel.DefaultLookAndFeel defaultLookAndFeel1;
        private System.Windows.Forms.Label label1;
    }
}