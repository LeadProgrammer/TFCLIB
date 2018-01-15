using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace LIBDI
{
	/// <summary>
	/// Summary description for SAPPopUpPrompt.
	/// </summary>
	public class SAPPopUpPrompt : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox saName;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox saPassword;
        private System.Windows.Forms.Button CreateSQLuser;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        private string gDBName;
        private string gUser;
        private string gPassword;
        private System.Windows.Forms.Button buttonCancel;
		private System.Windows.Forms.TextBox textBoxServer;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox textBoxDBName;
		public bool CancelWasPressed;

		public SAPPopUpPrompt(string ServerName, string DBName)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			textBoxServer.Text=ServerName;
			textBoxDBName.Text=DBName;
			CancelWasPressed=false;
            MicroSoftWindows.WindowWrapper oWindow = new MicroSoftWindows.WindowWrapper();
            ShowDialog(oWindow);
            //
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public void GetUserAndPassowrd(ref string User, ref string Password)
		{

			User = gUser;
			Password = gPassword;
		}
        public string GetUser()
        {
            return gUser;
        }
        public string GetPassword()
        {
            return gPassword;
        }
        public string GetDBName()
        {
            return gDBName;
        }
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SAPPopUpPrompt));
            this.label1 = new System.Windows.Forms.Label();
            this.saName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.saPassword = new System.Windows.Forms.TextBox();
            this.CreateSQLuser = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.textBoxServer = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxDBName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(400, 32);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please enter the SQL Administator User and Administator Password. User sa is the " +
                "most common user name for the Administator.";
            // 
            // saName
            // 
            this.saName.Location = new System.Drawing.Point(168, 120);
            this.saName.Name = "saName";
            this.saName.Size = new System.Drawing.Size(184, 20);
            this.saName.TabIndex = 1;
            this.saName.Text = "sa";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 120);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(152, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "SQL Administator Name:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(8, 152);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(152, 16);
            this.label3.TabIndex = 4;
            this.label3.Text = "SQL Administator Password:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // saPassword
            // 
            this.saPassword.Location = new System.Drawing.Point(168, 152);
            this.saPassword.Name = "saPassword";
            this.saPassword.Size = new System.Drawing.Size(184, 20);
            this.saPassword.TabIndex = 3;
            this.saPassword.Text = "sa";
            // 
            // CreateSQLuser
            // 
            this.CreateSQLuser.Location = new System.Drawing.Point(24, 184);
            this.CreateSQLuser.Name = "CreateSQLuser";
            this.CreateSQLuser.Size = new System.Drawing.Size(248, 23);
            this.CreateSQLuser.TabIndex = 5;
            this.CreateSQLuser.Text = "Create Addon SQL User";
            this.CreateSQLuser.Click += new System.EventHandler(this.CreateSQLuser_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(344, 184);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 7;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // textBoxServer
            // 
            this.textBoxServer.CausesValidation = false;
            this.textBoxServer.Location = new System.Drawing.Point(168, 56);
            this.textBoxServer.Name = "textBoxServer";
            this.textBoxServer.Size = new System.Drawing.Size(184, 20);
            this.textBoxServer.TabIndex = 8;
            this.textBoxServer.TextChanged += new System.EventHandler(this.textBoxServer_TextChanged);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(8, 56);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(152, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "Server Name:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(8, 88);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(152, 16);
            this.label5.TabIndex = 11;
            this.label5.Text = "Database Name:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxDBName
            // 
            this.textBoxDBName.CausesValidation = false;
            this.textBoxDBName.Location = new System.Drawing.Point(168, 88);
            this.textBoxDBName.Name = "textBoxDBName";
            this.textBoxDBName.Size = new System.Drawing.Size(184, 20);
            this.textBoxDBName.TabIndex = 10;
            // 
            // SAPPopUpPrompt
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(440, 320);
            this.ControlBox = false;
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBoxDBName);
            this.Controls.Add(this.textBoxServer);
            this.Controls.Add(this.saPassword);
            this.Controls.Add(this.saName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.CreateSQLuser);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SAPPopUpPrompt";
            this.Text = "SAPPopUpPrompt";
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void CreateSQLuser_Click(object sender, System.EventArgs e)
		{
			gUser = saName.Text;
			gPassword = saPassword.Text;
			CancelWasPressed=false;
			Close();
		}

		private void buttonCancel_Click(object sender, System.EventArgs e)
		{
			CancelWasPressed=true;
			Close();

		}

		private void textBoxServer_TextChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
