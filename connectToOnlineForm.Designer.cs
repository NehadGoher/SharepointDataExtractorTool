namespace ContentTypeExtractor
{
    partial class connectToOnlineForm
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
            this.label_username = new System.Windows.Forms.Label();
            this.label_password = new System.Windows.Forms.Label();
            this.btn_login = new System.Windows.Forms.Button();
            this.txtbx_username = new System.Windows.Forms.TextBox();
            this.txtbx_password = new System.Windows.Forms.TextBox();
            this.label_url = new System.Windows.Forms.Label();
            this.txtbx_url = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label_status = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label_username
            // 
            this.label_username.AutoSize = true;
            this.label_username.Location = new System.Drawing.Point(197, 26);
            this.label_username.Name = "label_username";
            this.label_username.Size = new System.Drawing.Size(73, 17);
            this.label_username.TabIndex = 0;
            this.label_username.Text = "Username";
            // 
            // label_password
            // 
            this.label_password.AutoSize = true;
            this.label_password.Location = new System.Drawing.Point(197, 78);
            this.label_password.Name = "label_password";
            this.label_password.Size = new System.Drawing.Size(69, 17);
            this.label_password.TabIndex = 1;
            this.label_password.Text = "Password";
            // 
            // btn_login
            // 
            this.btn_login.Location = new System.Drawing.Point(361, 170);
            this.btn_login.Name = "btn_login";
            this.btn_login.Size = new System.Drawing.Size(78, 40);
            this.btn_login.TabIndex = 2;
            this.btn_login.Text = "Connect";
            this.btn_login.UseVisualStyleBackColor = true;
            this.btn_login.Click += new System.EventHandler(this.btn_login_Click);
            // 
            // txtbx_username
            // 
            this.txtbx_username.Location = new System.Drawing.Point(313, 23);
            this.txtbx_username.Name = "txtbx_username";
            this.txtbx_username.Size = new System.Drawing.Size(177, 22);
            this.txtbx_username.TabIndex = 3;
            // 
            // txtbx_password
            // 
            this.txtbx_password.Location = new System.Drawing.Point(313, 78);
            this.txtbx_password.Name = "txtbx_password";
            this.txtbx_password.Size = new System.Drawing.Size(177, 22);
            this.txtbx_password.TabIndex = 4;
            // 
            // label_url
            // 
            this.label_url.AutoSize = true;
            this.label_url.Location = new System.Drawing.Point(200, 131);
            this.label_url.Name = "label_url";
            this.label_url.Size = new System.Drawing.Size(36, 17);
            this.label_url.TabIndex = 5;
            this.label_url.Text = "URL";
            // 
            // txtbx_url
            // 
            this.txtbx_url.Location = new System.Drawing.Point(313, 131);
            this.txtbx_url.Name = "txtbx_url";
            this.txtbx_url.Size = new System.Drawing.Size(177, 22);
            this.txtbx_url.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 241);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Status: ";
            // 
            // label_status
            // 
            this.label_status.AutoSize = true;
            this.label_status.Location = new System.Drawing.Point(74, 241);
            this.label_status.Name = "label_status";
            this.label_status.Size = new System.Drawing.Size(0, 17);
            this.label_status.TabIndex = 8;
            // 
            // connectToOnlineForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(764, 289);
            this.Controls.Add(this.label_status);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtbx_url);
            this.Controls.Add(this.label_url);
            this.Controls.Add(this.txtbx_password);
            this.Controls.Add(this.txtbx_username);
            this.Controls.Add(this.btn_login);
            this.Controls.Add(this.label_password);
            this.Controls.Add(this.label_username);
            this.Name = "connectToOnlineForm";
            this.Text = "Connect To Online";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label_username;
        private System.Windows.Forms.Label label_password;
        private System.Windows.Forms.Button btn_login;
        private System.Windows.Forms.TextBox txtbx_username;
        private System.Windows.Forms.TextBox txtbx_password;
        private System.Windows.Forms.Label label_url;
        private System.Windows.Forms.TextBox txtbx_url;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label_status;
    }
}