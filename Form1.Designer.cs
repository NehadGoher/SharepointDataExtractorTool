namespace ContentTypeExtractor
{
    partial class Form1
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
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.test_btn = new System.Windows.Forms.Button();
            this.bt_createContentType = new System.Windows.Forms.Button();
            this.btn_createSiteColumn = new System.Windows.Forms.Button();
            this.btn_createLibrary = new System.Windows.Forms.Button();
            this.btn_deleteContentTypes = new System.Windows.Forms.Button();
            this.btn_deleteSiteColumn = new System.Windows.Forms.Button();
            this.btn_deleteLibrary = new System.Windows.Forms.Button();
            this.listContent = new System.Windows.Forms.ListBox();
            this.btn_clear = new System.Windows.Forms.Button();
            this.listSite = new System.Windows.Forms.ListBox();
            this.listLib = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(28, 44);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(145, 39);
            this.buttonBrowse.TabIndex = 0;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.button1_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(621, 131);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(450, 446);
            this.richTextBox1.TabIndex = 6;
            this.richTextBox1.Text = "";
            // 
            // test_btn
            // 
            this.test_btn.Location = new System.Drawing.Point(28, 89);
            this.test_btn.Name = "test_btn";
            this.test_btn.Size = new System.Drawing.Size(145, 32);
            this.test_btn.TabIndex = 8;
            this.test_btn.Text = "Load Data To List";
            this.test_btn.UseVisualStyleBackColor = true;
            this.test_btn.Click += new System.EventHandler(this.test_btn_Click);
            // 
            // bt_createContentType
            // 
            this.bt_createContentType.Location = new System.Drawing.Point(621, 44);
            this.bt_createContentType.Name = "bt_createContentType";
            this.bt_createContentType.Size = new System.Drawing.Size(167, 39);
            this.bt_createContentType.TabIndex = 9;
            this.bt_createContentType.Text = "Create Content Types";
            this.bt_createContentType.UseVisualStyleBackColor = true;
            this.bt_createContentType.Click += new System.EventHandler(this.bt_createContentType_Click);
            // 
            // btn_createSiteColumn
            // 
            this.btn_createSiteColumn.Location = new System.Drawing.Point(795, 44);
            this.btn_createSiteColumn.Name = "btn_createSiteColumn";
            this.btn_createSiteColumn.Size = new System.Drawing.Size(144, 39);
            this.btn_createSiteColumn.TabIndex = 10;
            this.btn_createSiteColumn.Text = "Create Site Column";
            this.btn_createSiteColumn.UseVisualStyleBackColor = true;
            this.btn_createSiteColumn.Click += new System.EventHandler(this.btn_createSiteColumn_Click);
            // 
            // btn_createLibrary
            // 
            this.btn_createLibrary.Location = new System.Drawing.Point(945, 44);
            this.btn_createLibrary.Name = "btn_createLibrary";
            this.btn_createLibrary.Size = new System.Drawing.Size(120, 39);
            this.btn_createLibrary.TabIndex = 11;
            this.btn_createLibrary.Text = "Create Library";
            this.btn_createLibrary.UseVisualStyleBackColor = true;
            this.btn_createLibrary.Click += new System.EventHandler(this.btn_createLibrary_Click);
            // 
            // btn_deleteContentTypes
            // 
            this.btn_deleteContentTypes.Location = new System.Drawing.Point(621, 88);
            this.btn_deleteContentTypes.Name = "btn_deleteContentTypes";
            this.btn_deleteContentTypes.Size = new System.Drawing.Size(167, 33);
            this.btn_deleteContentTypes.TabIndex = 12;
            this.btn_deleteContentTypes.Text = "Delete Content Types";
            this.btn_deleteContentTypes.UseVisualStyleBackColor = true;
            this.btn_deleteContentTypes.Click += new System.EventHandler(this.btn_deleteContentTypes_Click);
            // 
            // btn_deleteSiteColumn
            // 
            this.btn_deleteSiteColumn.Location = new System.Drawing.Point(794, 89);
            this.btn_deleteSiteColumn.Name = "btn_deleteSiteColumn";
            this.btn_deleteSiteColumn.Size = new System.Drawing.Size(145, 33);
            this.btn_deleteSiteColumn.TabIndex = 13;
            this.btn_deleteSiteColumn.Text = "Delete Site Columns";
            this.btn_deleteSiteColumn.UseVisualStyleBackColor = true;
            this.btn_deleteSiteColumn.Click += new System.EventHandler(this.btn_deleteSiteColumn_Click);
            // 
            // btn_deleteLibrary
            // 
            this.btn_deleteLibrary.Location = new System.Drawing.Point(945, 88);
            this.btn_deleteLibrary.Name = "btn_deleteLibrary";
            this.btn_deleteLibrary.Size = new System.Drawing.Size(120, 33);
            this.btn_deleteLibrary.TabIndex = 14;
            this.btn_deleteLibrary.Text = "Delete Library";
            this.btn_deleteLibrary.UseVisualStyleBackColor = true;
            this.btn_deleteLibrary.Click += new System.EventHandler(this.btn_deleteLibrary_Click);
            // 
            // listContent
            // 
            this.listContent.FormattingEnabled = true;
            this.listContent.ItemHeight = 16;
            this.listContent.Location = new System.Drawing.Point(28, 199);
            this.listContent.Name = "listContent";
            this.listContent.Size = new System.Drawing.Size(155, 388);
            this.listContent.TabIndex = 15;
            // 
            // btn_clear
            // 
            this.btn_clear.Location = new System.Drawing.Point(28, 127);
            this.btn_clear.Name = "btn_clear";
            this.btn_clear.Size = new System.Drawing.Size(145, 31);
            this.btn_clear.TabIndex = 16;
            this.btn_clear.Text = "Clear";
            this.btn_clear.UseVisualStyleBackColor = true;
            this.btn_clear.Click += new System.EventHandler(this.btn_clear_Click);
            // 
            // listSite
            // 
            this.listSite.FormattingEnabled = true;
            this.listSite.ItemHeight = 16;
            this.listSite.Location = new System.Drawing.Point(219, 199);
            this.listSite.Name = "listSite";
            this.listSite.Size = new System.Drawing.Size(167, 388);
            this.listSite.TabIndex = 17;
            // 
            // listLib
            // 
            this.listLib.FormattingEnabled = true;
            this.listLib.ItemHeight = 16;
            this.listLib.Location = new System.Drawing.Point(423, 199);
            this.listLib.Name = "listLib";
            this.listLib.Size = new System.Drawing.Size(169, 388);
            this.listLib.TabIndex = 18;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 180);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 17);
            this.label1.TabIndex = 19;
            this.label1.Text = "Content Type List";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(216, 179);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(116, 17);
            this.label2.TabIndex = 20;
            this.label2.Text = "Site Columns List";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(430, 179);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 17);
            this.label3.TabIndex = 21;
            this.label3.Text = "Libraries Lists";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1083, 611);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listLib);
            this.Controls.Add(this.listSite);
            this.Controls.Add(this.btn_clear);
            this.Controls.Add(this.listContent);
            this.Controls.Add(this.btn_deleteLibrary);
            this.Controls.Add(this.btn_deleteSiteColumn);
            this.Controls.Add(this.btn_deleteContentTypes);
            this.Controls.Add(this.btn_createLibrary);
            this.Controls.Add(this.btn_createSiteColumn);
            this.Controls.Add(this.bt_createContentType);
            this.Controls.Add(this.test_btn);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.buttonBrowse);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button test_btn;
        private System.Windows.Forms.Button bt_createContentType;
        private System.Windows.Forms.Button btn_createSiteColumn;
        private System.Windows.Forms.Button btn_createLibrary;
        private System.Windows.Forms.Button btn_deleteContentTypes;
        private System.Windows.Forms.Button btn_deleteSiteColumn;
        private System.Windows.Forms.Button btn_deleteLibrary;
        private System.Windows.Forms.ListBox listContent;
        private System.Windows.Forms.Button btn_clear;
        private System.Windows.Forms.ListBox listSite;
        private System.Windows.Forms.ListBox listLib;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

