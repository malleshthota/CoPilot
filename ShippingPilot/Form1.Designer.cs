﻿namespace ShippingPilot
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblUserName = new System.Windows.Forms.Label();
            this.lblDateTime = new System.Windows.Forms.Label();
            this.lblInfo = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtResponsePath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnLineItem = new System.Windows.Forms.Button();
            this.txtLineItem = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblRespFilePath = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Verdana", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.Location = new System.Drawing.Point(-3, 136);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Pilot Data File";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnSubmit
            // 
            this.btnSubmit.BackColor = System.Drawing.Color.OrangeRed;
            this.btnSubmit.Font = new System.Drawing.Font("Verdana", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubmit.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnSubmit.Location = new System.Drawing.Point(128, 331);
            this.btnSubmit.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnSubmit.Size = new System.Drawing.Size(266, 89);
            this.btnSubmit.TabIndex = 1;
            this.btnSubmit.Text = "Submit";
            this.btnSubmit.UseVisualStyleBackColor = false;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("Verdana", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.ForeColor = System.Drawing.Color.MediumBlue;
            this.txtFilePath.Location = new System.Drawing.Point(107, 133);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(342, 22);
            this.txtFilePath.TabIndex = 2;
            // 
            // btnBrowse
            // 
            this.btnBrowse.BackColor = System.Drawing.Color.Orange;
            this.btnBrowse.Font = new System.Drawing.Font("Verdana", 9.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBrowse.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnBrowse.Location = new System.Drawing.Point(457, 133);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(137, 22);
            this.btnBrowse.TabIndex = 3;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = false;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::ShippingPilot.Properties.Resources.logo_2010_3;
            this.pictureBox1.Location = new System.Drawing.Point(305, 12);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(274, 68);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // lblUserName
            // 
            this.lblUserName.AutoSize = true;
            this.lblUserName.Font = new System.Drawing.Font("Verdana", 9.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUserName.ForeColor = System.Drawing.Color.OrangeRed;
            this.lblUserName.Location = new System.Drawing.Point(38, 26);
            this.lblUserName.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(38, 16);
            this.lblUserName.TabIndex = 6;
            this.lblUserName.Text = "......";
            // 
            // lblDateTime
            // 
            this.lblDateTime.AutoSize = true;
            this.lblDateTime.Font = new System.Drawing.Font("Verdana", 9.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDateTime.ForeColor = System.Drawing.Color.Gray;
            this.lblDateTime.Location = new System.Drawing.Point(38, 43);
            this.lblDateTime.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDateTime.Name = "lblDateTime";
            this.lblDateTime.Size = new System.Drawing.Size(28, 16);
            this.lblDateTime.TabIndex = 7;
            this.lblDateTime.Text = "....";
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Font = new System.Drawing.Font("Verdana", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfo.ForeColor = System.Drawing.Color.Navy;
            this.lblInfo.Location = new System.Drawing.Point(125, 298);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(42, 18);
            this.lblInfo.TabIndex = 8;
            this.lblInfo.Text = "Info";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(-3, 231);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 14);
            this.label2.TabIndex = 9;
            this.label2.Text = "Response File";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // txtResponsePath
            // 
            this.txtResponsePath.ForeColor = System.Drawing.Color.Blue;
            this.txtResponsePath.Location = new System.Drawing.Point(107, 229);
            this.txtResponsePath.Name = "txtResponsePath";
            this.txtResponsePath.Size = new System.Drawing.Size(342, 22);
            this.txtResponsePath.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Verdana", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(107, 258);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(287, 12);
            this.label3.TabIndex = 12;
            this.label3.Text = "*Provide Response Folder. Ex:: C:\\PilotResponse\\";
            // 
            // btnLineItem
            // 
            this.btnLineItem.BackColor = System.Drawing.Color.Orange;
            this.btnLineItem.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnLineItem.Location = new System.Drawing.Point(457, 179);
            this.btnLineItem.Name = "btnLineItem";
            this.btnLineItem.Size = new System.Drawing.Size(137, 24);
            this.btnLineItem.TabIndex = 15;
            this.btnLineItem.Text = "Browse";
            this.btnLineItem.UseVisualStyleBackColor = false;
            this.btnLineItem.Click += new System.EventHandler(this.btnLineItem_Click);
            // 
            // txtLineItem
            // 
            this.txtLineItem.ForeColor = System.Drawing.Color.Blue;
            this.txtLineItem.Location = new System.Drawing.Point(107, 181);
            this.txtLineItem.Name = "txtLineItem";
            this.txtLineItem.Size = new System.Drawing.Size(342, 22);
            this.txtLineItem.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(-3, 183);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(105, 14);
            this.label4.TabIndex = 13;
            this.label4.Text = "Line Item Data";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblRespFilePath
            // 
            this.lblRespFilePath.AutoSize = true;
            this.lblRespFilePath.BackColor = System.Drawing.Color.Yellow;
            this.lblRespFilePath.ForeColor = System.Drawing.Color.DarkRed;
            this.lblRespFilePath.Location = new System.Drawing.Point(104, 101);
            this.lblRespFilePath.Name = "lblRespFilePath";
            this.lblRespFilePath.Size = new System.Drawing.Size(0, 14);
            this.lblRespFilePath.TabIndex = 16;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(661, 435);
            this.Controls.Add(this.lblRespFilePath);
            this.Controls.Add(this.btnLineItem);
            this.Controls.Add(this.txtLineItem);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtResponsePath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.lblDateTime);
            this.Controls.Add(this.lblUserName);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Verdana", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "COPILOT Submissin";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.Label lblDateTime;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtResponsePath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnLineItem;
        private System.Windows.Forms.TextBox txtLineItem;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblRespFilePath;
    }
}

