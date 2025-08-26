namespace ChiaseNoiBo
{
    partial class UpdateNotificationForm
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
            this.guna2MessageDialog1 = new Guna.UI2.WinForms.Guna2MessageDialog();
            this.lblMessage = new System.Windows.Forms.Label();
            this.guna2ProgressIndicator1 = new Guna.UI2.WinForms.Guna2ProgressIndicator();
            this.SuspendLayout();
            // 
            // guna2MessageDialog1
            // 
            this.guna2MessageDialog1.Buttons = Guna.UI2.WinForms.MessageDialogButtons.OK;
            this.guna2MessageDialog1.Caption = null;
            this.guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.None;
            this.guna2MessageDialog1.Parent = null;
            this.guna2MessageDialog1.Style = Guna.UI2.WinForms.MessageDialogStyle.Default;
            this.guna2MessageDialog1.Text = null;
            // 
            // lblMessage
            // 
            this.lblMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMessage.Location = new System.Drawing.Point(12, 102);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(492, 41);
            this.lblMessage.TabIndex = 0;
            // 
            // guna2ProgressIndicator1
            // 
            this.guna2ProgressIndicator1.Location = new System.Drawing.Point(217, -2);
            this.guna2ProgressIndicator1.Name = "guna2ProgressIndicator1";
            this.guna2ProgressIndicator1.ProgressColor = System.Drawing.Color.DeepSkyBlue;
            this.guna2ProgressIndicator1.Size = new System.Drawing.Size(90, 90);
            this.guna2ProgressIndicator1.TabIndex = 2;
            // 
            // UpdateNotificationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(516, 152);
            this.Controls.Add(this.guna2ProgressIndicator1);
            this.Controls.Add(this.lblMessage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "UpdateNotificationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "UpdateNotificationForm";
            this.Load += new System.EventHandler(this.UpdateNotificationForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private Guna.UI2.WinForms.Guna2MessageDialog guna2MessageDialog1;
        private System.Windows.Forms.Label lblMessage;
        private Guna.UI2.WinForms.Guna2ProgressIndicator guna2ProgressIndicator1;
    }
}