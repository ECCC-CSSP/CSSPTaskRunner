namespace CSSPTaskRunner
{
    partial class CSSPTaskRunner
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.richTextBoxStatus = new System.Windows.Forms.RichTextBox();
            this.timerCheckTask = new System.Windows.Forms.Timer(this.components);
            this.panelStatus = new System.Windows.Forms.Panel();
            this.lblStatusText = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.butGetTasksInformation = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panelStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.butGetTasksInformation);
            this.splitContainer1.Panel1.Controls.Add(this.panelStatus);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.richTextBoxStatus);
            this.splitContainer1.Size = new System.Drawing.Size(800, 450);
            this.splitContainer1.SplitterDistance = 81;
            this.splitContainer1.TabIndex = 1;
            // 
            // richTextBoxStatus
            // 
            this.richTextBoxStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBoxStatus.Location = new System.Drawing.Point(0, 0);
            this.richTextBoxStatus.Name = "richTextBoxStatus";
            this.richTextBoxStatus.Size = new System.Drawing.Size(800, 365);
            this.richTextBoxStatus.TabIndex = 0;
            this.richTextBoxStatus.Text = "";
            // 
            // timerCheckTask
            // 
            this.timerCheckTask.Interval = 1000;
            this.timerCheckTask.Tick += new System.EventHandler(this.timerCheckTask_Tick);
            // 
            // panelStatus
            // 
            this.panelStatus.Controls.Add(this.lblStatusText);
            this.panelStatus.Controls.Add(this.lblStatus);
            this.panelStatus.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelStatus.Location = new System.Drawing.Point(0, 50);
            this.panelStatus.Name = "panelStatus";
            this.panelStatus.Size = new System.Drawing.Size(800, 31);
            this.panelStatus.TabIndex = 31;
            // 
            // lblStatusText
            // 
            this.lblStatusText.AutoSize = true;
            this.lblStatusText.Location = new System.Drawing.Point(12, 9);
            this.lblStatusText.Name = "lblStatusText";
            this.lblStatusText.Size = new System.Drawing.Size(40, 13);
            this.lblStatusText.TabIndex = 3;
            this.lblStatusText.Text = "Status:";
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(66, 9);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 13);
            this.lblStatus.TabIndex = 4;
            // 
            // butGetTasksInformation
            // 
            this.butGetTasksInformation.Location = new System.Drawing.Point(15, 12);
            this.butGetTasksInformation.Name = "butGetTasksInformation";
            this.butGetTasksInformation.Size = new System.Drawing.Size(165, 23);
            this.butGetTasksInformation.TabIndex = 32;
            this.butGetTasksInformation.Text = "Get Tasks Information";
            this.butGetTasksInformation.UseVisualStyleBackColor = true;
            this.butGetTasksInformation.Click += new System.EventHandler(this.butGetTasksInformation_Click);
            // 
            // CSSPTaskRunner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.splitContainer1);
            this.Name = "CSSPTaskRunner";
            this.Text = "CSSP Task Runner";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panelStatus.ResumeLayout(false);
            this.panelStatus.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.RichTextBox richTextBoxStatus;
        private System.Windows.Forms.Timer timerCheckTask;
        private System.Windows.Forms.Panel panelStatus;
        private System.Windows.Forms.Label lblStatusText;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button butGetTasksInformation;
    }
}

