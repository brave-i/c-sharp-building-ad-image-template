namespace BuildingTemplate
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
            m_progressBar = new System.Windows.Forms.ProgressBar();


            this.ExcelLabel = new System.Windows.Forms.Label();
            this.Excel_Path_Edit = new System.Windows.Forms.RichTextBox();
            this.Excel_Path_Btn = new System.Windows.Forms.Button();
            this.PROGRESS_LABEL = new System.Windows.Forms.Label();
            
            Label_Percent = new System.Windows.Forms.Label();
            this.m_startButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ExcelLabel
            // 
            this.ExcelLabel.AutoSize = true;
            this.ExcelLabel.Location = new System.Drawing.Point(11, 32);
            this.ExcelLabel.Name = "ExcelLabel";
            this.ExcelLabel.Size = new System.Drawing.Size(101, 13);
            this.ExcelLabel.TabIndex = 0;
            this.ExcelLabel.Text = "EXCEL FILE PATH:";
            // 
            // Excel_Path_Edit
            // 
            this.Excel_Path_Edit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.Excel_Path_Edit.Location = new System.Drawing.Point(118, 26);
            this.Excel_Path_Edit.Name = "Excel_Path_Edit";
            this.Excel_Path_Edit.ReadOnly = true;
            this.Excel_Path_Edit.Size = new System.Drawing.Size(432, 29);
            this.Excel_Path_Edit.TabIndex = 1;
            this.Excel_Path_Edit.Text = "Select Spread Sheet File!";
            // 
            // Excel_Path_Btn
            // 
            this.Excel_Path_Btn.Location = new System.Drawing.Point(556, 25);
            this.Excel_Path_Btn.Name = "Excel_Path_Btn";
            this.Excel_Path_Btn.Size = new System.Drawing.Size(36, 30);
            this.Excel_Path_Btn.TabIndex = 2;
            this.Excel_Path_Btn.Text = "...";
            this.Excel_Path_Btn.UseVisualStyleBackColor = true;
            this.Excel_Path_Btn.Click += new System.EventHandler(this.Excel_Path_Btn_Click);
            // 
            // PROGRESS_LABEL
            // 
            this.PROGRESS_LABEL.AutoSize = true;
            this.PROGRESS_LABEL.Location = new System.Drawing.Point(42, 78);
            this.PROGRESS_LABEL.Name = "PROGRESS_LABEL";
            this.PROGRESS_LABEL.Size = new System.Drawing.Size(70, 13);
            this.PROGRESS_LABEL.TabIndex = 0;
            this.PROGRESS_LABEL.Text = "PROGRESS:";
            // 
            // m_progressBar
            // 
            m_progressBar.Location = new System.Drawing.Point(117, 76);
            m_progressBar.Name = "m_progressBar";
            m_progressBar.Size = new System.Drawing.Size(433, 20);
            m_progressBar.TabIndex = 3;
            // 
            // Label_Percent
            // 
            Label_Percent.AutoSize = true;
            Label_Percent.Location = new System.Drawing.Point(561, 79);
            Label_Percent.Name = "Label_Percent";
            Label_Percent.Size = new System.Drawing.Size(24, 13);
            Label_Percent.TabIndex = 0;
            Label_Percent.Text = "0 %";
            // 
            // m_startButton
            // 
            this.m_startButton.Location = new System.Drawing.Point(408, 116);
            this.m_startButton.Name = "m_startButton";
            this.m_startButton.Size = new System.Drawing.Size(142, 37);
            this.m_startButton.TabIndex = 4;
            this.m_startButton.Text = "Go !";
            this.m_startButton.UseVisualStyleBackColor = true;
            this.m_startButton.Click += new System.EventHandler(this.m_startButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(612, 167);
            this.Controls.Add(this.m_startButton);
            this.Controls.Add(m_progressBar);
            this.Controls.Add(this.Excel_Path_Btn);
            this.Controls.Add(this.Excel_Path_Edit);
            this.Controls.Add(Label_Percent);
            this.Controls.Add(this.PROGRESS_LABEL);
            this.Controls.Add(this.ExcelLabel);
            this.Name = "Form1";
            this.Text = "Building Template";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label ExcelLabel;
        private System.Windows.Forms.RichTextBox Excel_Path_Edit;
        private System.Windows.Forms.Button Excel_Path_Btn;
        private System.Windows.Forms.Label PROGRESS_LABEL;
        private System.Windows.Forms.Button m_startButton;

        public static System.Windows.Forms.ProgressBar m_progressBar;
        public static System.Windows.Forms.Label Label_Percent;
    }
}

