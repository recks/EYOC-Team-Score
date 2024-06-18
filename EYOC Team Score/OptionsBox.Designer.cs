namespace EYOC_Team_Score
{
    partial class OptionsBox
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
            ExportCSS = new CheckBox();
            OK = new Button();
            Cancel = new Button();
            SuspendLayout();
            // 
            // ExportCSS
            // 
            ExportCSS.AutoSize = true;
            ExportCSS.Location = new Point(18, 21);
            ExportCSS.Name = "ExportCSS";
            ExportCSS.Size = new Size(168, 19);
            ExportCSS.TabIndex = 0;
            ExportCSS.Text = "Export CSS with report files";
            ExportCSS.UseVisualStyleBackColor = true;
            ExportCSS.CheckedChanged += ExportCSS_CheckedChanged;
            // 
            // OK
            // 
            OK.Location = new Point(198, 81);
            OK.Name = "OK";
            OK.Size = new Size(75, 23);
            OK.TabIndex = 1;
            OK.Text = "OK";
            OK.UseVisualStyleBackColor = true;
            OK.Click += OK_Click;
            // 
            // Cancel
            // 
            Cancel.Location = new Point(287, 81);
            Cancel.Name = "Cancel";
            Cancel.Size = new Size(75, 23);
            Cancel.TabIndex = 2;
            Cancel.Text = "Cancel";
            Cancel.UseVisualStyleBackColor = true;
            Cancel.Click += Cancel_Click;
            // 
            // OptionsBox
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(374, 116);
            Controls.Add(Cancel);
            Controls.Add(OK);
            Controls.Add(ExportCSS);
            Name = "OptionsBox";
            Text = "Options";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private CheckBox ExportCSS;
        private Button OK;
        private Button Cancel;
    }
}