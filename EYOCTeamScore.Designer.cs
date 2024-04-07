﻿using TheArtOfDev.HtmlRenderer.WinForms;

namespace EYOC_Team_Score
{
    partial class EYOCTeamScore
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EYOCTeamScore));
            btn_ImportFile = new Button();
            groupBox1 = new GroupBox();
            btn_CalculateTeamScores = new Button();
            listbox_ResultFiles = new ListBox();
            eventBindingSource = new BindingSource(components);
            openResultFileDialog = new OpenFileDialog();
            lbl_TeamScores = new Label();
            tabs_TeamScore = new TabControl();
            Total = new TabPage();
            htmlPanel_Total = new HtmlPanel();
            IndividualPage = new TabPage();
            htmlPanel_Individual = new HtmlPanel();
            ctxmenu_DeleteResultFile = new ContextMenuStrip(components);
            mitem_DeleteResultFile = new ToolStripMenuItem();
            menuTop = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            menuTop_OpenFile = new ToolStripMenuItem();
            toolStripSeparator = new ToolStripSeparator();
            menuTop_Export = new ToolStripMenuItem();
            toolStripSeparator1 = new ToolStripSeparator();
            printToolStripMenuItem = new ToolStripMenuItem();
            printPreviewToolStripMenuItem = new ToolStripMenuItem();
            toolStripSeparator2 = new ToolStripSeparator();
            menuTop_Exit = new ToolStripMenuItem();
            toolsToolStripMenuItem = new ToolStripMenuItem();
            menuTop_Options = new ToolStripMenuItem();
            helpToolStripMenuItem = new ToolStripMenuItem();
            toolStripSeparator5 = new ToolStripSeparator();
            menuTop_about = new ToolStripMenuItem();
            exportTeamScoresIndividualDialog = new SaveFileDialog();
            exportTeamScoresTotalDialog = new SaveFileDialog();
            groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)eventBindingSource).BeginInit();
            tabs_TeamScore.SuspendLayout();
            Total.SuspendLayout();
            IndividualPage.SuspendLayout();
            ctxmenu_DeleteResultFile.SuspendLayout();
            menuTop.SuspendLayout();
            SuspendLayout();
            // 
            // btn_ImportFile
            // 
            btn_ImportFile.AccessibleRole = AccessibleRole.None;
            btn_ImportFile.Location = new Point(6, 22);
            btn_ImportFile.Name = "btn_ImportFile";
            btn_ImportFile.Size = new Size(142, 23);
            btn_ImportFile.TabIndex = 1;
            btn_ImportFile.Text = "Import Result File";
            btn_ImportFile.UseVisualStyleBackColor = true;
            btn_ImportFile.Click += ImportFile_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btn_CalculateTeamScores);
            groupBox1.Controls.Add(listbox_ResultFiles);
            groupBox1.Controls.Add(btn_ImportFile);
            groupBox1.Location = new Point(12, 37);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(776, 97);
            groupBox1.TabIndex = 2;
            groupBox1.TabStop = false;
            groupBox1.Text = "Import results";
            // 
            // btn_CalculateTeamScores
            // 
            btn_CalculateTeamScores.Font = new Font("Segoe UI", 9F);
            btn_CalculateTeamScores.Location = new Point(6, 63);
            btn_CalculateTeamScores.Name = "btn_CalculateTeamScores";
            btn_CalculateTeamScores.Size = new Size(142, 23);
            btn_CalculateTeamScores.TabIndex = 3;
            btn_CalculateTeamScores.Text = "Calculate Team Scores";
            btn_CalculateTeamScores.UseVisualStyleBackColor = true;
            btn_CalculateTeamScores.Click += btn_CalculateTeamScores_Click;
            // 
            // listbox_ResultFiles
            // 
            listbox_ResultFiles.DataSource = eventBindingSource;
            listbox_ResultFiles.FormattingEnabled = true;
            listbox_ResultFiles.ItemHeight = 15;
            listbox_ResultFiles.Location = new Point(163, 22);
            listbox_ResultFiles.Name = "listbox_ResultFiles";
            listbox_ResultFiles.Size = new Size(607, 64);
            listbox_ResultFiles.TabIndex = 2;
            listbox_ResultFiles.MouseUp += listbox_ResultFiles_MouseUp;
            // 
            // eventBindingSource
            // 
            eventBindingSource.DataSource = typeof(Model.Event);
            // 
            // openResultFileDialog
            // 
            openResultFileDialog.Filter = "\"XML files|*.xml|All files|*.*\"";
            // 
            // lbl_TeamScores
            // 
            lbl_TeamScores.AutoSize = true;
            lbl_TeamScores.Font = new Font("Segoe UI", 14F);
            lbl_TeamScores.Location = new Point(355, 133);
            lbl_TeamScores.Name = "lbl_TeamScores";
            lbl_TeamScores.Size = new Size(116, 25);
            lbl_TeamScores.TabIndex = 4;
            lbl_TeamScores.Text = "Team Scores";
            // 
            // tabs_TeamScore
            // 
            tabs_TeamScore.Controls.Add(Total);
            tabs_TeamScore.Controls.Add(IndividualPage);
            tabs_TeamScore.Location = new Point(18, 146);
            tabs_TeamScore.Name = "tabs_TeamScore";
            tabs_TeamScore.SelectedIndex = 0;
            tabs_TeamScore.Size = new Size(770, 573);
            tabs_TeamScore.TabIndex = 5;
            // 
            // Total
            // 
            Total.Controls.Add(htmlPanel_Total);
            Total.Location = new Point(4, 24);
            Total.Name = "Total";
            Total.Padding = new Padding(3);
            Total.Size = new Size(762, 545);
            Total.TabIndex = 0;
            Total.Text = "Total";
            Total.UseVisualStyleBackColor = true;
            // 
            // htmlPanel_Total
            // 
            htmlPanel_Total.AutoScroll = true;
            htmlPanel_Total.BackColor = SystemColors.Window;
            htmlPanel_Total.BaseStylesheet = null;
            htmlPanel_Total.Dock = DockStyle.Fill;
            htmlPanel_Total.Location = new Point(3, 3);
            htmlPanel_Total.Name = "htmlPanel_Total";
            htmlPanel_Total.Size = new Size(756, 539);
            htmlPanel_Total.TabIndex = 6;
            htmlPanel_Total.Text = null;
            // 
            // IndividualPage
            // 
            IndividualPage.Controls.Add(htmlPanel_Individual);
            IndividualPage.Location = new Point(4, 24);
            IndividualPage.Name = "IndividualPage";
            IndividualPage.Padding = new Padding(3);
            IndividualPage.Size = new Size(762, 545);
            IndividualPage.TabIndex = 1;
            IndividualPage.Text = "Detailed";
            IndividualPage.UseVisualStyleBackColor = true;
            // 
            // htmlPanel_Individual
            // 
            htmlPanel_Individual.AutoScroll = true;
            htmlPanel_Individual.BackColor = SystemColors.Window;
            htmlPanel_Individual.BaseStylesheet = null;
            htmlPanel_Individual.Dock = DockStyle.Fill;
            htmlPanel_Individual.Location = new Point(3, 3);
            htmlPanel_Individual.Name = "htmlPanel_Individual";
            htmlPanel_Individual.Size = new Size(756, 539);
            htmlPanel_Individual.TabIndex = 6;
            htmlPanel_Individual.Text = null;
            // 
            // ctxmenu_DeleteResultFile
            // 
            ctxmenu_DeleteResultFile.Items.AddRange(new ToolStripItem[] { mitem_DeleteResultFile });
            ctxmenu_DeleteResultFile.Name = "ctxmenu_DeleteResultFile";
            ctxmenu_DeleteResultFile.Size = new Size(181, 26);
            // 
            // mitem_DeleteResultFile
            // 
            mitem_DeleteResultFile.Name = "mitem_DeleteResultFile";
            mitem_DeleteResultFile.Size = new Size(180, 22);
            mitem_DeleteResultFile.Text = "Delete this result file";
            mitem_DeleteResultFile.Click += ctxmenu_DeleteResultFile_Click;
            // 
            // menuTop
            // 
            menuTop.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, toolsToolStripMenuItem, helpToolStripMenuItem });
            menuTop.Location = new Point(0, 0);
            menuTop.Name = "menuTop";
            menuTop.Size = new Size(800, 24);
            menuTop.TabIndex = 6;
            menuTop.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { menuTop_OpenFile, toolStripSeparator, menuTop_Export, toolStripSeparator1, printToolStripMenuItem, printPreviewToolStripMenuItem, toolStripSeparator2, menuTop_Exit });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // menuTop_OpenFile
            // 
            menuTop_OpenFile.Image = (Image)resources.GetObject("menuTop_OpenFile.Image");
            menuTop_OpenFile.ImageTransparentColor = Color.Magenta;
            menuTop_OpenFile.Name = "menuTop_OpenFile";
            menuTop_OpenFile.ShortcutKeys = Keys.Control | Keys.O;
            menuTop_OpenFile.Size = new Size(209, 22);
            menuTop_OpenFile.Text = "&Import Result File";
            menuTop_OpenFile.Click += ImportFile_Click;
            // 
            // toolStripSeparator
            // 
            toolStripSeparator.Name = "toolStripSeparator";
            toolStripSeparator.Size = new Size(206, 6);
            // 
            // menuTop_Export
            // 
            menuTop_Export.Image = (Image)resources.GetObject("menuTop_Export.Image");
            menuTop_Export.ImageTransparentColor = Color.Magenta;
            menuTop_Export.Name = "menuTop_Export";
            menuTop_Export.ShortcutKeys = Keys.Control | Keys.S;
            menuTop_Export.Size = new Size(209, 22);
            menuTop_Export.Text = "&Export Reports";
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(206, 6);
            // 
            // printToolStripMenuItem
            // 
            printToolStripMenuItem.Enabled = false;
            printToolStripMenuItem.Image = (Image)resources.GetObject("printToolStripMenuItem.Image");
            printToolStripMenuItem.ImageTransparentColor = Color.Magenta;
            printToolStripMenuItem.Name = "printToolStripMenuItem";
            printToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.P;
            printToolStripMenuItem.Size = new Size(209, 22);
            printToolStripMenuItem.Text = "&Print";
            // 
            // printPreviewToolStripMenuItem
            // 
            printPreviewToolStripMenuItem.Enabled = false;
            printPreviewToolStripMenuItem.Image = (Image)resources.GetObject("printPreviewToolStripMenuItem.Image");
            printPreviewToolStripMenuItem.ImageTransparentColor = Color.Magenta;
            printPreviewToolStripMenuItem.Name = "printPreviewToolStripMenuItem";
            printPreviewToolStripMenuItem.Size = new Size(209, 22);
            printPreviewToolStripMenuItem.Text = "Print Pre&view";
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new Size(206, 6);
            // 
            // menuTop_Exit
            // 
            menuTop_Exit.Name = "menuTop_Exit";
            menuTop_Exit.Size = new Size(209, 22);
            menuTop_Exit.Text = "E&xit";
            // 
            // toolsToolStripMenuItem
            // 
            toolsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { menuTop_Options });
            toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            toolsToolStripMenuItem.Size = new Size(46, 20);
            toolsToolStripMenuItem.Text = "&Tools";
            // 
            // menuTop_Options
            // 
            menuTop_Options.Enabled = false;
            menuTop_Options.Name = "menuTop_Options";
            menuTop_Options.Size = new Size(116, 22);
            menuTop_Options.Text = "&Options";
            // 
            // helpToolStripMenuItem
            // 
            helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { toolStripSeparator5, menuTop_about });
            helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            helpToolStripMenuItem.Size = new Size(44, 20);
            helpToolStripMenuItem.Text = "&Help";
            // 
            // toolStripSeparator5
            // 
            toolStripSeparator5.Name = "toolStripSeparator5";
            toolStripSeparator5.Size = new Size(113, 6);
            // 
            // menuTop_about
            // 
            menuTop_about.Name = "menuTop_about";
            menuTop_about.Size = new Size(116, 22);
            menuTop_about.Text = "&About...";
            // 
            // exportTeamScoresIndividualDialog
            // 
            exportTeamScoresIndividualDialog.FileName = "teamscores_details.html";
            exportTeamScoresIndividualDialog.Filter = "\"HTML files|*.html\"";
            exportTeamScoresIndividualDialog.FileOk += exportIndividual_Ok;
            // 
            // exportTeamScoresTotalDialog
            // 
            exportTeamScoresTotalDialog.FileName = "teamscores_total.html";
            exportTeamScoresTotalDialog.Filter = "\"HTML files|*.html\"";
            exportTeamScoresTotalDialog.FileOk += exportTotal_Ok;
            // 
            // EYOCTeamScore
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 731);
            Controls.Add(menuTop);
            Controls.Add(lbl_TeamScores);
            Controls.Add(tabs_TeamScore);
            Controls.Add(groupBox1);
            Name = "EYOCTeamScore";
            Text = "EYOC Team Score Calculator";
            groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)eventBindingSource).EndInit();
            tabs_TeamScore.ResumeLayout(false);
            Total.ResumeLayout(false);
            IndividualPage.ResumeLayout(false);
            ctxmenu_DeleteResultFile.ResumeLayout(false);
            menuTop.ResumeLayout(false);
            menuTop.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btn_ImportFile;
        private GroupBox groupBox1;
        private ListBox listbox_ResultFiles;
        private Button btn_CalculateTeamScores;
        private OpenFileDialog openResultFileDialog;
        private Label lbl_TeamScores;
        private TabControl tabs_TeamScore;
        private TabPage Total;
        private TabPage IndividualPage;
        private BindingSource eventBindingSource;
        private ContextMenuStrip ctxmenu_DeleteResultFile;
        private ToolStripMenuItem mitem_DeleteResultFile;
        private HtmlPanel htmlPanel_Individual;
        private HtmlPanel htmlPanel_Total;
        private MenuStrip menuTop;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem menuTop_OpenFile;
        private ToolStripSeparator toolStripSeparator;
        private ToolStripMenuItem menuTop_Export;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem printToolStripMenuItem;
        private ToolStripMenuItem printPreviewToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripMenuItem menuTop_Exit;
        private ToolStripMenuItem toolsToolStripMenuItem;
        private ToolStripMenuItem menuTop_Options;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripMenuItem menuTop_about;
        private SaveFileDialog exportTeamScoresIndividualDialog;
        private SaveFileDialog exportTeamScoresTotalDialog;
    }
}
