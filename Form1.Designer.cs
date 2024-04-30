namespace ExcelQuery;

partial class Form1
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
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnPickFolder = new System.Windows.Forms.Button();
            this.msgOut = new System.Windows.Forms.RichTextBox();
            this.queryKeyWord = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.filePw = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.btnQuery = new System.Windows.Forms.Button();
            this.rdoMatchEql = new System.Windows.Forms.RadioButton();
            this.rdoMatchGreater = new System.Windows.Forms.RadioButton();
            this.rdoMatchLess = new System.Windows.Forms.RadioButton();
            this.grpMatch = new System.Windows.Forms.GroupBox();
            this.grpMatch.SuspendLayout();
            this.SuspendLayout();
            this.Load += new System.EventHandler(this.Form1_Load);
            // 
            // btnPickFolder
            // 
            this.btnPickFolder.Location = new System.Drawing.Point(81, 12);
            this.btnPickFolder.Name = "btnPickFolder";
            this.btnPickFolder.Size = new System.Drawing.Size(78, 34);
            this.btnPickFolder.TabIndex = 0;
            this.btnPickFolder.Text = "選擇資料夾";
            this.btnPickFolder.UseVisualStyleBackColor = true;
            this.btnPickFolder.Click += new System.EventHandler(this.btnPickFolder_Click);
            // 
            // msgOut
            // 
            this.msgOut.Location = new System.Drawing.Point(240, 12);
            this.msgOut.Name = "msgOut";
            this.msgOut.ReadOnly = true;
            this.msgOut.Size = new System.Drawing.Size(548, 426);
            this.msgOut.TabIndex = 999;
            this.msgOut.Text = "";
            // 
            // queryKeyWord
            // 
            this.queryKeyWord.Location = new System.Drawing.Point(12, 100);
            this.queryKeyWord.Name = "queryKeyWord";
            this.queryKeyWord.Size = new System.Drawing.Size(220, 23);
            this.queryKeyWord.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "查詢關鍵字";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 357);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(159, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "檔案密碼(可換行設多個密碼)";
            // 
            // filePw
            // 
            this.filePw.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.filePw.Location = new System.Drawing.Point(13, 375);
            this.filePw.Multiline = true;
            this.filePw.Name = "filePw";
            this.filePw.Size = new System.Drawing.Size(220, 50);
            this.filePw.TabIndex = 1;
            // 
            // lblPath
            // 
            this.lblPath.AutoEllipsis = true;
            this.lblPath.Location = new System.Drawing.Point(10, 48);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(220, 15);
            this.lblPath.TabIndex = 6;
            this.lblPath.Text = "[請選擇資料夾]";
            this.lblPath.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnQuery
            // 
            this.btnQuery.Location = new System.Drawing.Point(88, 209);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(64, 34);
            this.btnQuery.TabIndex = 5;
            this.btnQuery.Text = "查詢";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // rdoMatchEql
            // 
            this.rdoMatchEql.AutoSize = true;
            this.rdoMatchEql.Checked = true;
            this.rdoMatchEql.Location = new System.Drawing.Point(12, 15);
            this.rdoMatchEql.Name = "rdoMatchEql";
            this.rdoMatchEql.Size = new System.Drawing.Size(49, 19);
            this.rdoMatchEql.TabIndex = 2;
            this.rdoMatchEql.TabStop = true;
            this.rdoMatchEql.Text = "相等";
            this.rdoMatchEql.UseVisualStyleBackColor = true;
            // 
            // rdoMatchGreater
            // 
            this.rdoMatchGreater.AutoSize = true;
            this.rdoMatchGreater.Location = new System.Drawing.Point(67, 15);
            this.rdoMatchGreater.Name = "rdoMatchGreater";
            this.rdoMatchGreater.Size = new System.Drawing.Size(49, 19);
            this.rdoMatchGreater.TabIndex = 3;
            this.rdoMatchGreater.Text = "大於";
            this.rdoMatchGreater.UseVisualStyleBackColor = true;
            // 
            // rdoMatchLess
            // 
            this.rdoMatchLess.AutoSize = true;
            this.rdoMatchLess.Location = new System.Drawing.Point(122, 15);
            this.rdoMatchLess.Name = "rdoMatchLess";
            this.rdoMatchLess.Size = new System.Drawing.Size(49, 19);
            this.rdoMatchLess.TabIndex = 4;
            this.rdoMatchLess.Text = "小於";
            this.rdoMatchLess.UseVisualStyleBackColor = true;
            // 
            // grpMatch
            // 
            this.grpMatch.Controls.Add(this.rdoMatchGreater);
            this.grpMatch.Controls.Add(this.rdoMatchLess);
            this.grpMatch.Controls.Add(this.rdoMatchEql);
            this.grpMatch.Location = new System.Drawing.Point(9, 148);
            this.grpMatch.Name = "grpMatch";
            this.grpMatch.Size = new System.Drawing.Size(202, 40);
            this.grpMatch.TabIndex = 10;
            this.grpMatch.TabStop = false;
            this.grpMatch.Text = "比對模式";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.grpMatch);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.filePw);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.queryKeyWord);
            this.Controls.Add(this.msgOut);
            this.Controls.Add(this.btnPickFolder);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel整批查詢";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.grpMatch.ResumeLayout(false);
            this.grpMatch.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private FolderBrowserDialog folderBrowserDialog1;
    private Button btnPickFolder;
    private RichTextBox msgOut;
    private TextBox queryKeyWord;
    private Label label1;
    private Label label2;
    private TextBox filePw;
    private Label lblPath;
    private Button btnQuery;
    private RadioButton rdoMatchEql;
    private RadioButton rdoMatchGreater;
    private RadioButton rdoMatchLess;
    private GroupBox grpMatch;
}
