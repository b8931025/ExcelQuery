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
            this.label2 = new System.Windows.Forms.Label();
            this.filePw = new System.Windows.Forms.TextBox();
            this.lblPath = new System.Windows.Forms.Label();
            this.btnQuery = new System.Windows.Forms.Button();
            this.numBgn = new System.Windows.Forms.NumericUpDown();
            this.numEnd = new System.Windows.Forms.NumericUpDown();
            this.lblBetween = new System.Windows.Forms.Label();
            this.rdoKeyWord = new System.Windows.Forms.RadioButton();
            this.rdoNumberRange = new System.Windows.Forms.RadioButton();
            this.grpQueryType = new System.Windows.Forms.GroupBox();
            this.chkNumberInWord = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.numBgn)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numEnd)).BeginInit();
            this.grpQueryType.SuspendLayout();
            this.SuspendLayout();
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
            this.queryKeyWord.Location = new System.Drawing.Point(12, 130);
            this.queryKeyWord.Name = "queryKeyWord";
            this.queryKeyWord.Size = new System.Drawing.Size(220, 23);
            this.queryKeyWord.TabIndex = 1;
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
            this.filePw.PasswordChar = '*';
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
            this.btnQuery.Location = new System.Drawing.Point(81, 230);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(64, 34);
            this.btnQuery.TabIndex = 5;
            this.btnQuery.Text = "查詢";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // numBgn
            // 
            this.numBgn.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numBgn.Location = new System.Drawing.Point(12, 170);
            this.numBgn.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numBgn.Minimum = new decimal(new int[] {
            1000000,
            0,
            0,
            -2147483648});
            this.numBgn.Name = "numBgn";
            this.numBgn.Size = new System.Drawing.Size(85, 23);
            this.numBgn.TabIndex = 1000;
            this.numBgn.Visible = false;
            // 
            // numEnd
            // 
            this.numEnd.Increment = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numEnd.Location = new System.Drawing.Point(120, 170);
            this.numEnd.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numEnd.Minimum = new decimal(new int[] {
            1000000,
            0,
            0,
            -2147483648});
            this.numEnd.Name = "numEnd";
            this.numEnd.Size = new System.Drawing.Size(85, 23);
            this.numEnd.TabIndex = 1001;
            this.numEnd.Value = new decimal(new int[] {
            2000,
            0,
            0,
            0});
            this.numEnd.Visible = false;
            // 
            // lblBetween
            // 
            this.lblBetween.AutoSize = true;
            this.lblBetween.Location = new System.Drawing.Point(98, 175);
            this.lblBetween.Name = "lblBetween";
            this.lblBetween.Size = new System.Drawing.Size(19, 15);
            this.lblBetween.TabIndex = 1003;
            this.lblBetween.Text = "～";
            this.lblBetween.Visible = false;
            // 
            // rdoKeyWord
            // 
            this.rdoKeyWord.AutoSize = true;
            this.rdoKeyWord.Checked = true;
            this.rdoKeyWord.Location = new System.Drawing.Point(0, 15);
            this.rdoKeyWord.Name = "rdoKeyWord";
            this.rdoKeyWord.Size = new System.Drawing.Size(61, 19);
            this.rdoKeyWord.TabIndex = 1004;
            this.rdoKeyWord.TabStop = true;
            this.rdoKeyWord.Text = "關鍵字";
            this.rdoKeyWord.UseVisualStyleBackColor = true;
            this.rdoKeyWord.Click += new System.EventHandler(this.rdoKeyWord_Click);
            // 
            // rdoNumberRange
            // 
            this.rdoNumberRange.AutoSize = true;
            this.rdoNumberRange.Location = new System.Drawing.Point(94, 15);
            this.rdoNumberRange.Name = "rdoNumberRange";
            this.rdoNumberRange.Size = new System.Drawing.Size(73, 19);
            this.rdoNumberRange.TabIndex = 1005;
            this.rdoNumberRange.Text = "數字範圍";
            this.rdoNumberRange.UseVisualStyleBackColor = true;
            this.rdoNumberRange.Click += new System.EventHandler(this.rdoNumberRange_Click);
            // 
            // grpQueryType
            // 
            this.grpQueryType.Controls.Add(this.rdoNumberRange);
            this.grpQueryType.Controls.Add(this.rdoKeyWord);
            this.grpQueryType.Location = new System.Drawing.Point(13, 70);
            this.grpQueryType.Name = "grpQueryType";
            this.grpQueryType.Size = new System.Drawing.Size(200, 40);
            this.grpQueryType.TabIndex = 1006;
            this.grpQueryType.TabStop = false;
            this.grpQueryType.Text = "查詢類別";
            // 
            // chkNumberInWord
            // 
            this.chkNumberInWord.AutoSize = true;
            this.chkNumberInWord.Location = new System.Drawing.Point(12, 200);
            this.chkNumberInWord.Name = "chkNumberInWord";
            this.chkNumberInWord.Size = new System.Drawing.Size(98, 19);
            this.chkNumberInWord.TabIndex = 1007;
            this.chkNumberInWord.Text = "混合數字查詢";
            this.chkNumberInWord.UseVisualStyleBackColor = true;
            this.chkNumberInWord.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.chkNumberInWord);
            this.Controls.Add(this.grpQueryType);
            this.Controls.Add(this.lblBetween);
            this.Controls.Add(this.numEnd);
            this.Controls.Add(this.numBgn);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.filePw);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.queryKeyWord);
            this.Controls.Add(this.msgOut);
            this.Controls.Add(this.btnPickFolder);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel整批查詢";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numBgn)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numEnd)).EndInit();
            this.grpQueryType.ResumeLayout(false);
            this.grpQueryType.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private FolderBrowserDialog folderBrowserDialog1;
    private Button btnPickFolder;
    private RichTextBox msgOut;
    private TextBox queryKeyWord;
    private Label label2;
    private TextBox filePw;
    private Label lblPath;
    private Button btnQuery;
    private NumericUpDown numBgn;
    private NumericUpDown numEnd;
    private Label lblBetween;
    private RadioButton rdoKeyWord;
    private RadioButton rdoNumberRange;
    private GroupBox grpQueryType;
    private CheckBox chkNumberInWord;
}
