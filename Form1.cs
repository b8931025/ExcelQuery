namespace ExcelQuery;

using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
        numBgn.Location = new System.Drawing.Point(numBgn.Location.X, numBgn.Location.Y - 40);
        numEnd.Location = new System.Drawing.Point(numEnd.Location.X, numEnd.Location.Y - 40);
        lblBetween.Location = new System.Drawing.Point(lblBetween.Location.X, lblBetween.Location.Y - 40);
        chkNumberInWord.Location = new System.Drawing.Point(chkNumberInWord.Location.X, chkNumberInWord.Location.Y - 40);
        btnQuery.Location = new System.Drawing.Point(btnQuery.Location.X, btnQuery.Location.Y - 40);
    }

    public string QueryString()
    {
        return _qryString;
    }

    private void Form1_Load(object sender, EventArgs e)
    {
        // �Ыب����Splash Screen
        FormSplashScreen splashScreen = new FormSplashScreen();
        splashScreen.ShowDialog();
    }

    private string _qryString;
    private decimal _numBgn;
    private decimal _numEnd;
    StringBuilder _sbResult = new StringBuilder();

    /// <summary>
    /// �ҥ�vb��isNumeric
    /// </summary>
    /// <param name="Expression">Object</param>
    /// <returns>true:�O�Ʀr false:�D�Ʀr</returns>
    static bool IsNumeric(object Expression)
    {
        // Variable to collect the Return value of the TryParse method.
        bool isNum;

        // Define variable to collect out parameter of the TryParse method. If the conversion fails, the out parameter is zero.
        double retNum;

        // The TryParse method converts a string in a specified style and culture-specific format to its double-precision floating point number equivalent.
        // The TryParse method does not generate an exception if the conversion fails. If the conversion passes, True is returned. If it does not, False is returned.
        isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);

        return isNum;
    }

    /// <summary>
    /// ��ܱ��d�߸�Ƨ�
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void btnPickFolder_Click(object sender, EventArgs e)
    {
        var ret = folderBrowserDialog1.ShowDialog();
        if (ret != DialogResult.OK) return;
        lblPath.Text = folderBrowserDialog1.SelectedPath;
    }

    /// <summary>
    /// ��Ҧ�excel�ɧ@����j��
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private async void btnQuery_Click(object sender, EventArgs e)
    {
        FormWaiting formWaiting = new FormWaiting();

        try
        {
            string selectPath = folderBrowserDialog1.SelectedPath;
            _qryString = queryKeyWord.Text.Trim();
            _numBgn = numBgn.Value;
            _numEnd = numEnd.Value;

            msgOut.ResetText();
            _sbResult.Clear();

            if (selectPath == "")
            {
                MessageBox.Show("�п�ܸ�Ƨ�");
                return;
            }

            if (rdoKeyWord.Checked && queryKeyWord.Text.Trim() == "")
            {
                MessageBox.Show("�п�J�d������r");
                return;
            }

            if(rdoNumberRange.Checked && (string.IsNullOrEmpty(numBgn.Text) || string.IsNullOrEmpty(numEnd.Text)))
            {
                MessageBox.Show("�п�J�Ʀr�d��");
                return;
            }else if(rdoNumberRange.Checked && _numBgn > _numEnd)
            {
                MessageBox.Show("�_�l�Ȥ��i�j�󵲧���");
                return;
            }

            // Display form modelessly
            formWaiting.Show(); 
            Application.DoEvents();

            //Recursively search for all files in the selected folder
            var files = Directory.GetFiles(selectPath, "*.xls*", SearchOption.AllDirectories);
            Dictionary<string, IWorkbook> workbooks = new Dictionary<string, IWorkbook>();

            foreach (var filePath in files)
            {
                //Check if the file is an Excel file
                if (filePath.EndsWith(".xls") || filePath.EndsWith(".xlsx"))
                {
                    IWorkbook workbook = GetWorkbook(filePath);
                    if (workbook != null) workbooks.Add(filePath, workbook);
                }
            }

            if (msgOut.TextLength > 0) msgOut.AppendText("\r\n============[�d�ߵ��G]============\r\n\r\n");

            foreach (var obj in workbooks)
            {
                string filePath = obj.Key;
                IWorkbook workbook = obj.Value;
                SearchExcelFile(filePath, workbook);
            }

            if (_sbResult.Length == 0)
            {
                MsgOut("�d�L���");
            }
            else
            {
                MsgOut(_sbResult.ToString());
            }
        }
        catch(Exception ex)
        {

        }
        finally
        {
            formWaiting.Close();
        }

        MessageBox.Show("�d�ߵ���");
    }

    public void SearchExcelFile(string path, IWorkbook workbook)
    {
        int numberOfSheets = workbook.NumberOfSheets;

        for(int sIdx = 0; sIdx < numberOfSheets; sIdx++) // Loop through sheets
        {
            ISheet sheet = workbook.GetSheetAt(sIdx);

            for (int rowIdx = 0; rowIdx <= sheet.LastRowNum; rowIdx++) // Loop through rows
            {
                IRow row = sheet.GetRow(rowIdx);
                if (row == null) continue;

                for (int cellIdx = 0; cellIdx < row.LastCellNum; cellIdx++) // Loop through columns
                {
                    ICell cell = row.GetCell(cellIdx);
                    if (cell == null) continue;
                    //���o��椺��
                    string cellVal = GetCellValue(cell);
                    //���o���`��
                    string cellCmn = GetCellComment(cell);
                    //��椺��O�_�ŦX����
                    bool isMatchVal = false;
                    //���`�ѬO�_�ŦX����
                    bool isMatchCmn = false;

                    if (cellVal == "" && cellCmn == "") continue;

                    //����r�d��
                    if (rdoKeyWord.Checked)
                    {
                        isMatchVal = MatchKeyWord(cellVal);
                        isMatchCmn = MatchKeyWord(cellCmn);
                    }
                    else
                    {
                        //�Ʀr�d��d��
                        isMatchVal = MatchNumberRange(cellVal);
                        isMatchCmn = MatchNumberRange(cellCmn);
                    }

                    if (isMatchVal || isMatchCmn) {
                        _sbResult.Append($"[{path}][{sheet.SheetName}:{cell.Address}] >> ");
                        if (isMatchVal) _sbResult.Append(cellVal);
                        if (isMatchCmn) _sbResult.Append($"[����]{cellCmn}");
                        _sbResult.Append(Environment.NewLine);
                    }
                }
            }
        }
    }

    public string[] GetPasswords()
    {
        string[] pwList = filePw.Text.Split(Environment.NewLine);
        pwList = pwList.Where(pw => pw.Trim() != "").Select(pw => pw.Trim()).ToArray();
        return pwList;
    }

    public void MsgOut(string msg)
    {
        msgOut.Text += msg + Environment.NewLine;
    }

    public IWorkbook GetWorkbook(string path)
    {
        IWorkbook workbook = null;
        bool isXlsx = Path.GetExtension(path).Equals(".xlsx", StringComparison.OrdinalIgnoreCase);

        try
        {
            workbook = WorkbookFactory.Create(path, "", true);
            return workbook;
        }
        catch (Exception ex)
        {
            if (ex.Message.IndexOf("used by another process") > -1) {
                MsgOut($"�ɮפw�}��[{path}]");
                return null;
            }
            //if (ex.Message.IndexOf("not match") == -1) throw ex;
            string msg = ex.Message;
        }

        //�ݭn�K�X
        string[] list = GetPasswords();
        if (list.Length == 0) { 
            MsgOut($"�L�k�}�Ҫ��ɮ�[{path}]");
            return null;
        }

        foreach (string pw in list)
        {
            try
            {
                workbook = WorkbookFactory.Create(path, pw, true);
                return workbook;
            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
        }
 
        MsgOut($"�L�k�ѱK���ɮ�[{path}]");
        return workbook;
    }

    public string GetCellValue(ICell cell)
    {
        if (cell == null) return "";
        IComment comment = cell.CellComment;
        string ret = "";

        //�d������
        switch (cell.CellType)
        {
            case CellType.String:
                ret = cell.StringCellValue;
                Debug.WriteLine(cell.StringCellValue);
                break;
            case CellType.Numeric:
                if (DateUtil.IsCellDateFormatted(cell))
                {
                    ret = cell.DateCellValue.ToString();
                }
                else
                {
                    ret = cell.NumericCellValue.ToString();
                }
                break;
            case CellType.Boolean:
                ret = cell.BooleanCellValue.ToString();
                break;
            case CellType.Formula:
                if (cell.CachedFormulaResultType == CellType.Numeric)
                {
                    ret = cell.NumericCellValue.ToString();
                }
                else if (cell.CachedFormulaResultType == CellType.String)
                {
                    ret = cell.StringCellValue;
                }
                break;
            default:
                Debug.WriteLine($"{cell.Address} is unknown cell type");
                break;
        }

        return ret;
    }

    //���o����
    public string GetCellComment(ICell cell)
    {
        /*ps:
         * ����쪺���ѧ䤣��
         */
        if (cell == null) return "";
        IComment comment = cell.CellComment;
        if (comment == null) return "";
        return comment.String.ToString();
    }

    public bool MatchKeyWord(string txt)
    {
        if (string.IsNullOrEmpty(txt)) return false;
        return txt.IndexOf(QueryString()) > -1;
    }

    public bool MatchNumberRange(string txt)
    {
        if (string.IsNullOrEmpty(txt)) return false;

        Regex r = new Regex("\\d+");
        Match m = r.Match(txt);
        decimal d = 0;
        bool result = false;

        if(chkNumberInWord.Checked)
        {
            //�d�߲V�b��r�����Ʀr ex:����1000���
            while (m.Success)
            {
                for (int g = 0; g < m.Groups.Count; g++)
                {
                    Group group = m.Groups[g];
                    for (int c = 0; c < group.Captures.Count; c++)
                    {
                        Capture capture = group.Captures[c];
                        d = decimal.Parse(capture.ToString());
                        result = _numBgn <= d && _numEnd >= d;
                        if (result) break;
                    }
                    if (result) break;
                }
                if (result) break;
                m = m.NextMatch();
            }
        }
        else
        {
            //�u�d�߯¼Ʀr
            if (!IsNumeric(txt)) return false;
            d = Convert.ToDecimal(txt);
            result = _numBgn <= d && _numEnd >= d;
        }
        return result;
    }


    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
        // �Ыب����Splash Screen
        FormSplashScreen splashScreen = new FormSplashScreen("Good Bye");
        splashScreen.ShowDialog();
        Application.Exit();
    }

    private void rdoKeyWord_Click(object sender, EventArgs e)
    {
        lblBetween.Visible = false;
        numBgn.Visible = false;
        numEnd.Visible = false;
        chkNumberInWord.Visible = false;
        queryKeyWord.Visible = true;
    }

    private void rdoNumberRange_Click(object sender, EventArgs e)
    {
        lblBetween.Visible = true;
        numBgn.Visible = true;
        numEnd.Visible = true;
        chkNumberInWord.Visible = true;
        queryKeyWord.Visible = false;
    }
}
