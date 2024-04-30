namespace ExcelQuery;

using NPOI;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.Crypt;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Linq;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
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

    public double QueryDouble() {
        return _qryDouble;
    }

    private string _qryString;
    private double _qryDouble;
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
            formWaiting.Show(); // Display form modelessly
            Application.DoEvents();

            string selectPath = folderBrowserDialog1.SelectedPath;
            _qryString = queryKeyWord.Text.Trim();
            if (IsNumeric(_qryString)) _qryDouble = Double.Parse(_qryString);
            msgOut.ResetText();
            _sbResult.Clear();

            if (queryKeyWord.Text.Trim() == "")
            {
                MessageBox.Show("�п�J�d������r");
                return;
            }

            if (selectPath == "")
            {
                MessageBox.Show("�п�ܸ�Ƨ�");
                return;
            }

            if (rdoMatchGreater.Checked || rdoMatchLess.Checked)
            {
                if (!IsNumeric(QueryString()))
                {
                    MessageBox.Show("�ȭ��Ʀr�i�ϥΤj��p��");
                    return;
                }
            }

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

                    string cellVal = GetCellValue(cell);
                    string cellCmn = GetCellComment(cell);
                    bool isMatchVal = false;
                    bool isMatchCmn = false;

                    if (cellVal == "" && cellCmn == "") continue;

                    if (rdoMatchEql.Checked) //�۵�
                    {
                        isMatchVal = MatchCell(cellVal);
                        isMatchCmn = MatchCell(cellCmn);
                    }
                    else if (rdoMatchGreater.Checked) //�j��
                    {
                        isMatchVal = MatchCellGreater(cellVal);
                        isMatchCmn = MatchCellGreater(cellCmn);
                    }
                    else //�p��
                    {
                        isMatchVal = MatchCellLess(cellVal);
                        isMatchCmn = MatchCellLess(cellCmn);
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

    public bool MatchCell(string txt)
    {
        return txt.IndexOf(QueryString()) > -1;
    }

    public bool MatchCellGreater(string txt)
    {
        if(!IsNumeric(txt)) return false;
        double val = Double.Parse(txt);
        return val > QueryDouble();
    }

    public bool MatchCellLess(string txt)
    {
        if (!IsNumeric(txt)) return false;
        double val = Double.Parse(txt);
        return val < QueryDouble();
    }

    private void Form1_FormClosed(object sender, FormClosedEventArgs e)
    {
        // �Ыب����Splash Screen
        FormSplashScreen splashScreen = new FormSplashScreen("Good Bye");
        splashScreen.ShowDialog();
        Application.Exit();
    }
}
