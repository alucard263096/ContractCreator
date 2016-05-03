using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContractCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            OpenFileDialog docFileDialog = new OpenFileDialog();
            docFileDialog.Title = "选择一个合同模板文件";
            docFileDialog.Filter = "Word文件 (*.docx)|*.docx|2003Word文件 (*.doc)|*.doc";
            docFileDialog.FilterIndex = 1;
            if (docFileDialog.ShowDialog() == DialogResult.OK)
            {
                OpenFileDialog excelFileDialog = new OpenFileDialog();
                excelFileDialog.Title = "选择一个合同数据文件";
                excelFileDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx|2003Excel文件 (*.xls)|*.xls";
                excelFileDialog.FilterIndex = 1;
                if (excelFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FolderBrowserDialog savePathDialog = new FolderBrowserDialog();
                    savePathDialog.Description = "选择合同生成的存放路径";
                    if (savePathDialog.ShowDialog() == DialogResult.OK)
                    {
                        RDoC r=new RDoC(docFileDialog.FileName, excelFileDialog.FileName, savePathDialog.SelectedPath,this);
                        Thread t = new Thread(new ThreadStart(r.rdo));
                        t.Start();  
                        //StartGenerateContract(docFileDialog.FileName, excelFileDialog.FileName, savePathDialog.SelectedPath);
                    }
                }
            }
        }

        class RDoC
        {
            string doc, exc, svp;
            Form1 f;
            public RDoC(string doc, string exc, string svp,Form1 f)
            {
                this.doc = doc;
                this.exc = exc;
                this.svp = svp;
                this.f = f;
            }
            public void rdo()
            {
                StartGenerateContract(doc, exc, svp);
            }
            private void StartGenerateContract(string docFilePath, string excelFilePath, string savePath)
            {
                bool canOpenWord = TryToOpenWord(docFilePath);
                if (canOpenWord == false)
                {
                    MessageBox.Show("读取合同模板文档失败，请检查Word文件或者Office环境是否正确。");
                    return;
                }
                DataTable dtExcel = ExcelToDT(excelFilePath);
                if (dtExcel == null)
                {
                    MessageBox.Show("读取合同数据文档失败，请检查Excel文件或者Office环境是否正确。");
                    return;
                }
                f.btnCreate.Enabled = false;
                f.pgb.Maximum = dtExcel.Rows.Count;
                f.pgb.Value = 0;
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    GenerateContract(docFilePath, dtExcel.Columns, dtExcel.Rows[i], savePath, getNo(i + 1, dtExcel.Rows.Count));
                    f.pgb.Value++;
                    f.lblProgress.Text = (i + 1).ToString() + "/" + dtExcel.Rows.Count.ToString();
                }
                f.lblProgress.Text = "生成完成";
                f.btnCreate.Enabled = true;
            }

            private string getNo(int no, int total)
            {
                StringBuilder sb = new StringBuilder();
                string nostr = no.ToString();
                string totalstr = total.ToString();
                for (int i = 0; i < totalstr.Length - nostr.Length; i++)
                {
                    sb.Append("0");
                }
                sb.Append(no);
                return sb.ToString();
            }

            private void GenerateContract(string docFilePath, DataColumnCollection columns, DataRow row, string savePath, string no)
            {
                Microsoft.Office.Interop.Word.Application app = null;
                Microsoft.Office.Interop.Word.Document doc = null;
                //将要导出的新word文件名
                try
                {
                    app = new Microsoft.Office.Interop.Word.Application();//创建word应用程序
                    object fileName = docFilePath;//模板文件
                    //打开模板文件
                    object oMissing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(ref fileName,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
                    //构造数据

                    foreach (DataColumn col in columns)
                    {
                        app.Selection.Find.Replacement.ClearFormatting();
                        app.Selection.Find.ClearFormatting();
                        app.Selection.Find.Text = "{" + col.ColumnName + "}";//需要被替换的文本
                        app.Selection.Find.Replacement.Text = row[col.ColumnName].ToString();//替换文本 

                        //执行替换操作
                        app.Selection.Find.Execute(
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref replace,
                        ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);
                    }

                    //对替换好的word模板另存为一个新的word文档
                    object newfileName = savePath + "\\" + no.ToString() + row[0].ToString() + ".docx";//模板文件

                    doc.SaveAs(newfileName,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (System.Threading.ThreadAbortException ex)
                {
                    //这边为了捕获Response.End引起的异常
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (doc != null)
                    {
                        doc.Close();//关闭word文档
                    }
                    if (app != null)
                    {
                        app.Quit();//退出word应用程序
                    }
                }
            }

            private bool TryToOpenWord(string docFilePath)
            {
                Microsoft.Office.Interop.Word.Application app = null;
                Microsoft.Office.Interop.Word.Document doc = null;
                //将要导出的新word文件名
                try
                {
                    app = new Microsoft.Office.Interop.Word.Application();//创建word应用程序
                    object fileName = docFilePath;//模板文件
                    //打开模板文件
                    object oMissing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(ref fileName,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    return true;
                }
                catch (System.Threading.ThreadAbortException ex)
                {
                    //这边为了捕获Response.End引起的异常
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (doc != null)
                    {
                        doc.Close();//关闭word文档
                    }
                    if (app != null)
                    {
                        app.Quit();//退出word应用程序
                    }
                }
                return false;
            }

            public DataTable ExcelToDT(string Path)
            {
                try
                {
                    string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + Path + ";" + @"Extended Properties=""Excel 12.0 Xml;HDR=YES;""";
                    using (OleDbConnection conn = new OleDbConnection(strConn))
                    {
                        conn.Open();
                        string strExcel = "";
                        OleDbDataAdapter myCommand = null;
                        DataSet ds = null;
                        strExcel = "select * from [sheet1$]";
                        myCommand = new OleDbDataAdapter(strExcel, strConn);
                        ds = new DataSet();
                        myCommand.Fill(ds, "table1");
                        return ds.Tables[0];
                    }
                }
                catch
                {

                }
                return null;
            }
        }

        private void btnDownDemo_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog savePathDialog = new FolderBrowserDialog();
            savePathDialog.Description = "选择模板合同的存放路径";
            if (savePathDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo doc=new FileInfo("Templates\\合同模板.docx");
                doc.CopyTo(savePathDialog.SelectedPath + "\\" + doc.Name,true);

                FileInfo excel = new FileInfo("Templates\\合同数据.xlsx");
                excel.CopyTo(savePathDialog.SelectedPath + "\\" + excel.Name, true);


                string path = savePathDialog.SelectedPath;
                System.Diagnostics.Process.Start("explorer.exe", path);
            }
        }
    }
}
