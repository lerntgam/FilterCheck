using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace FilterCheck
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += Application_WorkbookBeforeSave;
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            if (activeWorksheet != null && activeWorksheet.FilterMode == true)
            {
                DialogResult ret;
                ret = MessageBox.Show("フィルターがかけられています。保存する前に確認してください。保存しますか？", "確認", buttons: MessageBoxButtons.OKCancel);
                if(ret == DialogResult.Cancel)
                {
                    Cancel = true;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
