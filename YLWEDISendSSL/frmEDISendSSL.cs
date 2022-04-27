using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using YLWService;
using YLWService.Extensions;

namespace YLWEDISendSSL
{
    public partial class frmEDISendSSL : Form
    {
        bool _bEvent = false;

        string selectedPath = "";
        string attachPath = "";

        public frmEDISendSSL()
        {
            InitializeComponent();

            SetInit();

            _bEvent = true;
        }

        private void SetInit()
        {
            DateTime today = DateTime.Today;
            dtpFrDt.Value = today.AddDays(1 - today.Day);
            dtpToDt.Value = dtpFrDt.Value.AddDays(17);

            selectedPath = YLWServiceModule.GetOutPath();
            attachPath = YLWServiceModule.GetSendfilePath();

            this.dgvList.AutoGenerateColumns = false;
            this.dgvList.AddColumn("TEXTBOX", "dsgnNo", "설계번호", 100);
            this.dgvList.AddColumn("TEXTBOX", "contrNm", "계약자명", 100);
            this.dgvList.AddColumn("TEXTBOX", "trDt", "조사요청일자", 80);
            this.dgvList.AddColumn("TEXTBOX", "CclsDt", "종결일자", 80);
            this.dgvList.AddColumn("TEXTBOX", "AsgnEmpSeq", "조사자내부코드", 10, false);
            this.dgvList.AddColumn("TEXTBOX", "AsgnEmpName", "조사자명", 100);
            this.dgvList.AddColumn("TEXTBOX", "req_id", "조사요청내부코드", 10, false);
            this.dgvList.AddColumn("TEXTBOX", "rtn_id", "조사결과내부코드", 10, false);
            this.dgvList.AddColumn("TEXTBOX", "edi_id", "결과_edi_id", 10, false);
        }

        private void tsbQuery_Click(object sender, EventArgs e)
        {
            try
            {
                string strSql = "";
                strSql += @" SELECT L.dsgnNo, L.contrNm, L.trDt, L.CclsDt, L.AsgnEmpSeq, emp.EmpName AS AsgnEmpName, L.id AS req_id, L.id AS rtn_id, L.send_edi_id AS edi_id ";
                strSql += @" FROM _TAdjSL_SSL_SUIT AS L WITH(NOLOCK) ";
                strSql += @" LEFT JOIN _TDAEmp AS emp WITH(NOLOCK) ON emp.CompanySeq = L.CompanySeq ";
                strSql += @"      AND emp.EmpSeq     = L.AsgnEmpSeq ";
                strSql += @" WHERE  L.CompanySeq = @CompanySeq ";
                strSql += @" AND    L.transOutDt = '' ";
                strSql += @" AND    L.CclsFg = '6' ";
                strSql += @" AND    L.trDt >= @frdt ";
                strSql += @" AND    L.trDt <= @todt ";
                strSql += @" FOR JSON PATH ";

                List<IDbDataParameter> lstPara = new List<IDbDataParameter>();
                lstPara.Clear();
                lstPara.Add(new SqlParameter("@CompanySeq", YLWService.MTRServiceModule.SecurityJson.companySeq));
                lstPara.Add(new SqlParameter("@frdt", dtpFrDt.Value.ToString("yyyy-MM-dd")));
                lstPara.Add(new SqlParameter("@todt", dtpToDt.Value.ToString("yyyy-MM-dd")));
                strSql = Utils.GetSQL(strSql, lstPara.ToArray());
                strSql = strSql.Replace("\r\n", "");
                DataTable dt = MTRServiceModule.GetMTRServiceDataTable(YLWService.MTRServiceModule.SecurityJson.companySeq, strSql);
                if (dt == null)
                {
                    DataTable dtr = (DataTable)this.dgvList.DataSource;
                    if (dtr != null) dtr.Rows.Clear();
                    return;
                }
                this.dgvList.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tsbRowDelete_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewRow row in dgvList.SelectedRows)
                {
                    dgvList.Rows.Remove(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tsbDownload_Click(object sender, EventArgs e)
        {
            string trDt = DateTime.Now.ToString("yyyy-MM-dd");
            string fileName = "";
            int fileSeq = 0;

            try
            {
                ClearFolder();
                ClearAttachFolder();

                //Commit용 데이타 블럭
                DataTable dtC = new DataTable("DataBlock1");
                dtC.Columns.Add("send_type");
                dtC.Columns.Add("success_fg");
                dtC.Columns.Add("cust_code");
                dtC.Columns.Add("file_name");
                dtC.Columns.Add("rtn_id");
                dtC.Columns.Add("edi_id");
                dtC.Columns.Add("trDt");
                dtC.Clear();

                for (int ii = 0; ii < dgvList.Rows.Count; ii++)
                {
                    string rtn_id = Utils.ConvertToString(this.dgvList.Rows[ii].Cells["rtn_id"].Value);
                    string edi_id = Utils.ConvertToString(this.dgvList.Rows[ii].Cells["edi_id"].Value);
                    if (ii % 200 == 0)
                    {
                        fileSeq ++;
                        fileName = "R.JUHSRS." + trDt.Replace("-", "") + "." + Utils.PadLeft(fileSeq, 3, '0') + ".txt";
                        if (File.Exists(fileName))
                        {
                            MessageBox.Show("있는 파일명입니다[" + fileName + "]");
                            return;
                        }
                    }
                    DataRow dr = dtC.Rows.Add();
                    if (WriteFileAppend(fileName, rtn_id, edi_id, trDt, dr) < 1)
                    {
                        MessageBox.Show("전문파일 생성중 에러 발생");
                        return;
                    }
                }
                if (CommitAll(dtC))
                {
                    MessageBox.Show("전문파일 다운로드 완료\r\n삼성서버로 업로드 하세요");
                    DataTable dtr = (DataTable)dgvList.DataSource;
                    if (dtr != null) dtr.Rows.Clear();
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ClearScreen()
        {
            _bEvent = false;

            DataTable dtr = (DataTable)dgvList.DataSource;
            if (dtr != null) dtr.Rows.Clear();

            _bEvent = true;
        }

        private void ClearFolder()
        {
            DirectoryInfo dir = new DirectoryInfo(selectedPath);
            var files = dir.GetFiles("*.*", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                file.Attributes = FileAttributes.Normal;
                file.Delete();
            }
        }

        private void ClearAttachFolder()
        {
            DirectoryInfo dir = new DirectoryInfo(attachPath);
            var files = dir.GetFiles("*.*", SearchOption.TopDirectoryOnly);
            foreach (var file in files)
            {
                file.Attributes = FileAttributes.Normal;
                file.Delete();
            }
        }

        public int WriteFileAppend(string fileName, string rtn_id, string edi_id, string trDt, DataRow pdr)
        {
            YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
            security.serviceId = "Metro.Package.AdjSL.BisAdjSLEDITransSSL";
            security.methodId = "out";

            DataSet ds = new DataSet("ROOT");
            DataTable dt = ds.Tables.Add("DataBlock1");

            dt.Columns.Add("companyseq");
            dt.Columns.Add("send_type");
            dt.Columns.Add("success_fg");
            dt.Columns.Add("cust_code");
            dt.Columns.Add("rtn_id");
            dt.Columns.Add("edi_id");
            dt.Columns.Add("trDt");

            dt.Clear();
            DataRow dr = dt.Rows.Add();

            dr["companyseq"] = security.companySeq;
            dr["send_type"] = 1;
            dr["success_fg"] = 0;
            dr["cust_code"] = "SSL";
            dr["rtn_id"] = rtn_id;
            dr["edi_id"] = edi_id;
            dr["trDt"] = trDt;

            DataSet yds = MTRServiceModule.CallMTRServiceCall(security, ds);
            if (yds != null && yds.Tables.Count > 0)
            {
                if (yds.Tables.Contains("ErrorMessage")) throw new Exception(yds.Tables["ErrorMessage"].Rows[0]["Message"].ToString());
                DataTable dataBlock1 = yds.Tables["DataBlock1"];
                if (dataBlock1 != null && dataBlock1.Rows.Count > 0)
                {
                    if (!Directory.Exists(selectedPath)) Directory.CreateDirectory(selectedPath);
                    string file = selectedPath + "/" + fileName;
                    FileStream fs = new FileStream(file, FileMode.Append, FileAccess.Write);
                    using (StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("euc-kr")))
                    {
                        sw.Write(dataBlock1.Rows[0]["edi_text"]);
                        sw.Write("\r\n");
                        sw.Close();
                    }
                    fs.Close();

                    DataTable dataBlock2 = yds.Tables["DataBlock2"];
                    if (dataBlock2 != null && dataBlock2.Rows.Count > 0)
                    {
                        for (int ii = 0; ii < dataBlock2.Rows.Count; ii++)
                        {
                            if (!WriteAttachFile(dataBlock2.Rows[ii]["parent_id"] + "", dataBlock2.Rows[ii]["id"] + "")) return -1;
                        }
                    }
                    pdr["send_type"] = 1;
                    pdr["success_fg"] = 1;
                    pdr["cust_code"] = "SSL";
                    pdr["file_name"] = fileName;
                    pdr["rtn_id"] = rtn_id;
                    pdr["edi_id"] = dataBlock1.Rows[0]["edi_id"];
                    pdr["trDt"] = trDt;
                    return 1;
                }
            }
            return 0;
        }

        public bool CommitAll(DataTable dt)
        {
            YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
            security.serviceId = "Metro.Package.AdjSL.BisAdjSLEDITransSSL";
            security.methodId = "commit";

            DataSet ds = new DataSet("ROOT");
            ds.Tables.Add(dt);

            DataSet yds = MTRServiceModule.CallMTRServiceCallPost(security, ds);
            if (yds == null) return false;
            return true;
        }

        public bool WriteAttachFile(string parent_id, string id)
        {
            YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
            security.serviceId = "Metro.Package.AdjSL.BisAdjSLEDITransSSL";
            security.methodId = "attachOut";

            DataSet ds = new DataSet("ROOT");
            DataTable dt = ds.Tables.Add("DataBlock1");

            dt.Columns.Add("companyseq");
            dt.Columns.Add("parent_id");
            dt.Columns.Add("id");

            dt.Clear();
            DataRow dr = dt.Rows.Add();

            dr["companyseq"] = security.companySeq;
            dr["parent_id"] = parent_id;
            dr["id"] = id;

            DataSet yds = MTRServiceModule.CallMTRServiceCall(security, ds);
            if (yds != null && yds.Tables.Count > 0)
            {
                if (yds.Tables.Contains("ErrorMessage")) throw new Exception(yds.Tables["ErrorMessage"].Rows[0]["Message"].ToString());
                DataTable dataBlock1 = yds.Tables["DataBlock1"];
                if (dataBlock1 != null && dataBlock1.Rows.Count > 0)
                {
                    if (!Directory.Exists(attachPath)) Directory.CreateDirectory(attachPath);
                    string fileName = dataBlock1.Rows[0]["file_name"] + "";
                    string fileSeq = dataBlock1.Rows[0]["file_seq"] + "";
                    string file = attachPath + "/" + fileName;
                    if (File.Exists(file)) File.Delete(file);
                    string fileBase64 = YLWService.MTRServiceModule.CallMTRFileDownloadBase64(security, fileSeq, "0", "0");
                    if (fileBase64 != "")
                    {
                        byte[] bytes_file = Convert.FromBase64String(fileBase64);
                        FileStream fs = new FileStream(file, FileMode.Create);
                        fs.Write(bytes_file, 0, bytes_file.Length);
                        fs.Close();
                    }
                }
            }
            return true;
        }
    }
}