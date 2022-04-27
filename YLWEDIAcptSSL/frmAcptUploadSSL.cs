using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using YLWService;
using YLWService.Extensions;

namespace YLWEDISendSSL
{
    public partial class frmAcptUploadSSL : Form
    {
        bool _bEvent = false;

        string dept7 = "";
        string selectedPath = "";
        string getAttachPath = "";

        public frmAcptUploadSSL()
        {
            InitializeComponent();

            this.Load += FrmAcptUploadSSL_Load;
            this.dgvList.RowPostPaint += DgvList_RowPostPaint;
            this.dgvList.KeyDown += DgvList_KeyDown;

            SetInit();

            _bEvent = true;
        }

        private void DgvList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                tsbRowDelete_Click(dgvList, new EventArgs());
            }
        }

        private void DgvList_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dgvList.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }

        private void FrmAcptUploadSSL_Load(object sender, EventArgs e)
        {
        }

        private void SetInit()
        {
            try
            {
                selectedPath = YLWServiceModule.GetInPath();
                getAttachPath = YLWServiceModule.GetGetfilePath();

                SetComboDept(cboDept);
                //Utils.SetComboSelectedValue(cboDept, "7", "BeDeptSeq");  //하드코딩을 배제하기 위해서 막음 2021-06-08

                DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
                col1.HeaderText = "결과";
                col1.DataPropertyName = "success_fg";
                col1.Name = "success_fg";
                col1.Width = 50;
                col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                DataGridViewTextBoxColumn col2 = new DataGridViewTextBoxColumn();
                col2.HeaderText = "전문내용";
                col2.DataPropertyName = "edi_text";
                col2.Name = "edi_text";
                col2.DefaultCellStyle.Font = new Font("굴림체", 9);
                dgvList.Columns.AddRange(col1, col2);
                Type dgvType = dgvList.GetType();
                PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
                pi.SetValue(dgvList, true, null);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tsbReadDirectory_Click(object sender, EventArgs e)
        {
            try
            {
                txtFileName.Text = "";
                dgvList.DataSource = null;

                OpenFileDialog dlg = new OpenFileDialog();
                dlg.InitialDirectory = selectedPath;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    txtFileName.Text = dlg.FileName;
                    ReadFile(dlg.FileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ReadFile(string fileName)
        {
            try
            {
                int counter = 0;
                string line = "";
                DataTable dt = new DataTable();
                dt.Columns.Add("edi_text");
                using (System.IO.StreamReader file = new System.IO.StreamReader(fileName, Encoding.GetEncoding("euc-kr")))
                {
                    while ((line = file.ReadLine()) != null)
                    {
                        DataRow dr = dt.Rows.Add();
                        dr["edi_text"] = line;
                        counter++;
                    }
                    file.Close();
                }
                dgvList.DataSource = dt;
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

        private void tsbUpload_Click(object sender, EventArgs e)
        {
            if (txtFileName.Text == "")
            {
                MessageBox.Show("업로드 파일을 선택하세요");
                return;
            }
            string dept = Utils.GetComboSelectedValue(cboDept, "BeDeptSeq");
            if (dept == "")
            {
                MessageBox.Show("조사팀은 필수입니다");
                return;
            }

            try
            {
                string filename = Path.GetFileName(txtFileName.Text);
                for (int ii = 0; ii < dgvList.Rows.Count; ii++)
                {
                    dgvList.CurrentCell = dgvList.Rows[ii].Cells[0];
                    if (!ReadTxtFile(filename, dgvList.Rows[ii]))  //전문데이타 업로드
                    {
                        return;
                    }
                    Application.DoEvents();
                }
                for (int ii = dgvList.Rows.Count - 1; ii >= 0; ii--)
                {
                    dgvList.CurrentCell = dgvList.Rows[ii].Cells[0];
                    if (dgvList.Rows[ii].Cells["success_fg"].Value + "" == "OK")
                    {
                        dgvList.Rows.Remove(dgvList.Rows[ii]);
                    }
                    else if (dgvList.Rows[ii].Cells["success_fg"].Value + "" == "EXIST")
                    {
                        dgvList.Rows.Remove(dgvList.Rows[ii]);
                    }
                    Application.DoEvents();
                }
                if (dgvList.Rows.Count < 1)
                {
                    File.Delete(txtFileName.Text);
                    txtFileName.Text = "";
                    dgvList.DataSource = null;
                    MessageBox.Show("접수 업로드 완료");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public bool ReadTxtFile(string filename, DataGridViewRow grow)
        {
            YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
            security.serviceId = "Metro.Package.AdjSL.BisAdjSLEDITransSSL";
            security.methodId = "in";

            DataSet ds = new DataSet("ROOT");
            DataTable dt = ds.Tables.Add("DataBlock1");

            dt.Columns.Add("companyseq");
            dt.Columns.Add("send_type");
            dt.Columns.Add("success_fg");
            dt.Columns.Add("cust_code");
            dt.Columns.Add("trans_dtm");
            dt.Columns.Add("file_name");
            dt.Columns.Add("edi_text");
            dt.Columns.Add("CclsFg");
            dt.Columns.Add("AsgnTeamSeq");
            dt.Columns.Add("AsgnEmpSeq");

            try
            {
                string editext = grow.Cells["edi_text"].Value + "";
                string dept = Utils.GetComboSelectedValue(cboDept, "BeDeptSeq");
                string emp = txtEmpSeq.Text;
                string cclsfg = "0";
                if (dept != "") cclsfg = "1";
                //if (emp != "") cclsfg = "2";

                dt.Clear();
                DataRow dr = dt.Rows.Add();

                dr["companyseq"] = security.companySeq;
                dr["send_type"] = 0;
                dr["success_fg"] = 0;
                dr["cust_code"] = "SSL";
                dr["file_name"] = filename;
                dr["edi_text"] = editext;
                dr["CclsFg"] = cclsfg;
                dr["AsgnTeamSeq"] = dept;
                dr["AsgnEmpSeq"] = emp;

                DataSet yds = MTRServiceModule.CallMTRServiceCallPost(security, ds);
                if (yds != null)
                {
                    DataTable dataBlock1 = yds.Tables["DataBlock1"];
                    if (dataBlock1 != null && dataBlock1.Rows.Count > 0)
                    {
                        if (dataBlock1.Rows[0]["success_fg"] + "" == "1")
                        {
                            string edi_id = dataBlock1.Rows[0]["edi_id"] + "";

                            //첨부파일에 대한 설명이 없음
                            //var files = Directory.GetFiles(getAttachPath, "*.*", SearchOption.TopDirectoryOnly);
                            //foreach (var file in files)  //첨부파일
                            //{
                            //    if (Path.GetExtension(file).ToUpper() == ".TXT") continue;
                            //    if (!ReadAttachFile(edi_id, file)) return false;
                            //}
                            grow.Cells["success_fg"].Value = "OK";
                        }
                        else if (dataBlock1.Rows[0]["success_fg"] + "" == "2")
                        {
                            //있는 설계번호임
                            grow.Cells["success_fg"].Value = "EXIST";
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                YLWService.LogWriter.WriteLog(ex.Message);
                return false;
            }
        }

        public bool ReadAttachFile(string edi_id, string file)
        {
            YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
            security.serviceId = "Metro.Package.AdjSL.BisAdjSLEDITransSSL";
            security.methodId = "attach";

            DataSet ds = new DataSet("ROOT");
            DataTable dt = ds.Tables.Add("DataBlock1");

            dt.Columns.Add("companyseq");
            dt.Columns.Add("success_fg");
            dt.Columns.Add("trans_dtm");
            dt.Columns.Add("file_name");
            dt.Columns.Add("file_seq");
            dt.Columns.Add("edi_id");
            dt.Columns.Add("parent_id");
            dt.Columns.Add("id");

            try
            {
                string filename = Path.GetFileName(file);
                // File Info
                FileInfo finfo = new FileInfo(file);
                byte[] rptbyte = (byte[])MetroSoft.HIS.cFile.ReadBinaryFile(file);
                string fileBase64 = Convert.ToBase64String(rptbyte);
                // File Info
                //string fileSeq = YLWService.MTRServiceModule.CallMTRFileuploadGetSeq(security, finfo, fileBase64, "47820004");  // 이부분에서 오류남. CallMTRFileuploadGetSeq -> FileuploadGetSeq
                //string fileSeq = YLWService.YLWServiceModule.FileuploadGetSeq(security, finfo, fileBase64, "47820004");
                string fileSeq = YLWService.MTRServiceModule.CallMTRFileuploadGetSeq(security, finfo, fileBase64, "47820004");
                if (fileSeq == "") return false;

                dt.Clear();
                DataRow dr = dt.Rows.Add();
                dr["companyseq"] = security.companySeq;
                dr["success_fg"] = "0";
                dr["trans_dtm"] = "";
                dr["file_name"] = filename;
                dr["file_seq"] = fileSeq;
                dr["edi_id"] = edi_id;
                dr["parent_id"] = "0";
                dr["id"] = "0";

                DataSet yds = MTRServiceModule.CallMTRServiceCallPost(security, ds);
                if (yds != null)
                {
                    DataTable dataBlock1 = yds.Tables["DataBlock1"];
                    if (dataBlock1 != null && dataBlock1.Rows.Count > 0)
                    {
                        if (dataBlock1.Rows[0]["success_fg"] + "" != "1") return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                YLWService.LogWriter.WriteLog(ex.Message);
                return false;
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ClearScreen()
        {
            _bEvent = false;

            cboDept.SelectedIndex = -1;
            txtEmp.Text = "";
            txtEmpSeq.Text = "";
            DataTable dtr = (DataTable)dgvList.DataSource;
            if (dtr != null) dtr.Rows.Clear();

            _bEvent = true;
        }

        private void SetComboDept(ComboBox cboObj)
        {
            List<IDbDataParameter> lstPara = new List<IDbDataParameter>();
            string strSql = "";

            strSql = "SELECT b.MinorName FROM _TDAUMinor b WHERE b.CompanySeq = @CompanySeq AND b.MajorSeq = '300163' AND b.MinorSeq = '300163002' FOR JSON PATH ";  //300163002:7(물)
            lstPara.Clear();
            lstPara.Add(new SqlParameter("@CompanySeq", YLWService.MTRServiceModule.SecurityJson.companySeq));
            strSql = Utils.GetSQL(strSql, lstPara.ToArray());
            DataTable dt7 = MTRServiceModule.GetMTRServiceDataTable(YLWService.MTRServiceModule.SecurityJson.companySeq, strSql);
            dept7 = "";
            if (dt7 != null && dt7.Rows.Count > 0) dept7 = dt7.Rows[0]["MinorName"] + "";

            strSql = "";
            strSql += @" SELECT A.DeptName AS BeDeptName ";
            strSql += @"     ,A.BegDate AS BeBegDate ";
            strSql += @"     ,A.EndDate AS BeEndDate ";
            strSql += @"     ,A.Remark AS DeptRemark ";
            strSql += @"     ,A.DeptSeq AS BeDeptSeq ";
            strSql += @"     ,CASE WHEN A.EndDate >= CONVERT(NCHAR(8), GETDATE(), 112) THEN '1' ELSE '0' END AS IsUse  ";  /* 현재일을 기준으로 사용중인 부서인지 아닌지 판단 */
            strSql += @" FROM _TDADept AS A WITH(NOLOCK) ";
            strSql += @"      JOIN [dbo].[_fnOrgDeptHR](@CompanySeq, 1, @HeadDeptSeq, CONVERT(NCHAR(8), GETDATE(), 112)) AS hddept ON hddept.DeptSeq = A.DeptSeq ";
            strSql += @"      LEFT JOIN _THROrgDeptCCtr AS B WITH(NOLOCK) ON A.CompanySeq = B.CompanySeq AND A.DeptSeq = B.DeptSeq AND B.IsLast = '1' ";   /* 무조건 최종 활동센터를 표시한다. 2011.12.08 민형준 */
            strSql += @"      LEFT JOIN _TDACCtr AS C WITH(NOLOCK) ON A.CompanySeq = C.CompanySeq AND B.CCtrSeq = C.CCtrSeq ";
            strSql += @" WHERE A.CompanySeq = @CompanySeq ";
            strSql += @" AND   A.SMDeptType NOT IN(3051003, 3051004) ";  /* TFT제외, BPM부서 제외 */
            strSql += @" ";  /* AND CASE @DefQueryOption WHEN 0 THEN CONVERT(NVARCHAR(10), A.DeptSeq) ELSE A.DeptName END LIKE @Keyword */
            strSql += @" ORDER BY A.DeptName, A.DispSeq ";
            strSql += @" FOR JSON PATH ";
            lstPara.Clear();
            lstPara.Add(new SqlParameter("@CompanySeq", YLWService.MTRServiceModule.SecurityJson.companySeq));
            lstPara.Add(new SqlParameter("@HeadDeptSeq", dept7));
            strSql = Utils.GetSQL(strSql, lstPara.ToArray());
            DataTable dt = MTRServiceModule.GetMTRServiceDataTable(YLWService.MTRServiceModule.SecurityJson.companySeq, strSql);
            Utils.SetCombo(cboObj, dt, "BeDeptSeq", "BeDeptName", true);
        }

        private void txtEmp_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string dept = Utils.GetComboSelectedValue(cboDept, "BeDeptSeq");
            FrmFindEmp f = new FrmFindEmp(dept, txtEmp.Text);
            if (f.ShowDialog(this) == DialogResult.OK)
            {
                _bEvent = false;
                txtEmp.Text = f.ReturnFields.Find(x => x.FieldCode == "EmpName").FieldValue.ToString();
                txtEmpSeq.Text = f.ReturnFields.Find(x => x.FieldCode == "EmpSeq").FieldValue.ToString();
                dept = f.ReturnFields.Find(x => x.FieldCode == "DeptSeq").FieldValue.ToString();
                Utils.SetComboSelectedValue(cboDept, dept, "BeDeptSeq");
                _bEvent = true;
            }
        }

        private void cboDept_SelectedValueChanged(object sender, EventArgs e)
        {
            if (!_bEvent) return;
            txtEmp.Text = "";
            txtEmpSeq.Text = "";
        }
    }
}