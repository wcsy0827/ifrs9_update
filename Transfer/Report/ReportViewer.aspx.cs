using System;
using System.Linq;
using Transfer.Utility;
using Microsoft.Reporting.WebForms;

namespace Transfer.Report
{
    public partial class ReportViewer : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                GenerateReport();
            }
        }

        private void GenerateReport()
        {
            try
            {
                // 20.03.04 John.下載多張報表修正
                //var ReportWrapperSessionKey = "ReportWrapper";                
                var ReportWrapperSessionKey = Request.QueryString["ReportId"];
                var rw = (ReportWrapper)base.Session[ReportWrapperSessionKey];
                if (rw != null)
                {
                    var RptViewer = this.rptViewer;

                    RptViewer.LocalReport.DataSources.Clear();

                    // Rdlc location
                    RptViewer.LocalReport.ReportPath = rw.ReportPath;

                    // Set report data source
                    RptViewer.LocalReport.DataSources.Clear();
                    foreach (var reportDataSource in rw.ReportDataSources)
                    { RptViewer.LocalReport.DataSources.Add(reportDataSource); }


                    // Set DownLoad Name
                    string Name = "";
                    var _name = rw.ReportParameters.FirstOrDefault(x => x.Name == "Title");
                    if (_name != null)
                    {
                        string _DisplayName = _name.Values[0].ToString();
                        _DisplayName = _DisplayName.Replace("(", "-").Replace(")", "");
                        Name = _DisplayName;
                        RptViewer.LocalReport.DisplayName = _DisplayName;
                    }

                    //Set DownLoad Format
                    string Format = "";
                    ReportParameter _Format = rw.ReportParameters.FirstOrDefault(x => x.Name == "Format");
                    if (_Format != null)
                    {
                        Format = _Format.Values[0].ToString();
                    }

                    // Set report parameters
                    RptViewer.LocalReport.SetParameters(rw.ReportParameters);


                    // Refresh report
                    RptViewer.LocalReport.Refresh();

                    //Format Conversion
                    if (Format == "Excel" || Format == "PDF")
                    {
                        Warning[] warnings;
                        string[] streamIds;
                        string mimeType = string.Empty;
                        string encoding = string.Empty;
                        string extension = string.Empty;

                        byte[] bytes = RptViewer.LocalReport.Render(Format, null, out mimeType, out encoding, out extension, out streamIds, out warnings);

                        Response.Buffer = true;
                        Response.Clear();
                        Response.ContentType = mimeType;
                        Response.AddHeader("content-disposition", "attachment; filename=" + Name + "." + extension);
                        Response.BinaryWrite(bytes);
                        Response.Flush();
                        Response.End();
                    }

                    // Remove session
                    Session[ReportWrapperSessionKey] = null;
                }
                else
                {
                    Response.Redirect("/Account/Login");
                    //Response.End();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}