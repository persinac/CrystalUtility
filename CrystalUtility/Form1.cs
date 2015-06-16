using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CrystalUtility
{
    public partial class Form1 : Form
    {
        Utility u;
        DirectoryUtility du;
        List<String> procedures;
        List<String> views;
        List<String> sr_views;

        public Form1()
        {
            InitializeComponent();
            u = new Utility();
            du = new DirectoryUtility();
            procedures = new List<String>();
            views = new List<String>();
            sr_views = new List<String>();

            string t_path = @"C:\SVN\CrystalReports";

            foreach (string dir in Directory.GetDirectories(t_path))
            {
                cboDirectoryList.Items.Add(dir);
            }
        }

        private void btn_Run_Click(object sender, EventArgs e)
        {
            if (txtUsername.Text.Length > 1 && txtPassword.Text.Length > 1)
            {
                String username = txtUsername.Text;
                String password = txtPassword.Text;
                rTextBoxOutput.AppendText("Starting...\n\n");
                rTextBoxOutput.ScrollToCaret();
                String path = this.cboDirectoryList.Text;//"C:\\SVN\\CrystalReports\\OBS";
                CreateDirectoryStructure();
                foreach (string file in Directory.GetFiles(path, "*.rpt"))
                {

                    Console.WriteLine(String.Format("Processing {0}...", file));
                    rTextBoxOutput.AppendText(String.Format("Processing {0}...\n", file));
                    rTextBoxOutput.ScrollToCaret();

                    // Declarations
                    CrystalDecisions.CrystalReports.Engine.ReportDocument boReportDocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                    CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument boReportClientDocument;
                    CrystalDecisions.ReportAppServer.Controllers.RowsetController boRowsetController;
                    CrystalDecisions.ReportAppServer.DataDefModel.ISCRGroupPath boGroupPath;
                    string temp = "";
                    try
                    {
                        // Load the report from the application directory
                        boReportDocument.Load(file);

                        // Access the ReportClientDocument in the ReportDocument (EROM bridge)
                        // Note this is available without a dedicated RAS with SP2 for XI R2
                        boReportClientDocument = boReportDocument.ReportClientDocument;
                        boReportDocument.DataSourceConnections[0].SetLogon(username, password);

                        if (boReportDocument.Subreports.Count == 0)
                        {

                            CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions parameterList = boReportDocument.DataDefinition.ParameterFields;
                            for (int i = 0; i < parameterList.Count; i++)
                            {
                                Console.WriteLine("i = " + i + "... " + parameterList[i].ParameterFieldName.ToString());
                                if (parameterList[i].ParameterFieldName.ToString() == "ReportType")
                                {
                                    boReportDocument.SetParameterValue(i, "D");
                                }
                                else if (parameterList[i].ParameterFieldName.ToString() == "@ControlID")
                                {
                                    boReportDocument.SetParameterValue(i, 76461450);
                                }
                                else if (parameterList[i].ParameterFieldName.ToString() == "@RevisionCode")
                                {
                                    boReportDocument.SetParameterValue(i, 0);
                                }
                            }

                            // use the RowsetController to get the SQL query
                            // Note: If a report has parameters they must be supplied before getting the 
                            // SQL query.
                            boRowsetController = boReportClientDocument.RowsetController;
                            boGroupPath = new CrystalDecisions.ReportAppServer.DataDefModel.GroupPath();

                            String s = boRowsetController.GetSQLStatement(boGroupPath, out temp);

                            s = u.RemoveQuotes(s);
                            bool ret = false;
                            if(u.HasSQL(s)) {
                                if (u.HasProduct(s))
                                {
                                    if (du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\HasProductReference"))
                                    {
                                        ret = du.MoveFileToDirectory(file, path + "\\HasEmbeddedSQL\\HasProductReference\\" + Path.GetFileName(file));
                                        temp = "Processed " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\HasEmbeddedSQL\\HasProductReference\\";
                                    }
                                } else if(u.HasView(s)) {
                                    if (du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\HasView"))
                                    {
                                        ret = du.MoveFileToDirectory(file, path + "\\HasEmbeddedSQL\\HasView\\" + Path.GetFileName(file));
                                        temp = "Processed " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\HasEmbeddedSQL\\HasView\\";
                                        views.Add(boReportDocument.Database.Tables[0].Name);
                                    }
                                }
                                else
                                {
                                    if (du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\ProbablyOK"))
                                    {
                                        ret = du.MoveFileToDirectory(file, path + "\\HasEmbeddedSQL\\ProbablyOK\\" + Path.GetFileName(file));
                                        temp = "Processed " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\HasEmbeddedSQL\\ProbablyOK\\";
                                    }
                                }
                            } 
                            else if (u.UsesExec(s))
                            {
                                if (du.DoesDirectoryExist(path + "\\NoEmbeddedSQL"))
                                {
                                    ret = du.MoveFileToDirectory(file, path + "\\NoEmbeddedSQL\\" + Path.GetFileName(file));
                                    temp = "Processed " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\NoEmbeddedSQL\\";
                                }
                            }
                            else
                            {
                                if (du.DoesDirectoryExist(path + "\\ShouldReview"))
                                {
                                    ret = du.MoveFileToDirectory(file, path + "\\ShouldReview\\" + Path.GetFileName(file));
                                    temp = "Processed " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\ShouldReview\\";
                                }
                            }

                            rTextBoxOutput.AppendText(temp + "\n");
                            rTextBoxOutput.ScrollToCaret();
                        }
                        else
                        {
                            ConnectionInfo ci = new ConnectionInfo();
                            ci.UserID = username;
                            ci.Password = password;
                            ApplyLogOnInfoForSubreports(boReportDocument, ci);

                            if (du.DoesDirectoryExist(path + "\\HasSubReports"))
                            {
                                bool ret = du.MoveFileToDirectory(file, path + "\\HasSubReports\\" + Path.GetFileName(file));
                                temp = "HAS SUBREPORT(S)" + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\HasSubReports\\";
                            }
                            rTextBoxOutput.AppendText(temp + "\n");
                            rTextBoxOutput.ScrollToCaret();
                        }
                    }
                    catch (System.ArgumentOutOfRangeException aor)
                    {
                        if (du.DoesDirectoryExist(path + "\\ShouldReview"))
                        {
                            bool ret = du.MoveFileToDirectory(file, path + "\\ShouldReview\\" + Path.GetFileName(file));
                            temp = "Should REVIEW " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\ShouldReview\\";
                        }
                        rTextBoxOutput.AppendText(temp + "\n");
                        rTextBoxOutput.AppendText(aor.Message + "\n");
                        rTextBoxOutput.ScrollToCaret();
                    }
                    catch (System.Runtime.InteropServices.COMException comex)
                    {
                        if (du.DoesDirectoryExist(path + "\\ShouldReview"))
                        {
                            bool ret = du.MoveFileToDirectory(file, path + "\\ShouldReview\\" + Path.GetFileName(file));
                            temp = "Should REVIEW " + Path.GetFileName(file) + ", moved file: " + ret + ", moved to " + path + "\\ShouldReview\\";
                        }
                        rTextBoxOutput.AppendText(temp + "\n");
                        rTextBoxOutput.AppendText(comex.Message + "\n");
                        rTextBoxOutput.ScrollToCaret();
                    }
                    finally
                    {
                        // Clean up by closing and disposing of the ReportDocument object
                        boReportDocument.Close();
                        boReportDocument.Dispose();
                    }

                }
                List<String> noDupes_procs = procedures.Distinct().ToList();
                List<String> noDupes_sr_views = sr_views.Distinct().ToList();
                List<String> noDupes_views = views.Distinct().ToList();
                u.WriteToFile(path + "\\HasSubReports", "_subreportSPs.txt", noDupes_procs);
                u.WriteToFile(path + "\\HasSubReports", "_subreportViews.txt", noDupes_sr_views);
                u.WriteToFile(path + "\\HasEmbeddedSQL\\HasView", "_reportViews.txt", noDupes_views);
            }
            else
            {
                MessageBox.Show("Please enter a username and password!");
            }
        }

        private void ApplyLogOnInfoForSubreports(ReportDocument reportDocument, ConnectionInfo connectionInfo)
        {
            Sections sections = reportDocument.ReportDefinition.Sections;
            foreach (Section section in sections)
            {
                ReportObjects reportObjects = section.ReportObjects;
                foreach (ReportObject reportObject in reportObjects)
                {
                    //_log.InfoFormat("Type = {0}, Name = {1}", reportObject.Name, reportObject.Kind);
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        var subreportObject = (SubreportObject)reportObject;
                        ReportDocument subReportDocument = subreportObject.OpenSubreport(subreportObject.SubreportName);
                        SearchSubreport(subReportDocument);
                    }
                }
            }
        }

        private void SearchSubreport(ReportDocument reportDocument)
        {
            if(u.HasView(reportDocument.Database.Tables[0].Location)) {
                sr_views.Add(reportDocument.Database.Tables[0].Location);
            }
            else if (u.HasProc(reportDocument.Database.Tables[0].Location))
            {
                procedures.Add(reportDocument.Database.Tables[0].Location);
            }
        }

        private void ApplyLogOnInfo(ReportDocument reportDocument, ConnectionInfo connectionInfo)
        {
            foreach (Table table in reportDocument.Database.Tables)
            {
                table.LogOnInfo.ConnectionInfo.AllowCustomConnection = true;
                TableLogOnInfo tableLogonInfo = table.LogOnInfo;
                tableLogonInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(tableLogonInfo);
            }
        }

        private void CreateDirectoryStructure()
        {
            String path = this.cboDirectoryList.Text;//"C:\\SVN\\CrystalReports\\OBS";

            if (!du.DoesDirectoryExist(path + "\\HasEmbeddedSQL"))
            {
                if (du.CreateDirectory(path + "\\HasEmbeddedSQL"))
                {
                    Console.WriteLine("SUCCESS - HasEmbeddedSQL");
                }
                else
                {
                    Console.WriteLine("FAIL - HasEmbeddedSQL");
                }
            }
            if (!du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\HasProductReference"))
            {
                if (du.CreateDirectory(path + "\\HasEmbeddedSQL\\HasProductReference"))
                {
                    Console.WriteLine("SUCCESS - HasEmbeddedSQL\\HasProductReference");
                }
                else
                {
                    Console.WriteLine("FAIL - HasEmbeddedSQL\\HasProductReference");
                }
                    du.CreateDirectory(path + "\\HasEmbeddedSQL\\ProbablyOK");
                    du.CreateDirectory(path + "\\HasEmbeddedSQL\\HasView");
            }

            if (!du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\HasView"))
            {
                if (du.CreateDirectory(path + "\\HasEmbeddedSQL\\HasView"))
                {
                    Console.WriteLine("SUCCESS - HasEmbeddedSQL\\HasView");
                }
                else
                {
                    Console.WriteLine("FAIL - HasEmbeddedSQL\\HasView");
                }
                du.CreateDirectory(path + "\\HasEmbeddedSQL\\ProbablyOK");
            }

            if (!du.DoesDirectoryExist(path + "\\HasEmbeddedSQL\\ProbablyOK"))
            {
                if (du.CreateDirectory(path + "\\HasEmbeddedSQL\\ProbablyOK"))
                {
                    Console.WriteLine("SUCCESS - HasEmbeddedSQL\\ProbablyOK");
                }
                else
                {
                    Console.WriteLine("FAIL - HasEmbeddedSQL\\ProbablyOK");
                }
            }

            if (!du.DoesDirectoryExist(path + "\\NoEmbeddedSQL"))
            {
                if (du.CreateDirectory(path + "\\NoEmbeddedSQL"))
                {
                    Console.WriteLine("SUCCESS - NoEmbeddedSQL");
                }
                else
                {
                    Console.WriteLine("FAIL - NoEmbeddedSQL");
                }
            }

            if (!du.DoesDirectoryExist(path + "\\ShouldReview"))
            {
                if (du.CreateDirectory(path + "\\ShouldReview"))
                {
                    Console.WriteLine("SUCCESS - ShouldReview");
                }
                else
                {
                    Console.WriteLine("FAIL - ShouldReview");
                }
            }

            if (!du.DoesDirectoryExist(path + "\\HasSubReports"))
            {
                if (du.CreateDirectory(path + "\\HasSubReports"))
                {
                    Console.WriteLine("SUCCESS - HasSubReports");
                }
                else
                {
                    Console.WriteLine("FAIL - HasSubReports");
                }
            }
        }
    }
}
