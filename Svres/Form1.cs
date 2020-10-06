using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace Svres
{
    public partial class Form1 : Form
    {
        public string PathFile { get; set; }
        public List<WarnInfo> lWarnings{ get; set; }
        public List<WarnInfoEx> lWarnInfoEx { get; set; }
        public string ProjectName { get; set; }
        public string WorkFolder { get; set; }
        public Form1()
        {
            InitializeComponent();
            lWarnings = new List<WarnInfo>();
            lWarnInfoEx = new List<WarnInfoEx>();
        }

        private void bOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "svres files (*.svres)|*.svres|All files (*.*)|*.*";

            

            Regex rWarnInfo = new Regex(@"WarnInfo\s");
            Regex rWarnInfoEx = new Regex(@"WarnInfoEx\s");
            Regex rProjectName = new Regex(@"<projectName");
            Regex rLocInfo = new Regex(@"<LocInfo\s");
            Regex rSeverity = new Regex(@">severity<");

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                PathFile = ofd.FileName;
                WorkFolder = Path.GetDirectoryName(PathFile);
                string[] lines = File.ReadAllLines(PathFile);
                int n = 0;
                WarnInfoEx warnInfoEx = new WarnInfoEx();
                foreach (string s in lines)
                {
                    MatchCollection matches = rProjectName.Matches(s);
                    if (matches.Count > 0)
                    {
                        string[] projectName = s.Split(new char[] { '>', '<' });
                        ProjectName = projectName[2];
                    }
                        
                    matches = rWarnInfo.Matches(s);
                    if (matches.Count > 0)
                    {
                        WarnInfo warnInfo = new WarnInfo();
                        string[] date = s.Split(new char[] { '\"'});
                        warnInfo.Id = int.Parse(date[1]);
                        warnInfo.WarnClass = date[3];
                        warnInfo.NumberLine = int.Parse(date[5]);
                        warnInfo.Pathfile = date[7];
                        warnInfo.Message = date[9];
                        warnInfo.Status = date[11];
                        warnInfo.Details = date[13];
                        warnInfo.Comment = date[15];
                        warnInfo.Function = date[17];
                        warnInfo.Mtid = date[19];
                        warnInfo.Tool = date[21];
                        warnInfo.Lang = date[23];
                        lWarnings.Add(warnInfo);
                    }

                     
                    matches = rWarnInfoEx.Matches(s);
                    if (matches.Count > 0)
                    {
                        warnInfoEx = new WarnInfoEx();
                        string[] date = s.Split(new char[] { '\"' });
                        warnInfoEx.Id = int.Parse(date[1]);
                    }
                    matches = rLocInfo.Matches(s);
                    if (matches.Count > 0)
                    {
                        string[] date = s.Split(new char[] { '\"' });
                        LocInfo locInfo = new LocInfo();
                        locInfo.PathFile = date[1];
                        locInfo.Line = int.Parse(date[3]);
                        locInfo.Spec = date[5];
                        locInfo.Info = date[7];
                        locInfo.Col = int.Parse(date[9]);
                        warnInfoEx.LocInfo.Add(locInfo);

                    }
                    
                    matches = rSeverity.Matches(s);
                    if (matches.Count > 0)
                    {
                        //string[] date = s.Split(new char[] { '>', '<' });
                        string[] severity = lines[n+1].Split(new char[] { '>', '<' });
                        warnInfoEx.Severity = severity[2];
                        lWarnInfoEx.Add(warnInfoEx);
                    }
                    ++n;
                }
                MessageBox.Show(String.Format("Find {0} warnings.", lWarnings.Count));

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Report Svace 1");
                    worksheet.Cells[1, 1].Value = "Id";
                    worksheet.Cells[1, 2].Value = "WarnClass";
                    worksheet.Cells[1, 3].Value = "NumberLine";
                    worksheet.Cells[1, 4].Value = "Pathfile";
                    worksheet.Cells[1, 5].Value = "Message";
                    worksheet.Cells[1, 6].Value = "Status";
                    worksheet.Cells[1, 7].Value = "Details";
                    worksheet.Cells[1, 8].Value = "Comment";
                    worksheet.Cells[1, 9].Value = "Function";
                    worksheet.Cells[1, 10].Value = "Mtid";
                    worksheet.Cells[1, 11].Value = "Tool";
                    worksheet.Cells[1, 12].Value = "Lang";
                    worksheet.Cells[1, 13].Value = "Severity";

                    int i;
                    //
                    if (lWarnings.Count > 0)
                    {
                        for (i = 0; i < lWarnings.Count; i++)
                        {
                            worksheet.Cells["A" + (i + 2)].Value = lWarnings[i].Id;
                            worksheet.Cells["B" + (i + 2)].Value = lWarnings[i].WarnClass;
                            worksheet.Cells["C" + (i + 2)].Value = lWarnings[i].NumberLine;
                            worksheet.Cells["D" + (i + 2)].Value = lWarnings[i].Pathfile;
                            worksheet.Cells["E" + (i + 2)].Value = lWarnings[i].Message;
                            worksheet.Cells["F" + (i + 2)].Value = lWarnings[i].Status;
                            worksheet.Cells["G" + (i + 2)].Value = lWarnings[i].Details;
                            worksheet.Cells["H" + (i + 2)].Value = lWarnings[i].Comment;
                            worksheet.Cells["I" + (i + 2)].Value = lWarnings[i].Function;
                            worksheet.Cells["J" + (i + 2)].Value = lWarnings[i].Mtid;
                            worksheet.Cells["K" + (i + 2)].Value = lWarnings[i].Tool;
                            worksheet.Cells["L" + (i + 2)].Value = lWarnings[i].Lang;

                            for (int j = 0; j < lWarnInfoEx.Count; j++)
                            {
                                if(lWarnInfoEx[j].Id.Equals(lWarnings[i].Id))
                                {
                                    worksheet.Cells["M" + (i + 2)].Value = lWarnInfoEx[j].Severity;
                                }
                            }
                        }
                        //worksheet.Cells.AutoFitColumns(1);
                    }

                    try
                    {
                        package.SaveAs(new FileInfo(WorkFolder + "\\" + String.Format(ProjectName + "_svace_report.xlsx")));
                        MessageBox.Show("File excel create.");
                    }
                    catch(IOException ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    
                }
            }
            else
            {
                return;
            }
        }
    }
}
