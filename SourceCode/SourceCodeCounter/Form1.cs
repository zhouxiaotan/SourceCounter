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

namespace SourceCodeCounter
{
    public partial class Form1 : Form, IWin32Window
    {
        private List<PackageInfo> packageList = new List<PackageInfo>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.textPath.Text == string.Empty)
            {
                return;
            }

            GetPackageList();

            this.listBox1.Items.Clear();
            foreach (var item in packageList)
            {
                this.listBox1.Items.Add(item.PackageId);
            }

            GetPackageDetail();


        }

        public bool IsCommentLine(string line, string fileName)
        {
            bool ret = false;
            if (fileName.Contains(".cs"))
            {
                ret = line.Contains("//") && line.Trim().Substring(0, 2) == "//";
            }
            else if (fileName.Contains(".xaml"))
            {
                ret = line.Contains("<!--") && line.Trim().Substring(0, 4) == "<!--";
            }
            else if (fileName.Contains(".cpp") || fileName.Contains(".h"))
            {
                ret = line.Contains("//") && line.Trim().Substring(0, 2) == "//";
                if (!ret)
                {
                    ret = line.Contains("/*") && line.Trim().Substring(0, 2) == "/*";
                }
                if (!ret)
                {
                    ret = line.Contains("*/") && line.Trim().Substring(line.Length - 3, 2) == "*/";
                }
            }
            return ret;
        }

        public void GetPackageDetail()
        {
            progressBar1.Maximum = packageList.Count;
            progressBar1.Value = 0;

            foreach (var pack in packageList)
            {
                progressBar1.Value = packageList.IndexOf(pack) + 1;
                this.listBox1.SelectedItem = pack.PackageId;

                var appendStr = pack.PackageId + "A";
                var replaceStr = pack.PackageId + "R";
                var deleteStr = pack.PackageId + "D";
                var endStr = pack.PackageId + "E";

                foreach (var packfile in pack.PackFileList)
                {
                    var fileAllLines = File.ReadAllLines(packfile.FilePath);

                    // When Created File
                    if ((fileAllLines.Count(obj => obj.Contains(pack.PackageId)) == 1) ||
                        fileAllLines.FirstOrDefault(obj => obj.Contains(appendStr) || obj.Contains(replaceStr) || obj.Contains(deleteStr)) == null)
                    {
                        foreach (var line in fileAllLines)
                        {
                            if (IsCommentLine(line, packfile.FileName))
                            {
                                packfile.CommentCount += 1;
                            }
                            else
                            {
                                packfile.AppendCount += 1;
                            }
                        }
                        //packfile.AppendCount = packfile.LineCount = fileAllLines.Length;
                    }
                    bool isAppend = false;
                    bool isDelete = false;
                    bool isReplace = false;
                    int packIDLen = pack.PackageId.Length;
                    foreach (var line in fileAllLines)
                    {
                        if (line.Contains(pack.PackageId) && line.Contains(appendStr))
                        {
                            isAppend = true;
                        }
                        else if (!line.Contains(pack.PackageId) && isAppend == true)
                        {
                            if (IsCommentLine(line, packfile.FileName))
                            {
                                packfile.CommentCount += 1;
                            }
                            else
                            {
                                packfile.AppendCount += 1;
                            }
                        }
                        else if (line.Contains(pack.PackageId) && isAppend && line.Contains(endStr))
                        {
                            isAppend = false;
                        }

                        if (line.Contains(pack.PackageId) && line.Contains(replaceStr))
                        {
                            isReplace = true;
                        }
                        else if (!line.Contains(pack.PackageId) && isReplace == true)
                        {
                            if (IsCommentLine(line, packfile.FileName))
                            {
                                packfile.CommentCount += 1;
                            }
                            else
                            {
                                packfile.ReplaceCount += 1;
                            }
                        }
                        else if (line.Contains(pack.PackageId) && isReplace && line.Contains(endStr))
                        {
                            isReplace = false;
                        }

                        if (line.Contains(pack.PackageId) && line.Contains(deleteStr))
                        {
                            isDelete = true;
                        }
                        else if (!line.Contains(pack.PackageId) && isDelete == true)
                        {

                            packfile.DeleteCount += 1;
                            packfile.CommentCount += 1;
                        }
                        else if (line.Contains(pack.PackageId) && isDelete && line.Contains(endStr))
                        {
                            isDelete = false;
                        }


                    }
                    pack.AppendCount += packfile.AppendCount;
                    pack.DeleteCount += packfile.DeleteCount;
                    pack.CommentCount += packfile.CommentCount;
                }
            }

        }

        private void ExportToExcel()
        {
            if (packageList.Count == 0)
            {
                MessageBox.Show("Search First,Please!!!!");
                return;
            }

            Type objExcelType = Type.GetTypeFromProgID("Excel.Application");
            if (objExcelType == null)
            {
                MessageBox.Show("Install Excel First,Please!!!!");
                return;
            }

            object objApp = Activator.CreateInstance(objExcelType);

            if (objApp == null)
            {
                MessageBox.Show("Install Excel First,Please!!!!");
                return;
            }


            object xlWorkBooks = objApp.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, objApp, null);
            object xlWorkBook = xlWorkBooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, xlWorkBooks, null);
            object xlSheet = xlWorkBook.GetType().InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, xlWorkBook, null);

            int rowIdx = 0;
            int colIdx = 0;
            object fCell = null;

            ++rowIdx;

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "パッケージID" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "ファイル" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "行数" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "コメント" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "追加" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "削除" });

            fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
            fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { "修正" });

            foreach (var pack in packageList)
            {
                ++rowIdx;

                colIdx = 0;

                object packageIdCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                packageIdCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, packageIdCell, new object[] { pack.PackageId });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Count });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Sum(file => file.CommentCount + file.AppendCount + file.ReplaceCount) });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Sum(file => file.CommentCount) });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Sum(file => file.AppendCount) });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Sum(file => file.DeleteCount) });

                fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++colIdx });
                fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { pack.PackFileList.Sum(file => file.ReplaceCount) });

                foreach (var file in pack.PackFileList)
                {
                    ++rowIdx;

                    int fileCholIdx = 1;

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.FileName });

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.CommentCount + file.AppendCount + file.ReplaceCount });

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.CommentCount });

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.AppendCount });

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.DeleteCount });

                    fCell = xlSheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, xlSheet, new object[] { rowIdx, ++fileCholIdx });
                    fCell.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, fCell, new object[] { file.ReplaceCount });
                }
            }

            objApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, objApp, new object[] { true });

            xlWorkBooks = null;
            objExcelType = null;

            GC.Collect();
        }

        public void GetPackageList()
        {
            var fileInfos = GetFilesCount(new DirectoryInfo(this.textPath.Text));
            this.listBox1.Items.Clear();

            progressBar1.Maximum = fileInfos.Count;
            progressBar1.Value = 0;
            foreach (var item in fileInfos)
            {
                progressBar1.Value = fileInfos.IndexOf(item) + 1;
                Application.DoEvents();
                var c = textPackage.Text;
                try
                {
                    var fileContent = File.ReadAllText(item.FullName);
                    if (fileContent.Contains(textPackage.Text))
                    {
                        if (item.FullName.Contains(@"EXTERN\sql") && item.Extension == ".c")
                        {
                            continue;
                        }
                        if (item.Extension == ".bat")
                        {
                            continue;
                        }
                        var fileAllLines = File.ReadAllLines(item.FullName);
                        foreach (var line1 in fileAllLines)
                        {
                            var pkId = string.Empty;
                            if (line1.Contains(textPackage.Text))
                            {
                                var line = line1.Replace("IBMSE", "");
                                line = line.Substring(line.IndexOf(textPackage.Text));
                                line = line.Replace("*/", "");
                                line = line.TrimEnd(";. ".ToCharArray());
                                line = line.Split("\t .".ToCharArray())[0];
                                if (line.Contains("2020/01/09"))
                                {
                                    line = line.Substring(0, line.IndexOf("2020/01/09"));
                                }
                                if (line.Trim().Last() == 'A')
                                {
                                    pkId = line.Trim().Trim("A /@".ToCharArray());
                                }
                                else if (line.Trim().Last() == 'E')
                                {
                                    pkId = line.Substring(line.IndexOf(textPackage.Text)).Trim('E');

                                }
                                else if (line.Trim().Last() == 'R' || line.Trim().Last() == 'D')
                                {
                                    pkId = line.Replace("#region", "").Replace("#if 0", "").Trim().Trim("<!-RD /@".ToCharArray());

                                }
                                else
                                {
                                    pkId = line.Replace("</term>", "").Substring(line.IndexOf(textPackage.Text)).Split("\t ".ToCharArray())[0];
                                    if (pkId.Contains(' ')) pkId = pkId.Split(' ')[0];
                                    if ("-->".Contains(pkId.Last())) pkId = pkId.Replace("-->", "");
                                    if ("ARDE".Contains(pkId.Last())) pkId = pkId.TrimEnd("ARDE".ToArray());

                                }
                            }
                            if (pkId != string.Empty)
                            {
                                if (!packageList.Exists(d => d.PackageId == pkId))
                                {
                                    var packInfo = new PackageInfo() { PackageId = pkId };
                                    packInfo.PackFileList.Add(new PackFileInfo() { FilePath = item.FullName, FileName = item.Name });
                                    packageList.Add(packInfo);
                                    this.listBox1.Items.Add(pkId);
                                }
                                else
                                {
                                    if (!packageList.Exists(obj => obj.PackageId == pkId && obj.PackFileList.Exists(obj2 => obj2.FilePath == item.FullName)))
                                    {
                                        packageList.Find(d => d.PackageId == pkId).PackFileList.Add(new PackFileInfo() { FilePath = item.FullName, FileName = item.Name });
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    continue;
                }
            }
            packageList.Sort((o1, o2) => string.Compare(o1.PackageId, o2.PackageId));
        }

        public List<FileInfo> GetFilesCount(DirectoryInfo dirInfo)
        {
            var fileInfos = new List<FileInfo>();

            fileInfos.AddRange(dirInfo.GetFiles());

            foreach (DirectoryInfo subdir in dirInfo.GetDirectories())
            {
                fileInfos.AddRange(GetFilesCount(subdir));
            }
            return fileInfos;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var dia = new FileFolderDialog();
            if (dia.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                if (!Directory.Exists(dia.SelectedPath))
                {
                    MessageBox.Show("Select Folder, Please!!!!!");
                    return;
                }
                textPath.Text = dia.SelectedPath;
            }
            // this.textBox1.Text = ExportToExcel().ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            ExportToExcel();
            Cursor = Cursors.Default;

        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {

        }


        //        Range("A1:I1").Select
        //With Selection.Interior
        //    .Pattern = xlSolid
        //    .PatternColorIndex = xlAutomatic
        //    .ThemeColor = xlThemeColorDark1
        //    .TintAndShade = -0.249977111117893
        //    .PatternTintAndShade = 0
        //End With

        //Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        //Selection.Borders(xlDiagonalUp).LineStyle = xlNone

        //With Selection.Borders(xlEdgeLeft)
        //    .LineStyle = xlContinuous
        //    .ColorIndex = 0
        //    .TintAndShade = 0
        //    .Weight = xlThin
        //End With

        //With Selection.Borders(xlEdgeTop)
        //    .LineStyle = xlContinuous
        //    .ColorIndex = 0
        //    .TintAndShade = 0
        //    .Weight = xlThin
        //End With
        //With Selection.Borders(xlEdgeBottom)
        //    .LineStyle = xlContinuous
        //    .ColorIndex = 0
        //    .TintAndShade = 0
        //    .Weight = xlThin
        //End With
        //With Selection.Borders(xlEdgeRight)
        //    .LineStyle = xlContinuous
        //    .ColorIndex = 0
        //    .TintAndShade = 0
        //    .Weight = xlThin
        //End With

        //Columns("A:I").EntireColumn.AutoFit
        //Columns("A:I").EntireColumn.AutoFit
    }
}
