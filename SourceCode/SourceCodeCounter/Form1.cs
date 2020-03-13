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
            textPath.Text = Properties.Settings.Default.FilePath;
            textPackage.Text = Properties.Settings.Default.PackageBefore;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.FilePath = textPath.Text;
                Properties.Settings.Default.PackageBefore = textPackage.Text;
                Properties.Settings.Default.Save();

                button1.Enabled = false;
                if (this.textPath.Text == string.Empty || this.textPackage.Text == string.Empty)
                {
                    return;
                }

                this.packageList.Clear();
                this.textBox1.Text = string.Empty;

                GetPackageList();

                this.listBox1.Items.Clear();
                foreach (var item in packageList)
                {
                    this.listBox1.Items.Add(item.PackageId);
                }

                GetPackageDetail();

                if (listBox1.Items.Count > 0)
                {
                    listBox1.SelectedIndex = 0;
                    listBox1_DoubleClick(listBox1, new EventArgs());
                }
            }
            finally
            {
                button1.Enabled = true;
            }

        }

        public bool IsCommentLine(string line, string fileName, ref bool isCommntStarted)
        {
            bool ret = false;
            line = line.Trim().Trim('\t');
            if (fileName.Contains(".cs"))
            {
                ret = line.Contains("//") && line.Trim().Substring(0, 2) == "//";
            }
            else if (fileName.Contains(".xaml") || fileName.Contains(".xml"))
            {
                ret = line.Contains("<!--") && line.Trim().Substring(0, 4) == "<!--";
                if (ret)
                {
                    if (!line.Contains("-->") && line.Trim().Substring(line.Trim().Length - 3, 3) == "-->")
                    {
                        isCommntStarted = true;
                    }
                }
                else
                {
                    if (isCommntStarted && !(line.Contains("-->") && line.Trim().Substring(line.Trim().Length - 3, 3) == "-->"))
                    {
                        ret = true;
                        isCommntStarted = false;
                    }
                }
                if (!ret)
                {
                    ret = isCommntStarted;
                }
            }
            else if (fileName.Contains(".properties"))
            {
                ret = line.Contains("#") && line.Trim().Substring(0, 1) == "#";
            }
            else if (fileName.Contains(".cpp") || fileName.Contains(".h") || fileName.Contains(".java") || fileName.Contains(".js"))
            {
                ret = line.Contains("//") && line.Trim().Substring(0, 2) == "//";
                if (!ret)
                {
                    ret = line.Contains("/*") && line.Trim().Substring(0, 2) == "/*";
                    if (ret) isCommntStarted = true;
                }
                if (!ret)
                {
                    ret = line.Contains("*/") && line.Trim().Substring(line.Trim().Length - 2, 2) == "*/";
                    if (ret) isCommntStarted = false;
                }
                if (fileName.Contains(".cpp") || fileName.Contains(".h"))
                {
                    if (isCommntStarted && line.Contains("*/") && line.Trim().Substring(line.Trim().Length - 2, 2) == "*/")
                    {
                        isCommntStarted = false;
                    }
                }
                if (!ret)
                {
                    ret = isCommntStarted;
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
                var propAppendStr = pack.PackageId + " ADD";
                var propEndStr = pack.PackageId + " END";
                var titleStr = pack.PackageId + "\t";
                var titleStrCs = "<term>" + pack.PackageId + "</term>";

                foreach (var packfile in pack.PackFileList)
                {
                    label3.Text = packfile.FilePath;
                    Application.DoEvents();

                    var fileAllLines = File.ReadAllLines(packfile.FilePath);

                    // When Created File
                    if ((fileAllLines.Count(obj => obj.Contains(pack.PackageId)) == 1) ||
                        fileAllLines.FirstOrDefault(obj =>
                            obj.Contains(appendStr) ||
                            obj.Contains(replaceStr) ||
                            obj.Contains(deleteStr) ||
                            obj.Contains(propAppendStr) ||
                            obj.Contains(propEndStr)
                            ) == null)
                    {
                        bool isCommntStarted = false;
                        foreach (var line in fileAllLines)
                        {
                            var isComment = IsCommentLine(line, packfile.FileName, ref isCommntStarted);
                            if (isComment || isCommntStarted)
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
                    else
                    {
                        bool isAppend = false;
                        bool isDelete = false;
                        bool isReplace = false;
                        int packIDLen = pack.PackageId.Length;

                        bool isPropAppend = false;
                        bool isCommntStarted = false;

                        foreach (var line in fileAllLines)
                        {
                            var isComment = IsCommentLine(line, packfile.FileName, ref isCommntStarted);
                            if (isCommntStarted && line.Contains(pack.PackageId))
                            {
                                packfile.CommentCount += 1;
                            }
                            else
                            {
                                if (line.Contains(pack.PackageId) && (line.Contains(titleStr) || line.Contains(titleStrCs)))
                                {
                                    packfile.CommentCount += 1;
                                }
                            }


                            if (line.Contains(pack.PackageId) && line.Contains(appendStr))
                            {
                                packfile.CommentCount += 1;
                                isAppend = true;
                            }
                            else if (!line.Contains(pack.PackageId) && isAppend == true)
                            {
                                if (IsCommentLine(line, packfile.FileName, ref isCommntStarted))
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
                                packfile.CommentCount += 1;
                                isAppend = false;
                            }
                            // .properties file
                            if (line.Contains(pack.PackageId) && line.Contains(propAppendStr))
                            {
                                packfile.CommentCount += 1;
                                isPropAppend = true;
                            }
                            else if (!line.Contains(pack.PackageId) && isPropAppend == true)
                            {
                                if (IsCommentLine(line, packfile.FileName, ref isCommntStarted))
                                {
                                    packfile.CommentCount += 1;
                                }
                                else
                                {
                                    packfile.AppendCount += 1;
                                }
                            }
                            else if (line.Contains(pack.PackageId) && isPropAppend && line.Contains(propEndStr))
                            {
                                packfile.CommentCount += 1;
                                isPropAppend = false;
                            }

                            if (line.Contains(pack.PackageId) && line.Contains(replaceStr))
                            {
                                packfile.CommentCount += 1;
                                isReplace = true;
                            }
                            else if (!line.Contains(pack.PackageId) && isReplace == true)
                            {
                                if (IsCommentLine(line, packfile.FileName, ref isCommntStarted))
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
                                packfile.CommentCount += 1;
                                isReplace = false;
                            }

                            if (line.Contains(pack.PackageId) && line.Contains(deleteStr))
                            {
                                packfile.CommentCount += 1;
                                isDelete = true;
                            }
                            else if (!line.Contains(pack.PackageId) && isDelete == true)
                            {

                                packfile.DeleteCount += 1;
                                packfile.CommentCount += 1;
                            }
                            else if (line.Contains(pack.PackageId) && isDelete && line.Contains(endStr))
                            {
                                packfile.CommentCount += 1;
                                isDelete = false;
                            }


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
                        if (item.FullName.Contains(@"\classes\") || item.FullName.Contains(@"\bin\"))
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
            if (listBox1.SelectedIndex < 0)
            {
                return;
            }

            var package = packageList.Find(obj => obj.PackageId == listBox1.SelectedItem.ToString());

            if (package == null)
            {
                return;
            }

            textBox1.Text = string.Empty;
            textBox1.Text += string.Format("ファイル数：\t\t{0}\r\n", package.PackFileList.Count);
            textBox1.Text += string.Format("行数：\t\t\t{0}\r\n", package.PackFileList.Sum(file => file.CommentCount + file.AppendCount + file.ReplaceCount));
            textBox1.Text += string.Format("コメント行数：\t\t{0}\r\n", package.PackFileList.Sum(file => file.CommentCount));
            textBox1.Text += string.Format("追加行数：\t\t{0}\r\n", package.PackFileList.Sum(file => file.AppendCount));
            textBox1.Text += string.Format("削除行数：\t\t{0}\r\n", package.PackFileList.Sum(file => file.DeleteCount));
            textBox1.Text += string.Format("修正行数：\t\t{0}\r\n", package.PackFileList.Sum(file => file.ReplaceCount));
            textBox1.Text += Environment.NewLine;

            textBox1.Text += "===============================";
            textBox1.Text += Environment.NewLine;
            foreach (var file in package.PackFileList)
            {
                textBox1.Text += string.Format("{0,-50}", file.FileName);
                textBox1.Text += string.Format("行数({0,5})", file.CommentCount + file.AppendCount + file.ReplaceCount).PadRight(20);
                textBox1.Text += string.Format("コメント行数({0,5})", file.CommentCount).PadRight(20);
                textBox1.Text += string.Format("追加行数({0,5})", file.AppendCount).PadRight(20);
                textBox1.Text += string.Format("削除行数({0,5})", file.DeleteCount).PadRight(20);
                textBox1.Text += string.Format("修正行数({0,5})", file.ReplaceCount).PadRight(20);
                textBox1.Text += Environment.NewLine;

            }
            textBox1.Text += Environment.NewLine;

        }

        public string PadRightA(string input, int len)
        {
            var lengthBtye = Encoding.Default.GetByteCount(input);
            var output = input.PadRight(len - (lengthBtye - input.Length));
            return output;
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            listBox1_DoubleClick(sender, new EventArgs());
        }
    }
}
