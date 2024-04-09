
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using OfficeOpenXml;
using static OfficeOpenXml.ExcelErrorValue;
using MessageBox = System.Windows.Forms.MessageBox;

namespace ExcelCV
{

    /*1. 打开 Visual Studio，选择 "工具" -> "NuGet 包管理器" -> "程序包管理器控制台"，然后手动执行安装命令：
    *           Install-Package EPPlus
    */
    public partial class 钟志平Ver2 : Form
    {
        private string excelAFilePath; 
        private string excelBFilePath;
        private Dictionary<string, string> rowMapping;

        public 钟志平Ver2()
        {
            InitializeComponent();
        }

        private void ButtonAClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialogA = new OpenFileDialog())
            {
                openFileDialogA.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                openFileDialogA.RestoreDirectory = true;

                if (openFileDialogA.ShowDialog() == DialogResult.OK)
                {
                    // A
                    excelAFilePath = openFileDialogA.FileName;

                    // B
                    using (OpenFileDialog openFileDialogB = new OpenFileDialog())
                    {
                        openFileDialogB.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                        openFileDialogB.RestoreDirectory = true;
                        if (openFileDialogB.ShowDialog() == DialogResult.OK)
                        {
                            excelBFilePath = openFileDialogB.FileName;
                        }
                    }


                    if (string.IsNullOrEmpty(excelAFilePath) || string.IsNullOrEmpty(excelBFilePath))
                    {
                        MessageBox.Show("请先选择 AExcel 和 BExcel 文件。");
                        return;
                    }

                    // 在处理Excel文件之前设置LicenseContext
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    rowMapping = new Dictionary<string, string>
                    {
                        {"A", "A" },   // 日期
                        {"C", "E" },    
                        {"D", "G" },
                        {"E", "K" },
                        {"F", "L" },
                        {"G", "D" },
                        {"H", "H" },
                        {"I", "I" },
                        {"J", "J" },
                        {"K", "O" },
                         {"L", "P" },

                    };

                    ProcessExcelFiles(excelAFilePath, 0, excelBFilePath , 0, rowMapping, 1);
                    // 在处理完后将LicenseContext重置为Default
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                }
            }
        }
        private void ButtonBClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialogA = new OpenFileDialog())
            {
                openFileDialogA.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                openFileDialogA.RestoreDirectory = true;

                if (openFileDialogA.ShowDialog() == DialogResult.OK)
                {
                    // A
                    excelAFilePath = openFileDialogA.FileName;

                    // B
                    using (OpenFileDialog openFileDialogB = new OpenFileDialog())
                    {
                        openFileDialogB.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                        openFileDialogB.RestoreDirectory = true;
                        if (openFileDialogB.ShowDialog() == DialogResult.OK)
                        {
                            excelBFilePath = openFileDialogB.FileName;
                        }
                    }


                    if (string.IsNullOrEmpty(excelAFilePath) || string.IsNullOrEmpty(excelBFilePath))
                    {
                        MessageBox.Show("请先选择 AExcel 和 BExcel 文件。");
                        return;
                    }

                    // 在处理Excel文件之前设置LicenseContext
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    rowMapping = new Dictionary<string, string>
                    {
                        {"C", "A" },
                        {"D", "C" },
                        {"E", "D" },
                        {"F", "E" },
                        {"H", "F" },
                        {"I", "G" },
                        {"J", "H" },
                        {"L", "I" },
                        {"M", "M" },
                        {"N", "J" },
                        {"O", "K" },
                        {"Q", "L" },

                    };

                    ProcessExcelFiles(excelAFilePath , 0 ,excelBFilePath , 1 , rowMapping , 1);
                    // 在处理完后将LicenseContext重置为Default
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                }
            }
        }
        private void ButtonCClick(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialogA = new OpenFileDialog())
            {
                openFileDialogA.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                openFileDialogA.RestoreDirectory = true;

                if (openFileDialogA.ShowDialog() == DialogResult.OK)
                {
                    // A
                    excelAFilePath = openFileDialogA.FileName;

                    // B
                    using (OpenFileDialog openFileDialogB = new OpenFileDialog())
                    {
                        openFileDialogB.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
                        openFileDialogB.RestoreDirectory = true;
                        if (openFileDialogB.ShowDialog() == DialogResult.OK)
                        {
                            excelBFilePath = openFileDialogB.FileName;
                        }
                    }


                    if (string.IsNullOrEmpty(excelAFilePath) || string.IsNullOrEmpty(excelBFilePath))
                    {
                        MessageBox.Show("请先选择 AExcel 和 BExcel 文件。");
                        return;
                    }

                    // 在处理Excel文件之前设置LicenseContext
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    rowMapping = new Dictionary<string, string>
                    {
                        {"A", "A" },
                        {"B", "B" },
                        {"C", "C" },
                        {"E", "O" },
                        {"G", "J" },
                        {"M", "M" },
                        {"K", "L" },
                        {"L", "H" },
                        {"F", "D" },

                    };

                    ProcessExcelFiles(excelAFilePath , 0 ,excelBFilePath , 2 , rowMapping , 3);
                    // 在处理完后将LicenseContext重置为Default
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                }
            }
        }
        
        static int LetterToInt(string letter)
        {
            if (string.IsNullOrEmpty(letter) || letter.Length != 1 || !char.IsUpper(letter[0]))
            {
                throw new ArgumentException("输入的不是大写字母");
            }

            char upperCaseA = 'A';
            int intValueA = (int)upperCaseA;

            char upperCaseLetter = char.ToUpper(letter[0]);
            int intValueLetter = (int)upperCaseLetter;

            // 将字母的ASCII码值减去A的ASCII码值，并加1
            int result = intValueLetter - intValueA + 1;

            return result;
        }
        private void ProcessExcelFiles(string aFilePath, int asheet , string bFilePath , int bsheet , Dictionary<string,string> rowMapping , int typeCase)
        {
            MessageBoxResult result = (MessageBoxResult)MessageBox.Show($"源文件【{aFilePath}】\n↓↓↓↓↓↓↓复制↓↓↓↓↓↓↓\n目标文件【{bFilePath}】\n策略（1添加/2覆盖/3已有数据忽略)：该命令类型{typeCase}", "提示", (MessageBoxButtons)MessageBoxButton.YesNo);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    using (var packageA = new ExcelPackage(new System.IO.FileInfo(aFilePath)))
                    using (var packageB = new ExcelPackage(new System.IO.FileInfo(bFilePath)))
                    {
                        var worksheetA = packageA.Workbook.Worksheets[asheet];
                        var worksheetB = packageB.Workbook.Worksheets[bsheet];   // sheet1/2/3  package.Workbook.Worksheets[1]  or  package.Workbook.Worksheets["Sheet2"]

                        int rowCountA = worksheetA.Dimension.End.Row;   // 从第二行开始 每行数据要拿着CV了

                        // 循环处理 AExcel 中除了第一行的每一行数据
                        for (int rowA = 2; rowA <= rowCountA; rowA++)
                        {
                            // 将数据复制到 BExcel 中的指定列
                            int rowB = 0;
                            bool isRowHidden = worksheetA.Row(rowA).Hidden;
                            switch (typeCase)
                            {
                                case 1:     // add 末尾
                                    rowB = worksheetB.Dimension.End.Row + 1;
                                    break;
                                case 2:     // repalce 替换
                                    int maxDataRowB = worksheetB.Dimension.End.Row + 1;
                                    for (int i = 0; i < maxDataRowB; i++)
                                    {
                                        if (worksheetB.Cells[i, LetterToInt("A")].Value == worksheetA.GetValue(rowA, LetterToInt("A")))
                                        {
                                            rowB = i;
                                            break;
                                        }
                                    }
                                    break;
                                case 3:     // 同日期（"A"）则忽略
                                    int maxB = worksheetB.Dimension.End.Row + 1;
                                    bool isSame = false;
                                    for (int i = 1; i < maxB; i++)  // 默认元素[0] 是标题
                                    {
                                        /*
                                         * 日期格式： Excel 中日期的显示格式可能不同，但实际存储的日期值是相等的。确保你比较的是日期的实际值，而不是格式化后的字符串。可以尝试将日期值转换为字符串进行比较，或者直接比较日期值。
                                         * 日期类型： 确保你在比较时考虑到日期的时间部分。如果两个日期值的时间部分不同，它们将被视为不相等。你可能需要将日期值转换为只包含日期部分的格式再进行比较。
                                         * 单元格类型： 确保你使用的 GetValue 方法正确地返回日期值的数据类型。你可以将获取的日期值强制转换为 DateTime 类型进行比较。
                                         */
                                        object shtBValue = worksheetB.GetValue(i, LetterToInt("A"));
                                        object shtAValue = worksheetA.GetValue(rowA, LetterToInt("A"));
                                        Debug.WriteLine($"第{i}行，是否隐藏{isRowHidden}，数值对比值{shtBValue} == {shtAValue}");
                                        if (!isRowHidden && shtAValue is DateTime dateA && shtBValue is DateTime dateB)
                                        {
                                            isSame = dateA.Date == dateB.Date;
                                            if (isSame)
                                            {

                                                break;
                                            }
                                        }
                                    }
                                    if (!isSame && !isRowHidden)
                                    {
                                        rowB = worksheetB.Dimension.End.Row + 1;
                                    }
                                    break;
                                default:
                                    break;
                            }

                            // 在这里根据实际需求设置要复制到 BExcel 中的列
                            if (rowB > 0)
                            {
                                foreach (var kvp in rowMapping)
                                {
                                    // 将大写字母转换为整数
                                    int key = LetterToInt(kvp.Key);
                                    int value = LetterToInt(kvp.Value);
                                    worksheetB.Cells[rowB, value].Value = worksheetA.GetValue(rowA, key);
                                }
                            }
                        }

                        packageB.Save();
                        MessageBox.Show("操作成功完成。");

                        OpenExcelFile(bFilePath);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"处理 Excel 文件时发生错误：{ex.Message}");
                }
            }
        }

        private void OpenExcelFile(string filePath)
        {
            try
            {
                Process.Start(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"无法打开 Excel 文件：{ex.Message}");
            }
        }
    }
}
