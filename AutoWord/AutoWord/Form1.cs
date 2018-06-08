using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace AutoWord
{
    public partial class MainFormUI : Form
    {
        public MainFormUI()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            //初始化一个OpenFileDialog类
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*.doc|*.docx";

            //判断用户是否正确的选择了文件
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                //获取用户选择文件的后缀名
                string extension = Path.GetExtension(fileDialog.FileName);
                //声明允许的后缀名
                if (extension.EndsWith(".doc") || extension.EndsWith(".docx"))
                {
                    HeadersAndFooters(fileDialog.FileName);
                }
            }
        }

        /// <summary>
                /// 设置文档的标题和页脚
                /// </summary>
                /// <param name="path">文档的路径</param>
        public static bool HeadersAndFooters(string path)
        {
            try
            {
                // 创建新文档
                using (var document = DocX.Create(path))
                {
                    // 这个文档添加页眉和页脚。
                    document.AddHeaders();
                    document.AddFooters();
                    // 强制第一个页面有一个不同的头和脚。
                    document.DifferentFirstPage = true;
                    // 奇偶页页眉页脚不同
                    document.DifferentOddAndEvenPages = true;
                    // 获取本文档的第一个、奇数和甚至是头文件。
                    Header headerFirst = document.Headers.First;
                    Header headerOdd = document.Headers.Odd;
                    Header headerEven = document.Headers.Even;
                    // 获取此文档的第一个、奇数和甚至脚注。
                    Footer footerFirst = document.Footers.First;
                    Footer footerOdd = document.Footers.Odd;
                    Footer footerEven = document.Footers.Even;
                    // 将一段插入到第一个头。
                    Paragraph p0 = headerFirst.InsertParagraph();
                    p0.Append("Hello First Header.").Bold();
                    // 在奇数头中插入一个段落。
                    Paragraph p1 = headerOdd.InsertParagraph();
                    p1.Append("Hello Odd Header.").Bold();
                    // 插入一个段落到偶数头中。
                    Paragraph p2 = headerEven.InsertParagraph();
                    p2.Append("Hello Even Header.").Bold();
                    // 将一段插入到第一个脚注中。
                    Paragraph p3 = footerFirst.InsertParagraph();
                    p3.Append("Hello First Footer.").Bold();
                    // 在奇数脚注中插入一个段落。
                    Paragraph p4 = footerOdd.InsertParagraph();
                    p4.Append("Hello Odd Footer.").Bold();
                    // 插入一个段落到偶数头中。
                    Paragraph p5 = footerEven.InsertParagraph();
                    p5.Append("Hello Even Footer.").Bold();
                    // 在文档中插入一个段落。
                    Paragraph p6 = document.InsertParagraph();
                    p6.AppendLine("Hello First page.");
                    // 创建一个第二个页面，显示第一个页面有自己的头和脚。
                    p6.InsertPageBreakAfterSelf();
                    // 在页面中断后插入一段。
                    Paragraph p7 = document.InsertParagraph();
                    p7.AppendLine("Hello Second page.");
                    // 创建三分之一页面显示，奇偶页不同的页眉和页脚。
                    p7.InsertPageBreakAfterSelf();
                    // 在页面中断后插入一段。
                    Paragraph p8 = document.InsertParagraph();
                    p8.AppendLine("Hello Third page.");
                    // 将属性保存入文档
                    document.Save();
                    return true;
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            //从内存中释放此文档。
        }
    }
}
