using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc.Documents;
using Spire.Doc;
using Spire.Xls;
using Spire.Doc.Fields;
using System.Drawing;

namespace 倒表工具
{
    public class ExcelOperation
    {
        public ExcelOperation(String filename)
        {
            this.filename = filename;
        }

        String filename = string.Empty;

        public void Excel2Docx()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filename);


            Document doc = new Document();
            Section section = doc.AddSection();
            //指定表格字体及大小
            ParagraphStyle sty = new ParagraphStyle(doc);
            sty.Name = "fortable";
            sty.CharacterFormat.FontName = "宋体";
            sty.CharacterFormat.FontSize = 12;
            doc.Styles.Add(sty);

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine(sheet.Name);
                if (sheet == workbook.Worksheets[0])
                    continue;
                else
                {

                    String headText = sheet.Range["B2"].Text + "  " + sheet.Range["B3"].Value;

                    Paragraph paragraph = section.AddParagraph();
                    paragraph.AppendText(headText);
                    paragraph.ApplyStyle(BuiltinStyle.Heading2);

                    //添加表格并将相应的位置合并
                    Table table = section.AddTable(true);
                    table.ResetCells(10, 6);

                    //前两行
                    TableRow row = table.Rows[0];
                    TextRange range = row.Cells[0].AddParagraph().AppendText("SIF描述");
                    range = row.Cells[3].AddParagraph().AppendText(sheet.Range["B4"].Text);
                    row = table.Rows[1];
                    range = row.Cells[0].AddParagraph().AppendText("事件后果");
                    range = row.Cells[3].AddParagraph().AppendText(sheet.Range["B5"].Text);

                    //危害程度
                    range = table.Rows[2].Cells[0].AddParagraph().AppendText("危害程度");
                    range = table.Rows[2].Cells[1].AddParagraph().AppendText("人员");
                    range = table.Rows[3].Cells[1].AddParagraph().AppendText("环境");
                    range = table.Rows[4].Cells[1].AddParagraph().AppendText("财产");
                    String str = sheet.GetText(42, 2);
                    if (str != null)
                    {
                        str = str.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");

                        range = table[2, 3].AddParagraph().AppendText(str);
                    }
                    str = sheet.GetText(43, 2);
                    if (str != null)
                    {
                        str = str.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");

                        range = table.Rows[3].Cells[3].AddParagraph().AppendText(str);
                    }
                    str = sheet.GetText(44, 2);
                    if(str != null)
                    {
                        str = str.Replace("：", "").Replace("0", "").Replace("N", "").Replace("L", "").Replace("M", "").Replace("H", "").Replace("E", "");
                        range = table[4, 3].AddParagraph().AppendText(str);
                    }
                    

                    //保护层&减缓措施
                    range = table[5, 0].AddParagraph().AppendText("独立保护层");
                    str = sheet.Range["D22"].Text + sheet.Range["D23"].Text + sheet.Range["D24"].Text;
                    range = table[5, 3].AddParagraph().AppendText(str);
                    range = table[6, 0].AddParagraph().AppendText("减缓措施");
                    str = sheet.Range["D12"].Text + sheet.Range["D13"].Text;
                    range = table[6, 3].AddParagraph().AppendText(str);

                    //SIL分项定级
                    range = table[7, 0].AddParagraph().AppendText("SIL分项定级");
                    range = table[7, 2].AddParagraph().AppendText("人员");
                    range = table[8, 2].AddParagraph().AppendText("环境");
                    range = table[9, 2].AddParagraph().AppendText("财产");
                    
                    str = string.Empty;
                    string[] head = new string[6] { "B", "C", "D", "E", "F", "G" };
                    int[] num = new int[4] { 30, 33, 36, 39 };
                    foreach (int i in num)
                    {
                        //人员
                        foreach (string h in head)
                        {
                            Console.WriteLine(sheet.Range[h + i.ToString()].Style.Color.ToString() + "\t" + h + i.ToString());
                            if (sheet.Range[h + i.ToString()].Style.Color.ToString() == @"Color [A=255, R=255, G=0, B=0]")
                            {
                                str = sheet.Range[h + i.ToString()].Text;
                                if (str == null || str.Replace(" ", "") == string.Empty )
                                {
                                    range = table[7, 3].AddParagraph().AppendText(str);
                                }
                                else
                                    range = table[7, 3].AddParagraph().AppendText("NA");
                                break;
                            }
                        }
                        if(str != string.Empty)
                        {
                            //环境
                            int l = i+1;
                            foreach (string h in head)
                            {
                                if (sheet.Range[h + l.ToString()].Style.Color == Color.Red)
                                {
                                    str = sheet.Range[h + l.ToString()].Text;
                                    if (str.Replace(" ", "") == string.Empty)
                                    {
                                        range = table[8, 3].AddParagraph().AppendText(str);
                                    }
                                    else
                                        range = table[8, 3].AddParagraph().AppendText("NA");
                                    break;
                                }
                            }
                            //财产
                            l++;
                            foreach (string h in head)
                            {
                                if (sheet.Range[h + l.ToString()].Style.Color == Color.Red)
                                {
                                    str = sheet.Range[h + l.ToString()].Text;
                                    if (str.Replace(" ", "") == string.Empty)
                                    {
                                        range = table[9, 3].AddParagraph().AppendText(str);
                                    }
                                    else
                                        range = table[9, 3].AddParagraph().AppendText("NA");
                                    break;
                                }
                            }
                        }
                    }
                    
                    
                    //最终定级
                    range = table[7, 4].AddParagraph().AppendText("SIL定级");
                    range = table[7, 5].AddParagraph().AppendText(sheet.GetText(44, 8));
                    
                    foreach (TableRow rowl in table.Rows)
                    {
                        foreach (TableCell cell in rowl.Cells)
                        {
                            foreach (Paragraph pg in cell.Paragraphs)
                            {
                                pg.ApplyStyle(sty.Name);
                            }
                        }

                    }

                    table.ApplyHorizontalMerge(0, 0, 2);
                    table.ApplyHorizontalMerge(0, 3, 5);
                    table.ApplyHorizontalMerge(1, 0, 2);
                    table.ApplyHorizontalMerge(5, 0, 2);
                    table.ApplyHorizontalMerge(6, 0, 2);
                    table.ApplyHorizontalMerge(1, 3, 5);
                    table.ApplyHorizontalMerge(2, 3, 5);
                    table.ApplyHorizontalMerge(3, 3, 5);
                    table.ApplyHorizontalMerge(4, 3, 5);
                    table.ApplyHorizontalMerge(5, 3, 5);
                    table.ApplyHorizontalMerge(6, 3, 5);
                    table.ApplyVerticalMerge(0, 2, 4);
                    table.ApplyHorizontalMerge(7, 0, 1);
                    table.ApplyHorizontalMerge(8, 0, 1);
                    table.ApplyHorizontalMerge(9, 0, 1);
                    table.ApplyVerticalMerge(0, 7, 9);
                    table.ApplyVerticalMerge(4, 7, 9);
                    table.ApplyVerticalMerge(5, 7, 9);

                    DefaultTableStyle style = DefaultTableStyle.TableGrid;
                    table.ApplyStyle(style);
                }
            }

            doc.SaveToFile(filename.Replace(".xlsm",".docx").Replace(".xls",".docx").Replace(".xlsx",".docx"),Spire.Doc.FileFormat.Docx2013);
            
        }
    }
}
