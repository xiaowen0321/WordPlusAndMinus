using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordPlusAndMinus
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            //设置页面尺寸及页边距
            float pageWidth = 21.0f;
            float pageHeight = 29.7f;
            float leftMargin = 2.54f;
            float rightMargin = 2.54f;
            float topMargin = 2.54f;
            float bottomMargin = 2.54f;
            float workPageWidth = pageWidth - leftMargin - rightMargin;
            doc.PageSetup.PageWidth = CentermeterToPound(pageWidth);
            doc.PageSetup.PageHeight = CentermeterToPound(pageHeight);
            doc.PageSetup.LeftMargin = CentermeterToPound(leftMargin);
            doc.PageSetup.RightMargin = CentermeterToPound(rightMargin);
            doc.PageSetup.TopMargin = CentermeterToPound(topMargin);
            doc.PageSetup.BottomMargin = CentermeterToPound(bottomMargin);

            //设置制表位
            for (int i = 0; i <= 4; i++)
            {
                doc.Paragraphs[1].Format.TabStops.Add(CentermeterToPound(workPageWidth / 5 * i), 0);
            }

            var focusSelect = Globals.ThisAddIn.Application.Selection;

            focusSelect.Font.Size = 16;
            focusSelect.Font.Name = "宋体";
            focusSelect.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
            focusSelect.TypeText("20以内加减法（100题）");
            focusSelect.TypeParagraph();
            focusSelect.Range.Paragraphs.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            Random r = new Random();
            for (int i = 1; i <= 20; i++)
            {
                for (int j = 1; j <= 5; j++)
                {
                newEquation:
                    int a = r.Next(20);
                    int b = r.Next(20);
                    if (r.Next(2) == 0) //减法
                    {
                        if (a >= b) //第一个数不小于第二个数才能相减
                        {
                            if (a < 10) //如果第一个数小于10，不用借位，可以相减
                            {
                                focusSelect.TypeText($"{a}-{b}=");
                            }
                            else if (a == 10) //如果第一个数为10，第二个数也应该等于10或0，这样相减才不用借位
                            {
                                if (b == 10 || b == 0)
                                {
                                    focusSelect.TypeText($"{a}-{b}=");
                                }
                                else
                                {
                                    goto newEquation;
                                }
                            }
                            else //第一个数大于10时，第二个数必须小于第一个数的个位，这样相减才不用借位
                            {
                                if (b <= a % 10)
                                {
                                    focusSelect.TypeText($"{a}-{b}=");
                                }
                                else
                                {
                                    goto newEquation;
                                }
                            }
                        }
                        else //第一个数小于第二个数，重新开始一个
                        {
                            goto newEquation;
                        }
                    }
                    else //加法
                    {
                        if (a + b < 20)
                        {
                            focusSelect.TypeText($"{a}+{b}=");
                        }
                        else
                        {
                            goto newEquation;
                        }
                    }
                    if (j != 5)
                    {
                        focusSelect.TypeText("\t");
                    }
                }
                focusSelect.TypeParagraph();
            }
        }

        private float PoundToCentermeter(float pound)
        {
            return pound * (float)2.54 / 72;
        }
        private float CentermeterToPound(float centermeter)
        {
            return centermeter / (float)2.54 * 72;
        }

    }
}
