using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Helpers;

namespace xuatfileWord.Util
{
    public class HtmlToWord
    {
        public static byte[] HtmlToWordMethod(String html)
        {
            string color = "#000";
            string width = "5000";
            const string filename = "test.docx";
            if (File.Exists(filename)) File.Delete(filename);

            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(
                       generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        new Document(new Body()).Save(mainPart);
                    }

                    //HtmlConverter converter = new HtmlConverter(mainPart);
                    Body body = mainPart.Document.Body;

                    //var paragraphs = converter.Parse(html);
                    //for (int i = 0; i < paragraphs.Count; i++)
                    //{
                    //    body.Append(paragraphs[i]);
                    //}
                    Table table = new Table();
                    //// Create the table properties

                    TableProperties tblProperties = new TableProperties();



                    //// Create Table Borders

                    TableBorders tblBorders = new TableBorders();



                    TopBorder topBorder = new TopBorder();

                    topBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    topBorder.Color = color;

                    tblBorders.AppendChild(topBorder);



                    BottomBorder bottomBorder = new BottomBorder();

                    bottomBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    bottomBorder.Color = color;

                    tblBorders.AppendChild(bottomBorder);



                    RightBorder rightBorder = new RightBorder();

                    rightBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    rightBorder.Color = color;

                    tblBorders.AppendChild(rightBorder);



                    LeftBorder leftBorder = new LeftBorder();

                    leftBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    leftBorder.Color = color;

                    tblBorders.AppendChild(leftBorder);



                    InsideHorizontalBorder insideHBorder = new InsideHorizontalBorder();

                    insideHBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    insideHBorder.Color = color;

                    tblBorders.AppendChild(insideHBorder);



                    InsideVerticalBorder insideVBorder = new InsideVerticalBorder();

                    insideVBorder.Val = new EnumValue<BorderValues>(BorderValues.Thick);

                    insideVBorder.Color = color;

                    tblBorders.AppendChild(insideVBorder);



                    //// Add the table borders to the properties

                    tblProperties.AppendChild(tblBorders);
                    TableWidth tableWidth = new TableWidth() { Width = width, Type = TableWidthUnitValues.Pct };

                    TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

                    tblProperties.AppendChild(tableStyle);
                    tblProperties.AppendChild(tableWidth);
                    //// Add the table properties to the table

                    table.AppendChild(tblProperties);

                    TableRow tr1 = new TableRow();
                    TableCell tc11 = new TableCell();
                    Paragraph p11 = new Paragraph(new Run(new Text("ID")));
                    tc11.Append(p11);
                    tr1.Append(tc11);

                    TableCell tc12 = new TableCell();
                    Paragraph p12 = new Paragraph();
                    Run r12 = new Run();
                    RunProperties rp12 = new RunProperties();
                    rp12.Bold = new Bold();
                    r12.Append(rp12);
                    r12.Append(new Text("Nice"));
                    p12.Append(r12);
                    tc12.Append(p12);

                    tr1.Append(tc12);
                    table.Append(tr1);

                    TableRow tr2 = new TableRow();


                    TableCell tc21 = new TableCell();
                    Paragraph p21 = new Paragraph(new Run(new Text("Name")));
                    tc21.Append(p21);
                    tr2.Append(tc21);

                    TableCell tc22 = new TableCell();
                    Paragraph p22 = new Paragraph();
                    ParagraphProperties pp22 = new ParagraphProperties();
                    pp22.Justification = new Justification() { Val = JustificationValues.Center };
                    p22.Append(pp22);
                    p22.Append(new Run(new Text("Table")));
                    tc22.Append(p22);

                    tr2.Append(tc22);
                    table.Append(tr2);

                    // Add your table to docx body
                    body.Append(table);


                    //Save
                    mainPart.Document.Save();
                }

                return generatedDocument.ToArray();
            }
        }
    }
}
