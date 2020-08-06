using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace GHIAProj
{
    public class Exporter
    {
        bool isTransgr = false;
        public Exporter(bool _isCommon) { this.isTransgr = _isCommon; }
        public void ToWord(DataGridView DGV, string filename)
        {
            if (DGV.RowCount != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop
                Word.Document oDoc = new Word.Document();
                oDoc.Application.Visible = true;               
                oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;//page orintation
                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //table format
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;
                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                  Type.Missing, Type.Missing, ref ApplyBorders,
                                  Type.Missing, Type.Missing, Type.Missing,
                                  Type.Missing, Type.Missing, Type.Missing,
                                  Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);
                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Font.Name = "Times New Roman";//
                oDoc.Application.Selection.Font.Size = 9; //
                oDoc.Application.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                oDoc.Application.Selection.Paragraphs.Space1();
                oDoc.Application.Selection.Paragraphs.SpaceAfter = 0;
                oDoc.Application.Selection.Paragraphs.SpaceBefore = 0;
                if (!isTransgr) { CommonTableFormat(DGV, RowCount, ColumnCount, oDoc); }
                else { TransgrTableFormat(DGV, RowCount, ColumnCount, oDoc); }
            }
        }
        private void CommonTableFormat(DataGridView DGV, int RowCount, int ColumnCount, Document oDoc)
        {
            //Center column with numbers
            oDoc.Application.Selection.Tables[1].Columns[4].Select();
            oDoc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            oDoc.Application.Selection.Tables[1].Columns[5].Select();
            oDoc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;        
            AddHeaders(DGV, ColumnCount, oDoc);//Headers
            //Add header to columns 5,6
            oDoc.Application.Selection.InsertRowsBelow(1);
            oDoc.Application.Selection.Tables[1].Cell(2, 5).Range.Text = oDoc.Application.Selection.Tables[1].Cell(1, 5).Range.Text;
            oDoc.Application.Selection.Tables[1].Cell(2, 6).Range.Text = oDoc.Application.Selection.Tables[1].Cell(1, 6).Range.Text;
            for (int i = 1; i < 5; i++)
            {
                oDoc.Application.Selection.Tables[1].Cell(1, i).Merge(oDoc.Application.Selection.Tables[1].Cell(2, i));
            }
            oDoc.Application.Selection.Tables[1].Cell(1, 7).Merge(oDoc.Application.Selection.Tables[1].Cell(2, 7));
            oDoc.Application.Selection.Tables[1].Cell(1, 8).Merge(oDoc.Application.Selection.Tables[1].Cell(2, 8));
            oDoc.Application.Selection.Tables[1].Cell(1, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 6));
            for(int i=1;i<=4;i++)
            {
                MergeEmptyCells(RowCount, oDoc, i);
            }
            MergeEmptyCells(RowCount, oDoc, 8);
            int checkCell = 3;
            for (int i = 3; i <= RowCount - 1; i++)
            {
                if (oDoc.Application.Selection.Tables[1].Cell(i + 1, 5).Range.Text.Length < 3)
                {                   
                    oDoc.Application.Selection.Tables[1].Cell(checkCell, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(i + 1, 5));
                    oDoc.Application.Selection.Tables[1].Cell(checkCell, 6).Merge(oDoc.Application.Selection.Tables[1].Cell(i + 1, 6));
                }
                else { checkCell = i + 1; }
            }
            FindAndReplace(oDoc.Application, "^13", ", ");
            oDoc.Application.Selection.Tables[1].Cell(1, 5).Range.Text = "Максимальная" + "\n" + "концентрация";
            oDoc.Application.Selection.Tables[1].Cell(1, 4).Range.Text = "Число" + "\n" + "случаев" + "\n" + "ВЗ";
            oDoc.Application.Selection.Tables[1].Cell(1, 7).Range.Text = "Субъект" + "\n" + "Российской Федерации";
            FindAndReplace(oDoc.Application, "ПДК, ", "ПДК");
            FindAndReplace(oDoc.Application, "Дата, ", "Дата");
        }
        private void TransgrTableFormat(DataGridView DGV, int RowCount, int ColumnCount, Document oDoc)
        {
            //Center column with numbers
            for(int i=4;i<=6;i++)
            {
                oDoc.Application.Selection.Tables[1].Columns[i].Select();
                oDoc.Application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }
            AddHeaders(DGV, ColumnCount, oDoc);//Headers
            //Add header to columns 5,6
            oDoc.Application.Selection.InsertRowsBelow(1);
            oDoc.Application.Selection.Tables[1].Cell(2, 5).Range.Text = oDoc.Application.Selection.Tables[1].Cell(1, 5).Range.Text;
            oDoc.Application.Selection.Tables[1].Cell(2, 6).Range.Text = oDoc.Application.Selection.Tables[1].Cell(1, 6).Range.Text;
            for (int i = 1; i < 5; i++)
            {
                oDoc.Application.Selection.Tables[1].Cell(1, i).Merge(oDoc.Application.Selection.Tables[1].Cell(2, i));
            }
            oDoc.Application.Selection.Tables[1].Cell(1, 7).Merge(oDoc.Application.Selection.Tables[1].Cell(2, 7));
            oDoc.Application.Selection.Tables[1].Cell(1, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 6));
            for (int i = 1; i <= 4; i++)
            {
                MergeEmptyCells(RowCount, oDoc, i);
            }
            int checkCell = 3;
            for (int i = 3; i <= RowCount - 1; i++)
            {
                if (oDoc.Application.Selection.Tables[1].Cell(i + 1, 5).Range.Text.Length < 3)
                {
                    oDoc.Application.Selection.Tables[1].Cell(checkCell, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(i + 1, 5));
                    oDoc.Application.Selection.Tables[1].Cell(checkCell, 6).Merge(oDoc.Application.Selection.Tables[1].Cell(i + 1, 6));
                }
                else { checkCell = i + 1; }
            }
            FindAndReplace(oDoc.Application, "^13", ", ");
            oDoc.Application.Selection.Tables[1].Cell(1, 5).Range.Text = "Максимальная" + "\n" + "концентрация";
            oDoc.Application.Selection.Tables[1].Cell(1, 4).Range.Text = "Число" + "\n" + "случаев" + "\n" + "ВЗ";
            FindAndReplace(oDoc.Application, "ПДК, ", "ПДК");
            FindAndReplace(oDoc.Application, "Дата, ", "Дата");
        }
        private void AddHeaders(DataGridView DGV, int ColumnCount, Document oDoc)
        {
            //Add row for headers
            oDoc.Application.Selection.Tables[1].Rows[1].Select();
            oDoc.Application.Selection.InsertRowsAbove(1);
            //Border`s style
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Tables[1].Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            oDoc.Application.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //Add table headers
            oDoc.Application.Selection.Tables[1].Rows[1].Select();
            for (int c = 0; c <= ColumnCount - 1; c++)
            {
                oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
            }
            oDoc.Application.Selection.Tables[1].Rows[1].Select();
            oDoc.Application.Selection.Font.Name = "Times New Roman";
            oDoc.Application.Selection.Font.Size = 9;
            oDoc.Application.Selection.Tables[1].Rows[1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            oDoc.Application.Selection.Tables[1].Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        }
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
        private void MergeEmptyCells(int RowCount, Document oDoc, int column)
        {
            int checkCell = 2;
            for (int i = 2; i <= RowCount - 1; i++)
            {
                if (oDoc.Application.Selection.Tables[1].Cell(i + 1, column).Range.Text.Length < 3)
                {
                    oDoc.Application.Selection.Tables[1].Cell(checkCell, column).Merge(oDoc.Application.Selection.Tables[1].Cell(i + 1, column));
                }
                else { checkCell = i + 1; }
            }
        }
    }
}
