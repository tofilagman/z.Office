using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace z.Office.Microsoft
{
    public class ExcelWriter : IDisposable
    {

        IWorkbook hssworkbook;
        public string FileName { get; set; }
        public bool IsNewFormat { get; private set; }
        Dictionary<string, ISheet> Sheets;
        private Dictionary<string, int> RowIndex;
        public bool DeleteFileOnDisposed { private get; set; } = false;
        public Dictionary<string, HSSFCellStyle> HFontStyle;
        public Dictionary<string, XSSFCellStyle> XFontStyle;


        public ExcelWriter(bool NewFormat)
        {
            this.IsNewFormat = NewFormat;
            this.Sheets = new Dictionary<string, ISheet>();
            this.RowIndex = new Dictionary<string, int>();
            HFontStyle = new Dictionary<string, HSSFCellStyle>();
            XFontStyle = new Dictionary<string, XSSFCellStyle>();

            if (IsNewFormat)
            {
                this.hssworkbook = new XSSFWorkbook();
            }
            else
            {
                this.hssworkbook = new HSSFWorkbook();
            }
        }

        public ExcelWriter(string FileName) : this(Excel.IsFileInNewFormat(FileName))
        {
            this.FileName = FileName;
            using (FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read))
            {
                if (IsNewFormat)
                {
                    this.hssworkbook = new XSSFWorkbook(fs);
                }
                else
                {
                    this.hssworkbook = new HSSFWorkbook(fs);
                }
            }
        }

        public ExcelWriter(Stream fs, bool NewFormat) : this(NewFormat)
        {
            if (IsNewFormat)
            {
                this.hssworkbook = new XSSFWorkbook(fs);
            }
            else
            {
                this.hssworkbook = new HSSFWorkbook(fs);
            }
        }



        public void ClearBook()
        {
            this.hssworkbook.Clear();
        }

        public void AddSheet(string SheetName)
        {
            Sheets.Add(SheetName, this.hssworkbook.CreateSheet(SheetName));
        }

        public IRow AddRow(string SheetName)
        {
            int idx = 0;
            if (!RowIndex.ContainsKey(SheetName))
            {
                RowIndex.Add(SheetName, idx);
            }
            else
            {
                idx = RowIndex[SheetName] + 1;
                RowIndex[SheetName] = idx;
            }

            IRow row = this.hssworkbook.GetSheet(SheetName).CreateRow(idx);

            return row;
        }

        public void AddCell(IRow row, int index, dynamic value, string Style = "")
        {
            row.CreateCell(index).SetCellValue(value);
            if (Style.Trim() != "")
            {
                if (this.IsNewFormat)
                {
                    if (this.XFontStyle.ContainsKey(Style))
                    {
                        row.Cells[index].CellStyle = this.XFontStyle[Style];
                    }
                }
                else if (this.HFontStyle.ContainsKey(Style))
                {
                    row.Cells[index].CellStyle = this.HFontStyle[Style];
                }
            }
        }

        public void AddCellStyle(String StyleName, String FontName = "Arial", Int16 FontSize = 8,
           Boolean IsItalic = false, FontUnderlineType UnderlineType = FontUnderlineType.None,
           FontBoldWeight BoldWeight = FontBoldWeight.None, HorizontalAlignment HorizontalAlign = HorizontalAlignment.Left,
           VerticalAlignment VerticalAlign = VerticalAlignment.Top, BorderStyle TopBorder = BorderStyle.None,
           BorderStyle BottomBorder = BorderStyle.None, BorderStyle RightBorder = BorderStyle.None,
           BorderStyle LeftBorder = BorderStyle.None, IndexedColors FontColor = null,
           IndexedColors BackgroundColor = null)
        {
            IFont font = this.hssworkbook.CreateFont();
            font.Color = ((FontColor == null) ? IndexedColors.Black.Index : FontColor.Index);
            font.FontName = FontName;
            font.FontHeightInPoints = FontSize;
            font.IsItalic = IsItalic;
            if (font.Underline != FontUnderlineType.None)
            {
                font.Underline = UnderlineType;
            }
            font.Boldweight = (short)BoldWeight;

            if (this.IsNewFormat)
            {
                XSSFCellStyle style = (XSSFCellStyle)this.hssworkbook.CreateCellStyle();
                style.SetFont(font);
                style.Alignment = HorizontalAlign;
                style.VerticalAlignment = VerticalAlign;
                style.BorderTop = TopBorder;
                style.BorderBottom = BottomBorder;
                style.BorderRight = RightBorder;
                style.BorderLeft = LeftBorder;

                if (BackgroundColor != null)
                {
                    style.FillForegroundColor = BackgroundColor.Index;
                    style.FillPattern = FillPattern.SolidForeground;
                }

                if (!this.XFontStyle.ContainsKey(StyleName))
                {
                    this.XFontStyle.Add(StyleName, style);
                }
                else
                {
                    this.XFontStyle[StyleName] = style;
                }
            }
            else
            {
                HSSFCellStyle style2 = (HSSFCellStyle)this.hssworkbook.CreateCellStyle();
                style2.SetFont(font);
                style2.Alignment = HorizontalAlign;
                style2.VerticalAlignment = VerticalAlign;
                style2.BorderTop = TopBorder;
                style2.BorderBottom = BottomBorder;
                style2.BorderRight = RightBorder;
                style2.BorderLeft = LeftBorder;

                if (BackgroundColor != null)
                {
                    style2.FillForegroundColor = BackgroundColor.Index;
                    style2.FillPattern = FillPattern.SolidForeground;
                }
                if (!this.HFontStyle.ContainsKey(StyleName))
                {
                    this.HFontStyle.Add(StyleName, style2);
                }
                else
                {
                    this.HFontStyle[StyleName] = style2;
                }
            }
        }

        public void Save()
        {
            foreach (KeyValuePair<string, ISheet> pair in this.Sheets)
            {
                int num = pair.Value.LastRowNum + 1;
                int num2 = 0;
                for (int i = 0; i < num; i++)
                {
                    int physicalNumberOfCells = pair.Value.GetRow(i).PhysicalNumberOfCells;
                    if (physicalNumberOfCells > num2)
                    {
                        num2 = physicalNumberOfCells;
                    }
                }
                for (int j = 0; j < num2; j++)
                {
                    pair.Value.AutoSizeColumn(j);
                }
            }

            using (MemoryStream ms = new MemoryStream())
            {

                if (this.IsNewFormat)
                {
                    ((XSSFWorkbook)this.hssworkbook).Write(ms);
                }
                else
                {
                    this.hssworkbook.Write(ms);
                }

                using (FileStream file = new FileStream(this.FileName, FileMode.Create, FileAccess.Write))
                {
                    ms.WriteTo(file);
                }
            }
        }

        public void SaveToStream(Stream ms)
        {
            foreach (KeyValuePair<string, ISheet> pair in this.Sheets)
            {
                int num = pair.Value.LastRowNum + 1;
                int num2 = 0;
                for (int i = 0; i < num; i++)
                {
                    int physicalNumberOfCells = pair.Value.GetRow(i).PhysicalNumberOfCells;
                    if (physicalNumberOfCells > num2)
                    {
                        num2 = physicalNumberOfCells;
                    }
                }
                for (int j = 0; j < num2; j++)
                {
                    pair.Value.AutoSizeColumn(j);
                }
            }

            //using (MemoryStream ms = new MemoryStream())
            //{

                if (this.IsNewFormat)
                {
                    ((XSSFWorkbook)this.hssworkbook).Write(ms);
                }
                else
                {
                    this.hssworkbook.Write(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                
            //}
        }


        public void Dispose()
        {
            if (this.hssworkbook != null) this.hssworkbook.Dispose();
            if (DeleteFileOnDisposed && File.Exists(this.FileName)) File.Delete(this.FileName);
            GC.Collect();
            GC.SuppressFinalize(this);
        }
    }
}
