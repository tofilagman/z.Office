using System;
using System.Globalization;
using System.Xml;
using System.Xml.XPath;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using System.Diagnostics;
using System.Data;
using System.IO;

namespace z.Office.PDF
{
    public class DataTableToPdf : IDisposable
    {
        Document document;
        DataTable dt;
        string title;
        Table table;

        public bool deleteFileOnDispose = false;

        private Orientation pageOrient = Orientation.Landscape;

        string docFile;

        public DataTableToPdf(DataTable mdt, string docTitle)
        {
            this.dt = mdt;
            this.title = docTitle;
        }

        /// <summary>
        ///  Must be PDF
        /// </summary>
        /// <param name="filename"></param>
        public void Save(string filename)
        {
            try
            {

                if (Path.GetExtension(filename).ToLower() != ".pdf")
                {
                    throw new Exception("Please specify pdf file");
                }

                // Create a MigraDoc document
                Document document = this.CreateDocument();
                document.UseCmykColor = true;

                // Create a renderer for PDF that uses Unicode font encoding
                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);

                // Set the MigraDoc document
                pdfRenderer.Document = document;


                // Create the PDF document
                pdfRenderer.RenderDocument();

                // Save the PDF document...
                
                pdfRenderer.Save(filename);

                this.docFile = filename;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region Rock and Roll

        public Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.document = new Document();
            this.document.Info.Title = "";
            this.document.Info.Subject = "";
            this.document.Info.Author = "LJ Gomez";
            this.document.DefaultPageSetup.Orientation = this.pageOrient;
            this.document.DefaultPageSetup.LeftMargin = Unit.FromInch(.3);
            this.document.DefaultPageSetup.RightMargin = Unit.FromInch(.3);

            DefineStyles();

            CreatePage();

            FillContent();

            return this.document;
        }

        /// <summary>
        /// Defines the styles used to format the MigraDoc document.
        /// </summary>
        void DefineStyles()
        {
            // Get the predefined style Normal.
            Style style = this.document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Verdana";

            style = this.document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("0cm", TabAlignment.Right);

            style = this.document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("2cm", TabAlignment.Center);

            // Create a new style called Table based on style Normal
            style = this.document.Styles.AddStyle("Table", "Normal");
            style.Font.Name = "Verdana";
            style.Font.Name = "Times New Roman";
            style.Font.Size = 9;

            // Create a new style called Reference based on style Normal
            style = this.document.Styles.AddStyle("Reference", "Normal");
            style.ParagraphFormat.SpaceBefore = "5mm";
            style.ParagraphFormat.SpaceAfter = "5mm";
            style.ParagraphFormat.TabStops.AddTabStop("2cm", TabAlignment.Right);

        }


        void CreatePage()
        {
            // Each MigraDoc document needs at least one section.
            Section section = this.document.AddSection();
 
            // Create footer
            Paragraph paragraph = section.Footers.Primary.AddParagraph();
            //paragraph.AddText("InSys");
            paragraph.Format.Font.Size = 9;
            paragraph.Format.Alignment = ParagraphAlignment.Center;
 
            // Add the print date field
            paragraph = section.AddParagraph();
            paragraph.Format.SpaceBefore = "1cm";
            paragraph.Style = "Reference";
            paragraph.AddFormattedText(this.title, TextFormat.Bold);
 
            // Create the item table
            this.table = section.AddTable();
            this.table.Style = "Table";
            this.table.Borders.Color = TableBorder;
            this.table.Borders.Width = 0.25;
            this.table.Borders.Left.Width = 0.5;
            this.table.Borders.Right.Width = 0.5;
            this.table.Rows.LeftIndent = 0;

            // Before you can add a row, you must define the columns
            Column column;
            foreach (DataColumn col in dt.Columns)
            {
                column = this.table.AddColumn(); //Unit.FromCentimeter(2)
                column.Format.Alignment = ParagraphAlignment.Center;
            }

            // Create the header of the table
            Row row = table.AddRow();
            row.HeadingFormat = true;
            row.Format.Alignment = ParagraphAlignment.Center;
            row.Format.Font.Bold = true;
            row.Shading.Color = TableBlue;

            for (int i = 0; i < dt.Columns.Count; i++)
            {

                row.Cells[i].AddParagraph(dt.Columns[i].ColumnName);
                row.Cells[i].Format.Font.Bold = false;
                row.Cells[i].Format.Alignment = ParagraphAlignment.Left;
                row.Cells[i].VerticalAlignment = VerticalAlignment.Bottom;

            }


            this.table.SetEdge(0, 0, dt.Columns.Count, 1, Edge.Box, BorderStyle.Single, 0.75, Color.Empty);

        }

        /// <summary>
        /// Creates the dynamic parts of the invoice.
        /// </summary>
        void FillContent()
        {
           
            Row row1;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                row1 = this.table.AddRow();


                row1.TopPadding = 1.5;


                for (int j = 0; j < dt.Columns.Count; j++)
                {

                    row1.Cells[j].Shading.Color = Color.Empty;
                    row1.Cells[j].VerticalAlignment = VerticalAlignment.Center;

                    row1.Cells[j].Format.Alignment = ParagraphAlignment.Left;
                    row1.Cells[j].Format.FirstLineIndent = 1;

                    row1.Cells[j].AddParagraph(dt.Rows[i][j].ToString());


                    this.table.SetEdge(0, this.table.Rows.Count - 2, dt.Columns.Count, 1, Edge.Box, BorderStyle.Single, 0.75);
                }
            }

        }

        // Some pre-defined colors
#if true
        // RGB colors
        readonly static Color TableBorder = new Color(81, 125, 192);
        readonly static Color TableBlue = new Color(235, 240, 249);
        readonly static Color TableGray = new Color(242, 242, 242);
#else
        // CMYK colors
        readonly static Color tableBorder = Color.FromCmyk(100, 50, 0, 30);
        readonly static Color tableBlue = Color.FromCmyk(0, 80, 50, 30);
        readonly static Color tableGray = Color.FromCmyk(30, 0, 0, 0, 100);
#endif

        #endregion

        void IDisposable.Dispose()
        {
            document = null;
            dt.Dispose();
            dt = null;
            table = null;

            if (deleteFileOnDispose)
            {
                if (this.docFile != null)
                {
                    try
                    {
                        if (File.Exists(this.docFile))
                        {
                            File.Delete(this.docFile);
                        }
                    }
                    catch { }
                }
            }

            GC.Collect();
            GC.SuppressFinalize(this);
        }

        public pOrientation pageOrientation
        {
            set
            {
                this.pageOrient = (Orientation)value;
            }
        }

        public enum pOrientation
        {
            Portrait = 0,
            Landscape = 1
        }

    }
}
