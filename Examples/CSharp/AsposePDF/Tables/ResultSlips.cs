using System;

namespace Aspose.Pdf.Examples.CSharp.AsposePDF.Tables
{
    public class ResultSlips
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_AsposePdf_ResultSlips();
            string outputFilename = Guid.NewGuid().ToString();

            Document doc = new Document();
            Page page = doc.Pages.Add();

            var docInfo = doc.Info;
            var pageInfo = page.PageInfo;

            // Create table
            Aspose.Pdf.Table tab1 = new Aspose.Pdf.Table();
            // Add the table in paragraphs collection of the desired section
            page.Paragraphs.Add(tab1);

            // Set with column widths of the table
            tab1.ColumnWidths = "150 150 150";
            tab1.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;

            //tab1.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;
            tab1.DefaultCellTextState = new Pdf.Text.TextState("Arial", 8f);

            // Set default cell border using BorderInfo object
            tab1.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f);

            // Set table border using another customized BorderInfo object
            tab1.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.4f);
            // Create MarginInfo object and set its left, bottom, right and top margins
            Aspose.Pdf.MarginInfo margin = new Aspose.Pdf.MarginInfo();
            margin.Top = 5f;
            margin.Left = 5f;
            margin.Right = 5f;
            margin.Bottom = 5f;

            // Set the default cell padding to the MarginInfo object
            tab1.DefaultCellPadding = margin;

            // Learner name and Learner ULN title
            Aspose.Pdf.Row row1 = tab1.Rows.Add();
            row1.BackgroundColor = Color.WhiteSmoke;
            row1.Cells.Add("Learner Name").ColSpan = 2;
            row1.Cells.Add("Learner ULN");

            // Learner name and Learner ULN value
            Aspose.Pdf.Row row2 = tab1.Rows.Add();
            row2.Cells.Add("Sajid Malik").ColSpan = 2;
            row2.Cells.Add("1234567890");

            // Provider name and UKPRN title
            Aspose.Pdf.Row row3 = tab1.Rows.Add();
            row3.BackgroundColor = Color.WhiteSmoke;
            row3.Cells.Add("Provier Name").ColSpan = 2;
            row3.Cells.Add("Provider UKPRN");

            // Provider name and UKPRN value
            Aspose.Pdf.Row row4 = tab1.Rows.Add();
            row4.Cells.Add("Abingdon and Witney College").ColSpan = 2;
            row4.Cells.Add("5000055");

            // Tlevel title
            Aspose.Pdf.Row row5 = tab1.Rows.Add();
            row5.BackgroundColor = Color.WhiteSmoke;
            row5.Cells.Add("T-Level").ColSpan = 3;

            // Tlevel value
            Aspose.Pdf.Row row6 = tab1.Rows.Add();
            row6.Cells.Add("T Level in Management and Administration").ColSpan = 3;

            // Core component name title
            Aspose.Pdf.Row row7 = tab1.Rows.Add();
            row7.BackgroundColor = Color.WhiteSmoke;
            row7.Cells.Add("Core Component Name").ColSpan = 3;

            // Core component name value
            Aspose.Pdf.Row row8 = tab1.Rows.Add();
            row8.Cells.Add("Maintenance, Installation and Repair for Engineering and Manufacturing").ColSpan = 3;

            // Core component code, exam period and result title
            Aspose.Pdf.Row row9 = tab1.Rows.Add();
            row9.BackgroundColor = Color.WhiteSmoke;
            row9.Cells.Add("Core Component Code");
            row9.Cells.Add("Core Component Exam Period");
            row9.Cells.Add("Core Component Result");

            // Core component code, exam period and result value
            Aspose.Pdf.Row row10 = tab1.Rows.Add();
            row10.Cells.Add("61001115");
            row10.Cells.Add("Autumn 2023");
            row10.Cells.Add("A*");

            // Occupational specialism name title
            Aspose.Pdf.Row row11 = tab1.Rows.Add();
            row11.BackgroundColor = Color.WhiteSmoke;
            row11.Cells.Add("Occupational Specialism(s)").ColSpan = 3;

            // Occupational sp12cialism name value
            Aspose.Pdf.Row row12 = tab1.Rows.Add();
            row12.Cells.Add("Supporting Healthcare - Supporting the Care of Children and Young People").ColSpan = 3;

            // Occupational specialism code, exam period and result title
            Aspose.Pdf.Row row13 = tab1.Rows.Add();
            row13.BackgroundColor = Color.WhiteSmoke;
            row13.Cells.Add("Specialism Code");
            row13.Cells.Add("Specialism Exam Period");
            row13.Cells.Add("Specialism Result");

            // Occupational specialism code, exam period and result title
            Aspose.Pdf.Row row14 = tab1.Rows.Add();
            row14.Cells.Add("ZTLOS039");
            row14.Cells.Add("Summer 2024");
            row14.Cells.Add("Distinction");

            // Occupational specialism name value
            Aspose.Pdf.Row row15 = tab1.Rows.Add();
            row15.Cells.Add("Maintenance Engineering Technologies: Control & Instrumentation").ColSpan = 3;

            // Occupational specialism code, exam period and result title
            Aspose.Pdf.Row row16 = tab1.Rows.Add();
            row16.BackgroundColor = Color.WhiteSmoke;
            row16.Cells.Add("Specialism Code");
            row16.Cells.Add("Specialism Exam Period");
            row16.Cells.Add("Specialism Result");

            // Occupational specialism code, exam period and result title
            Aspose.Pdf.Row row17 = tab1.Rows.Add();
            row17.Cells.Add("ZTLOS039");
            row17.Cells.Add("Summer 2024");
            row17.Cells.Add("Merit");

            // Industry placement status and overall result title
            Aspose.Pdf.Row row18 = tab1.Rows.Add();
            row18.BackgroundColor = Color.WhiteSmoke;
            row18.Cells.Add("Industry Placement Status").ColSpan = 2;
            row18.Cells.Add("Overall Result");

            // Industry placement status and overall result value
            Aspose.Pdf.Row row19 = tab1.Rows.Add();
            row19.Cells.Add("Completed").ColSpan = 2;
            row19.Cells.Add("Distinction");

            var path = dataDir + $"{outputFilename}.pdf";
            doc.Save(path);
            Console.WriteLine("\nFile saved at " + path);
        }
    }
}
