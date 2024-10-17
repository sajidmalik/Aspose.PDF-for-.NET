using Aspose.Pdf.Text;
using System;
using System.Collections.Generic;

namespace Aspose.Pdf.Examples.CSharp.AsposePDF.Tables
{
    public class ResultSlips
    {
        private Document Document;
        private readonly IList<Learner> Learners;

        public ResultSlips()
        {
            Document = new Document();

            Learners = new List<Learner>()
            {
                new Learner() { Name = "Russel Crowe", Uln = 4444444444 },
                new Learner() { Name = "Ricky Ponting", Uln = 5555555555 },
                new Learner() { Name = "Sylvester Stallone", Uln = 3333333333 },
                new Learner() { Name = "Imran Khan", Uln = 1111111111 }
            };
        }

        public void GenerateResultSlips()
        {
            string dataDir = RunExamples.GetDataDir_AsposePdf_ResultSlips();
            string outputFilename = Guid.NewGuid().ToString();

            for (var i = 0; i < Learners.Count; i++)
            {
                var page = Document.Pages.Add();
                page.PageInfo.IsLandscape = true;

                Aspose.Pdf.Rectangle r = page.MediaBox;
                double newHeight = r.Width;
                double newWidth = r.Height;
                double newLLX = r.LLX;

                double newLLY = r.LLY + (r.Height - newHeight);
                page.MediaBox = new Aspose.Pdf.Rectangle(newLLX, newLLY, newLLX + newWidth, newLLY + newHeight);
                page.CropBox = new Aspose.Pdf.Rectangle(newLLX, newLLY, newLLX + newWidth, newLLY + newHeight);

                Table table = new Table();
                SetTableProperties(table);

                page.Paragraphs.Add(GetTable(Learners[i], table));
            }

            var path = dataDir + $"{outputFilename}.pdf";
            Document.Save(path);
            Console.WriteLine("\nFile saved at " + path);
        }

        private Table GetTable(Learner learner, Table table)
        {
            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Learner Name", ColSpan = 2 } },
                { new RowProperty() { Name="Learner ULN" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name=learner.Name, ColSpan = 2 } },
                { new RowProperty() { Name=learner.Uln.ToString() } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Provier Name", ColSpan = 2 } },
                { new RowProperty() { Name="Provider UKPRN" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Abingdon and Witney College", ColSpan = 2 } },
                { new RowProperty() { Name="15000055" } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="T-Level", ColSpan = 3 } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="T Level in Management and Administration", ColSpan = 3 } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Core Component Name", ColSpan = 3 } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Maintenance, Installation and Repair for Engineering and Manufacturing", ColSpan = 3 } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Core Component Code" } },
                { new RowProperty() { Name="Core Component Exam Period" } },
                { new RowProperty() { Name="Core Component Result" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="61001115" } },
                { new RowProperty() { Name="Autumn 2023" } },
                { new RowProperty() { Name="A*" } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Occupational Specialism(s)", ColSpan = 3 } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Supporting Healthcare - Supporting the Care of Children and Young People", ColSpan = 3 } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Specialism Code" } },
                { new RowProperty() { Name="Specialism Exam Period" } },
                { new RowProperty() { Name="Specialism Result" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="ZTLOS039" } },
                { new RowProperty() { Name="Summer 2023" } },
                { new RowProperty() { Name="Distinction" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Maintenance Engineering Technologies: Control & Instrumentation", ColSpan = 3 } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Specialism Code" } },
                { new RowProperty() { Name="Specialism Exam Period" } },
                { new RowProperty() { Name="Specialism Result" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="ZTLOS039" } },
                { new RowProperty() { Name="Summer 2024" } },
                { new RowProperty() { Name="Merit" } }
            }));

            table.Rows.Add(new HeaderRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Industry Placement Status", ColSpan = 2 } },
                { new RowProperty() { Name="Overall Result" } }
            }));

            table.Rows.Add(new ValueRow().GetRow(new List<RowProperty>() {
                { new RowProperty() { Name="Completed", ColSpan = 2 } },
                { new RowProperty() { Name="Distinction" } }
            }));

            #region static_code

            //// Learner name and Learner ULN title
            //Aspose.Pdf.Row row1 = table.Rows.Add();
            //row1.BackgroundColor = Color.WhiteSmoke;
            //row1.Cells.Add("Learner Name").ColSpan = 2;
            //row1.Cells.Add("Learner ULN");

            //// Learner name and Learner ULN value
            //Aspose.Pdf.Row row2 = table.Rows.Add();
            //row2.Cells.Add(learner.Name).ColSpan = 2;
            //row2.Cells.Add(learner.Uln.ToString());

            //// Provider name and UKPRN title
            //Aspose.Pdf.Row row3 = table.Rows.Add();
            //row3.BackgroundColor = Color.WhiteSmoke;
            //row3.Cells.Add("Provier Name").ColSpan = 2;
            //row3.Cells.Add("Provider UKPRN");

            //// Provider name and UKPRN value
            //Aspose.Pdf.Row row4 = table.Rows.Add();
            //row4.Cells.Add("Abingdon and Witney College").ColSpan = 2;
            //row4.Cells.Add("15000055");

            //// Tlevel title
            //Aspose.Pdf.Row row5 = table.Rows.Add();
            //row5.BackgroundColor = Color.WhiteSmoke;
            //row5.Cells.Add("T-Level").ColSpan = 3;

            //// Tlevel value
            //Aspose.Pdf.Row row6 = table.Rows.Add();
            //row6.Cells.Add("T Level in Management and Administration").ColSpan = 3;

            //// Core component name title
            //Aspose.Pdf.Row row7 = table.Rows.Add();
            //row7.BackgroundColor = Color.WhiteSmoke;
            //row7.Cells.Add("Core Component Name").ColSpan = 3;

            //// Core component name value
            //Aspose.Pdf.Row row8 = table.Rows.Add();
            //row8.Cells.Add("Maintenance, Installation and Repair for Engineering and Manufacturing").ColSpan = 3;

            //// Core component code, exam period and result title
            //Aspose.Pdf.Row row9 = table.Rows.Add();
            //row9.BackgroundColor = Color.WhiteSmoke;
            //row9.Cells.Add("Core Component Code");
            //row9.Cells.Add("Core Component Exam Period");
            //row9.Cells.Add("Core Component Result");

            //// Core component code, exam period and result value
            //Aspose.Pdf.Row row10 = table.Rows.Add();
            //row10.Cells.Add("61001115");
            //row10.Cells.Add("Autumn 2023");
            //row10.Cells.Add("A*");

            //// Occupational specialism name title
            //Aspose.Pdf.Row row11 = table.Rows.Add();
            //row11.BackgroundColor = Color.WhiteSmoke;
            //row11.Cells.Add("Occupational Specialism(s)").ColSpan = 3;

            //// Occupational sp12cialism name value
            //Aspose.Pdf.Row row12 = table.Rows.Add();
            //row12.Cells.Add("Supporting Healthcare - Supporting the Care of Children and Young People").ColSpan = 3;

            //// Occupational specialism code, exam period and result title
            //Aspose.Pdf.Row row13 = table.Rows.Add();
            //row13.BackgroundColor = Color.WhiteSmoke;
            //row13.Cells.Add("Specialism Code");
            //row13.Cells.Add("Specialism Exam Period");
            //row13.Cells.Add("Specialism Result");

            //// Occupational specialism code, exam period and result title
            //Aspose.Pdf.Row row14 = table.Rows.Add();
            //row14.Cells.Add("ZTLOS039");
            //row14.Cells.Add("Summer 2024");
            //row14.Cells.Add("Distinction");

            //// Occupational specialism name value
            //Aspose.Pdf.Row row15 = table.Rows.Add();
            //row15.Cells.Add("Maintenance Engineering Technologies: Control & Instrumentation").ColSpan = 3;

            //// Occupational specialism code, exam period and result title
            //Aspose.Pdf.Row row16 = table.Rows.Add();
            //row16.BackgroundColor = Color.WhiteSmoke;
            //row16.Cells.Add("Specialism Code");
            //row16.Cells.Add("Specialism Exam Period");
            //row16.Cells.Add("Specialism Result");

            //// Occupational specialism code, exam period and result title
            //Aspose.Pdf.Row row17 = table.Rows.Add();
            //row17.Cells.Add("ZTLOS039");
            //row17.Cells.Add("Summer 2024");
            //row17.Cells.Add("Merit");

            //// Industry placement status and overall result title
            //Aspose.Pdf.Row row18 = table.Rows.Add();
            //row18.BackgroundColor = Color.WhiteSmoke;
            //row18.Cells.Add("Industry Placement Status").ColSpan = 2;
            //row18.Cells.Add("Overall Result");

            //// Industry placement status and overall result value
            //Aspose.Pdf.Row row19 = table.Rows.Add();
            //row19.Cells.Add("Completed").ColSpan = 2;
            //row19.Cells.Add("Distinction");

            #endregion static_code

            return table;
        }

        private void SetTableProperties(Table table)
        {
            table.ColumnWidths = "150 150 150";
            table.ColumnAdjustment = ColumnAdjustment.AutoFitToWindow;

            table.DefaultCellTextState = new Pdf.Text.TextState("Arial", 8f);
            table.DefaultCellBorder = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.1f);
            table.Border = new Aspose.Pdf.BorderInfo(Aspose.Pdf.BorderSide.All, 0.6f);

            Aspose.Pdf.MarginInfo margin = new Aspose.Pdf.MarginInfo();
            margin.Top = 5f;
            margin.Left = 5f;
            margin.Right = 5f;
            margin.Bottom = 5f;

            table.DefaultCellPadding = margin;
        }
    }

    public abstract class BaseRow
    {
        protected readonly Row row;
        private Cell cell;
        private TextFragment fragment;

        public abstract TextState RowStyle { get; }

        public BaseRow()
        {
            row = new Row();
        }

        public Row GetRow(IList<RowProperty> rowProperties)
        {
            row.DefaultCellTextState = RowStyle;

            foreach (RowProperty property in rowProperties)
            {
                cell = new Cell();
                fragment = new TextFragment(property.Name);

                if (property.ColSpan > 0)
                    cell.ColSpan = property.ColSpan;

                cell.Paragraphs.Add(fragment);
                row.Cells.Add(cell);
            }
            return row;
        }
    }

    public class RowProperty
    {
        public string Name { get; set; }
        public int ColSpan { get; set; }
    }

    public class ValueRow : BaseRow
    {
        public ValueRow()
        {
            row.BackgroundColor = Color.White;
        }
        public override TextState RowStyle
        {
            get
            {
                TextState textState = new TextState();
                textState.FontSize = 11f;
                return textState;
            }
        }
    }

    public class HeaderRow : BaseRow
    {
        public HeaderRow()
        {
            row.BackgroundColor = Color.LightGray;
        }
        public override TextState RowStyle
        {
            get
            {
                TextState textState = new TextState();
                textState.FontSize = 9f;
                textState.FontStyle = FontStyles.Bold;
                return textState;
            }
        }
    }

    public class Learner
    {
        public long Uln { get; set; }
        public string Name { get; set; }

    }
}
