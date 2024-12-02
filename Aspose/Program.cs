using Aspose.Imaging.Xmp.Types.Complex.Dimensions;
using Aspose.Pdf;
using Aspose.Pdf.Drawing;
using Aspose.Pdf.Text;
using Azure;
using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Drawing.Printing;
using static System.Net.Mime.MediaTypeNames;

internal class Program
{
    static void Main(string[] args)
    {

        Document pdfDocument = new Document();
        Page page = pdfDocument.Pages.Add();
        page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
        page.PageInfo.Margin.Left = 10;
        page.PageInfo.Margin.Right = 10;
        page.PageInfo.Margin.Top = 10;
        page.PageInfo.Margin.Bottom = 10;

        Page1(page);

        Page2(page);



        pdfDocument.Save(@"D:\Aspose_PdfAspose_PDF.pdf");
    }

    static void Page1(Page page)
    {
        PageHeaderContent(page);

        Table TableOfContext1 = new Table
        {
            ColumnWidths = "520",
            Margin = new MarginInfo(15, 1, 10, 2),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };

        // Add the table to the page
        page.Paragraphs.Add(TableOfContext1);

        // Define the header text with different styles
        string textPart1 = "The Limeneal Wheel® ";
        string textPart2 = "Model provides insight into individual's Power, Push and Pain dimensions\n\n";
        string textPart3 = "assessment tool, identifies individual's Power, Push and Pain dimensions through which an individual interacts with others, makes decisions or takes actions.";

        // Create a new row for the header
        Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        // Create a new cell
        Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText = new TextFragment();

        // Add the first part of the text with specific style
        TextSegment segment1 = new TextSegment(textPart1)
        {
            
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial"),
            }
        };
        styledText.Segments.Add(segment1);

        // Add the second part of the text with different style
        TextSegment segment2 = new TextSegment(textPart2)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Bold | FontStyles.Italic,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText.Segments.Add(segment2);



        TextSegment segment3 = new TextSegment(textPart1)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial")
            }
        };

        // Add the second part of the text with different style
        TextSegment segment4 = new TextSegment(textPart3)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };

        styledText.Segments.Add(segment3);
        styledText.Segments.Add(segment4);


        // Add the TextFragment to the cell
        headerCell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell.Margin = new MarginInfo(5, 5, 0, 5);

        PageWithSplitContent(page);

    }

    static void Page2(Page page)
    {
        PageHeaderContent(page);

        Table table = new Table()
        {
            ColumnWidths = "65",
            Margin = new MarginInfo(15, 0, 0, 0),
            BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#315496")),
            DefaultCellBorder = new BorderInfo(BorderSide.All, 1.0f, Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#315496"))),
        };
        page.Paragraphs.Add(table);

        // Create a new row for the header
        Row TableHeaderRow = table.Rows.Add();

        // Create a new cell
        Cell headerCell = TableHeaderRow.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText = new TextFragment();

        TextSegment segment = new TextSegment("VIRTUES")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#FFF")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial")
            }
        };

        styledText.Segments.Add(segment);

        // Add the TextFragment to the cell
        headerCell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#315496"));
        headerCell.Margin = new MarginInfo(3, 2, 5, 2);

        Table virtuesTable = new Table()
        {
            ColumnWidths = "350 200",
            Margin = new MarginInfo(15, 0, 0, 0),
            DefaultCellBorder = new BorderInfo(BorderSide.None, 0.1f),
        };


        Table virtuesTable1 = new Table()
        {
            ColumnWidths = "100 250",
            Margin = new MarginInfo(0, 0, 0, 0),
            DefaultCellBorder = new BorderInfo(BorderSide.None, 0.5f),
        };


        string[][] rowData = new string[][]
        {
            new string[] { "charisma", "persuade and attract others attention"},
            new string[] { "empathy", "sensitive to the needs and emotions of others"},
            new string[] { "Creativeness", "providing creative solutions"},

            new string[] { "understanding", "apply logic and reasoning to know cause and effect"},
            new string[] { "thankfulness", "acknowldges efforts and deeds of others"},
            new string[] { "orderliness", "precise and systematic"},

            new string[] { "Cooperation", "collaborating with teams and all stakeholders" },
            new string[] { "wisdom", "sharing perspectives and applying knowledge" },
            new string[] { "Ambition", "driving for results and targets" },

            new string[] { "friendliness", "friendly and pleasant disposition" },
            new string[] { "discernment", "distinguishing and choosing between options" },
            new string[] { "knowledge", "learn and gather information" },

            new string[] { "humility", "being modest and low-profile" },
            new string[] { "analytical", "evaluates options, weigh pros and cons to decide" },
            new string[] { "Justice", "unbiased, fair and sincere" },

            new string[] { "Kindness", "eagerly helps others" },
            new string[] { "hope", "Finding solutions and possibilities" },
            new string[] { "peace", "calm and unruffled in stressful situations" },

            new string[] { "Solertia", "developing strategies" },
            new string[] { "faith", "guided by instincts" },
            new string[] { "counsel", "advising others" },

            new string[] { "innovativeness", "original and innovative" },
            new string[] { "Trustworthiness", "reliable and committed" },
            new string[] { "meticulous", "detail oriented" },

            new string[] { "prudence", "thinking and acting to achieve the desired outcome" },
            new string[] { "love", "prioritizing to serve and care for others" },
            new string[] { "foresight", "foreseeing challenges" },

            new string[] { "vigilant", "alert and cautious" },
            new string[] { "forgiveness", "forgives others mistakes and looks forward" },
            new string[] { "Flexibility", "adapting to situations, people and environment" },

            new string[] { "compassion", "addresses the pains of others" },
            new string[] { "enthusiasm", "acts with zeal and eagerness" },
            new string[] { "temperance", "emotionally balanced" },

            new string[] { "patience", "sustaining despite provocation and challenges" },
            new string[] { "perseverance", "pursuing despite challenges and risks" },
            new string[] { "fortitude", "courageously dealing with challenges" },
        };

        // Populate the table
        foreach (string[] row in rowData)
        {
            Row tableRow = virtuesTable1.Rows.Add(); // Add a new row to the table

            foreach (string cellData in row)
            {
                Cell cell = tableRow.Cells.Add();
                
                TextFragment text = new TextFragment();

                TextSegment segme = new TextSegment(cellData)
                {
                    TextState = new TextState
                    {
                        ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                        FontSize = 9,
                        FontStyle = FontStyles.Regular,
                        Font = FontRepository.FindFont("Arial")
                    }
                };
                text.Segments.Add(segme);

                cell.Paragraphs.Add(text);

                cell.Margin = new MarginInfo(2, 2, 0, 2);

                cell.Border = new BorderInfo(BorderSide.None, 0.5f);
            }
        }

        page.Paragraphs.Add(virtuesTable);

        Row virtuesTableRow = virtuesTable.Rows.Add();

        Cell virtuesTableCell1 = virtuesTableRow.Cells.Add();

        virtuesTableCell1.Paragraphs.Add(virtuesTable1);

        Cell virtuesTableCell2 = virtuesTableRow.Cells.Add();





        Table virtuesTableBox1 = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(0, 0, 0, 0),
            Margin = new MarginInfo(0, 10, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        Row rowBox1 = virtuesTableBox1.Rows.Add();

        Cell cellBox1 = rowBox1.Cells.Add();
        cellBox1.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        TextFragment styledTextBox1 = new TextFragment();

        TextSegment segmentBox1 = new TextSegment("STRENGTHS")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 24, 
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledTextBox1.TextState.HorizontalAlignment = HorizontalAlignment.Center;

        styledTextBox1.Segments.Add(segmentBox1);

        cellBox1.Paragraphs.Add(styledTextBox1);
        cellBox1.Margin = new MarginInfo(5, 60, 5, 60);





        Table virtuesTableBox2 = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(0, 0, 0, 0),
            Margin = new MarginInfo(0, 10, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // Add a row to the virtues table
        Row rowBox2 = virtuesTableBox2.Rows.Add();

        // Create a cell with a background color and make it square (cube-shaped)
        Cell cellBox2 = rowBox2.Cells.Add();
        cellBox2.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));


        // Create the TextFragment with centered text
        TextFragment styledTextBox2 = new TextFragment();

        // Add the text with specific style
        TextSegment segmentBox2 = new TextSegment("CHALLENGES")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 24,
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledTextBox2.TextState.HorizontalAlignment = HorizontalAlignment.Center;

        styledTextBox2.Segments.Add(segmentBox2);

        // Add the TextFragment to the cell
        cellBox2.Paragraphs.Add(styledTextBox2);
        cellBox2.Margin = new MarginInfo(5, 60, 5, 60);




        Table virtuesTableBox3 = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(0, 0, 0, 0),
            Margin = new MarginInfo(0, 0, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // Add a row to the virtues table
        Row rowBox3 = virtuesTableBox3.Rows.Add();

        // Create a cell with a background color and make it square (cube-shaped)
        Cell cellBox3 = rowBox3.Cells.Add();
        cellBox3.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));


        // Create the TextFragment with centered text
        TextFragment styledTextBox3 = new TextFragment();

        // Add the text with specific style
        TextSegment segmentBox3 = new TextSegment("BARRIERS")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 24,
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledTextBox3.TextState.HorizontalAlignment = HorizontalAlignment.Center;

        styledTextBox3.Segments.Add(segmentBox3);

        // Add the TextFragment to the cell
        cellBox3.Paragraphs.Add(styledTextBox3);
        cellBox3.Margin = new MarginInfo(5, 60, 5, 60);


        virtuesTableCell2.Paragraphs.Add(virtuesTableBox1);
        virtuesTableCell2.Paragraphs.Add(virtuesTableBox2);
        virtuesTableCell2.Paragraphs.Add(virtuesTableBox3);

        EndOfThePageBox(page);

        

    }

    static void PageHeaderContent(Page page)
    {
        Table TableOfContext = new Table
        {
            ColumnWidths = "550",
            Margin = new MarginInfo(10, 1, 10, 10),
            DefaultCellBorder = new BorderInfo(BorderSide.Bottom, 1.0f, Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8"))),
        };

        page.Paragraphs.Add(TableOfContext);

        string[] TableOfContextHeaderTexts = { "LIMENEAL TALENT POTENTIAL for Mr. John Smith" };

        Row TableOfContexHeaderRow = TableOfContext.Rows.Add();
        foreach (string headerText in TableOfContextHeaderTexts)
        {
            TableOfContexHeaderRow.Cells.Add(headerText);
        }

        foreach (Cell cell in TableOfContexHeaderRow.Cells)
        {
            cell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
            cell.DefaultCellTextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 14,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial")
            };
            cell.Paragraphs[0].HorizontalAlignment = HorizontalAlignment.Center;
            cell.Paragraphs[0].Margin = new MarginInfo(0, 3, 0, 3);
        }

    }

    
    static void PageWithSplitContent(Page page)
    {
        Table layoutTable = new Table
        {
            ColumnWidths = "180 380",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            Margin = new MarginInfo(0, 0, 0, 0)
        };

        page.Paragraphs.Add(layoutTable);

        Row row = layoutTable.Rows.Add();

        // First column: Add graph content
        Cell graphCell = row.Cells.Add();


        //Graph graph = new Graph(100.0, 100.0);

        //// Add a circle to the graph as an example
        //Circle circle = new Circle(100, 100, 50);
        //circle.GraphInfo = new GraphInfo
        //{
        //    LineWidth = 2,
        //    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
        //    FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"))
        //};
        //graph.Shapes.Add(circle);

        // Add the graph to the cell
        //graphCell.Paragraphs.Add();



        //Second Column: Add Text content
        Cell textCell = row.Cells.Add();
        string textPart1 = "Power Dimensions ";
        string textPart2 = "represent high impact, most preferred, frequently used\b and naturally expressed dimension by an individual.";

        string textPart3 = "Push Dimensions ";
        string textPart4 = "represent medium impact, sometimes preferred\b dimensions which come with an extra or deliberate effort by the individual.";

        string textPart5 = "Pain Dimensions ";
        string textPart6 = "represent least impact, rarely preferred dimensions which\b are generally stressful and uncomfortable for an individual to express.";

        PageWithSplitContentRightSideContentTable(textCell, textPart1, textPart2);


        Table TableOfContext2 = new Table
        {
            ColumnWidths = "100 260",
            Margin = new MarginInfo(9, 15, 5, 2),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        textCell.Paragraphs.Add(TableOfContext2);

        // Define the header text with different styles
        string textPart21 = "Mentor\nBinder\nPrincipal";
        string textPart22 = "Most inclined to Guiding\nMost inclined to Bonding\nMost inclined to Research";

        // Create a new row for the header
        Row TableOfContexHeaderRow21 = TableOfContext2.Rows.Add();

        // Create a new cell
        Cell headerCell21 = TableOfContexHeaderRow21.Cells.Add();
        Cell headerCell22 = TableOfContexHeaderRow21.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText21 = new TextFragment();
        TextFragment styledText22 = new TextFragment();

        // Add the first part of the text with specific style
        TextSegment segment21 = new TextSegment(textPart21)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText21.Segments.Add(segment21);

        // Add the second part of the text with different style
        TextSegment segment22 = new TextSegment(textPart22)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText22.Segments.Add(segment22);

        // Add the TextFragment to the cell
        headerCell21.Paragraphs.Add(styledText21);
        headerCell22.Paragraphs.Add(styledText22);

        // Style the cell itself (background, padding, etc.)
        //headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell21.Margin = new MarginInfo(2, 10, 0, 0);
        headerCell22.Margin = new MarginInfo(5, 10, 0, 0);

        PageWithSplitContentRightSideContentTable(textCell, textPart3, textPart4);



        Table TableOfContext3 = new Table
        {
            ColumnWidths = "100 260",
            Margin = new MarginInfo(9, 5, 5, 15),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        textCell.Paragraphs.Add(TableOfContext3);

        // Define the header text with different styles
        string textPart31 = "Charmer\nGuardian\nDominion";
        string textPart32 = "Moderately inclined to Influencing\nModerately inclined to Discipline\nModerately inclined to Achieving in risks";

        // Create a new row for the header
        Row TableOfContexHeaderRow31 = TableOfContext3.Rows.Add();

        // Create a new cell
        Cell headerCell31 = TableOfContexHeaderRow31.Cells.Add();
        Cell headerCell32 = TableOfContexHeaderRow31.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText31 = new TextFragment();
        TextFragment styledText32 = new TextFragment();

        // Add the first part of the text with specific style
        TextSegment segment31 = new TextSegment(textPart31)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText31.Segments.Add(segment31);

        // Add the second part of the text with different style
        TextSegment segment32 = new TextSegment(textPart32)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText32.Segments.Add(segment32);

        // Add the TextFragment to the cell
        headerCell31.Paragraphs.Add(styledText31);
        headerCell32.Paragraphs.Add(styledText32);

        // Style the cell itself (background, padding, etc.)
        //headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell31.Margin = new MarginInfo(2, 10, 0, 0);
        headerCell32.Margin = new MarginInfo(5, 10, 0, 0);

        PageWithSplitContentRightSideContentTable(textCell, textPart5, textPart6);


        Table TableOfContext4 = new Table
        {
            ColumnWidths = "100 260",
            Margin = new MarginInfo(9, 5, 5, 15),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        textCell.Paragraphs.Add(TableOfContext4);

        // Define the header text with different styles
        string textPart41 = "Harmonizer\nVisualizer\nAngel";
        string textPart42 = "Less inclined to Harmony\nLess inclined to Value creating\nModerately inclined to Achieving in risks";

        // Create a new row for the header
        Row TableOfContexHeaderRow41 = TableOfContext4.Rows.Add();

        // Create a new cell
        Cell headerCell41 = TableOfContexHeaderRow41.Cells.Add();
        Cell headerCell42 = TableOfContexHeaderRow41.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText41 = new TextFragment();
        TextFragment styledText42 = new TextFragment();

        // Add the first part of the text with specific style
        TextSegment segment41 = new TextSegment(textPart41)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText41.Segments.Add(segment41);

        // Add the second part of the text with different style
        TextSegment segment42 = new TextSegment(textPart42)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText42.Segments.Add(segment42);

        // Add the TextFragment to the cell
        headerCell41.Paragraphs.Add(styledText41);
        headerCell42.Paragraphs.Add(styledText42);

        // Style the cell itself (background, padding, etc.)
        //headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell41.Margin = new MarginInfo(2, 10, 0, 0);
        headerCell42.Margin = new MarginInfo(5, 10, 0, 0);



    }

    static void PageWithSplitContentRightSideContentTable(Cell textCell, string textPart1, string textPart2)
    {
        Table TableOfContext1 = new Table
        {
            ColumnWidths = "360",
            Margin = new MarginInfo(5, 2, 5, 5),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        textCell.Paragraphs.Add(TableOfContext1);

        // Create a new row for the header
        Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        // Create a new cell
        Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText = new TextFragment();

        // Add the first part of the text with specific style
        TextSegment segment1 = new TextSegment(textPart1)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText.Segments.Add(segment1);

        // Add the second part of the text with different style
        TextSegment segment2 = new TextSegment(textPart2)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText.Segments.Add(segment2);

        // Add the TextFragment to the cell
        headerCell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell.Margin = new MarginInfo(5, 5, 0, 5);

    }

    static void EndOfThePageBox(Page page)
    {
        Table TableOfContext1 = new Table
        {
            ColumnWidths = "530",
            Margin = new MarginInfo(15, 2, 5, 5),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        page.Paragraphs.Add(TableOfContext1);

        // Create a new row for the header
        Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        // Create a new cell
        Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

        // Create a TextFragment to hold the styled text
        TextFragment styledText = new TextFragment();


        styledText.Segments.Add(SegmentBoldText("Power Virtues"));
        styledText.Segments.Add(SegmentRegularText("are highly expressed, come naturally, and effortlessly (Top 12)\n"));
        styledText.Segments.Add(SegmentBoldText("Push Virtues "));
        styledText.Segments.Add(SegmentRegularText("are moderately expressed, do not come naturally, and need deliberate effort (Mid 12)\n"));
        styledText.Segments.Add(SegmentBoldText("Pain Virtues "));
        styledText.Segments.Add(SegmentRegularText("are less expressed, puts an individual under stress, and needs extraordinary effort (Low 12)\n"));

        // Add the TextFragment to the cell
        headerCell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell.Margin = new MarginInfo(5, 5, 0, 5);

    }

    static TextSegment SegmentBoldText(string text)
    {
        TextSegment segment1 = new TextSegment(text)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };

        return segment1;
    }

    static TextSegment SegmentRegularText(string text)
    {
        TextSegment segment1 = new TextSegment(text)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Regular,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial")
            }
        };

        return segment1;
    }

    static void AddVirtuesBlockToRow(Cell targetCell, string text)
    {
        // Create a table for the virtues block
        Table virtuesTable = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(10, 10, 10, 10),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        // Add a row to the virtues table
        Row row = virtuesTable.Rows.Add();

        // Create a cell with a background color
        Cell cell = row.Cells.Add();
        cell.BackgroundColor = Color.FromRgb(System.Drawing.Color.LightBlue);

        TextFragment styledText = new TextFragment();

        // Add the text with specific style
        TextSegment segment = new TextSegment(text)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 24,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial")
            }
        };
        styledText.Segments.Add(segment);

        // Add the TextFragment to the cell
        cell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        cell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        cell.Margin = new MarginInfo(15, 15, 15, 15);

        // Add the virtues table as a paragraph inside the target cell
        targetCell.Paragraphs.Add(virtuesTable);
    }


}
