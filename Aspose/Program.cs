using Aspose.Pdf;
using Aspose.Pdf.Drawing;
using Aspose.Pdf.Text;
using Rectangle = Aspose.Pdf.Drawing.Rectangle;


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

        page.Paragraphs.Add(TableOfContext1);

        string textPart1 = "The Limeneal Wheel® ";
        string textPart2 = "Model provides insight into individual's Power, Push and Pain dimensions\n\n";
        string textPart3 = "assessment tool, identifies individual's Power, Push and Pain dimensions through which an individual interacts with others, makes decisions or takes actions.";

        Aspose.Pdf.Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        Aspose.Pdf.Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

        TextFragment styledText = new TextFragment();

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


        headerCell.Paragraphs.Add(styledText);

        headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell.Margin = new MarginInfo(5, 5, 0, 5);

        PageWithSplitContent(page);

    }

    static void Page2(Page page)
    {
        PageHeaderContent(page);

        var values = new object[][]
    {
        // Category "Power"
        new object[] { "Power", "Charisma", 87 },
        new object[] { "Power", "Empathy", 85 },
        new object[] { "Power", "Creativeness", 82 },
        new object[] { "Power", "Understanding", 79 },
        new object[] { "Power", "Thankfulness", 79 },
        new object[] { "Power", "Orderliness", 76 },
        new object[] { "Power", "Cooperation", 74 },
        new object[] { "Power", "Wisdom", 70 },
        new object[] { "Power", "Ambition", 69 },
        new object[] { "Power", "Friendliness", 69 },
        new object[] { "Power", "Discernment", 59 },
        new object[] { "Power", "Knowledge", 59 },
        new object[] { "Power", "Humility", 65 },
        new object[] { "Power", "Analytical", 63 },
        new object[] { "Power", "Justice", 61 },
        new object[] { "Power", "Kindness", 58 },
        new object[] { "Power", "Hope", 56 },
        new object[] { "Power", "Peace", 56 },
        new object[] { "Power", "Faith", 54 },
        new object[] { "Power", "Counsel", 51 },

        // Category "Push"
        new object[] { "Push", "Charisma", 63 },
        new object[] { "Push", "Empathy", 58 },
        new object[] { "Push", "Creativeness", 56 },
        new object[] { "Push", "Understanding", 52 },
        new object[] { "Push", "Thankfulness", 51 },
        new object[] { "Push", "Orderliness", 50 },
        new object[] { "Push", "Cooperation", 46 },
        new object[] { "Push", "Wisdom", 45 },
        new object[] { "Push", "Ambition", 41 },
        new object[] { "Push", "Friendliness", 39 },
        new object[] { "Push", "Discernment", 38 },
        new object[] { "Push", "Knowledge", 37 },
        new object[] { "Push", "Humility", 34 },
        new object[] { "Push", "Analytical", 32 },
        new object[] { "Push", "Justice", 30 },
        new object[] { "Push", "Kindness", 28 },
        new object[] { "Push", "Hope", 26 },
        new object[] { "Push", "Peace", 24 },
        new object[] { "Push", "Faith", 22 },
        new object[] { "Push", "Counsel", 20 },

        // Category "Pain"
        new object[] { "Pain", "Charisma", 30 },
        new object[] { "Pain", "Empathy", 26 },
        new object[] { "Pain", "Creativeness", 24 },
        new object[] { "Pain", "Understanding", 22 },
        new object[] { "Pain", "Thankfulness", 20 },
        new object[] { "Pain", "Orderliness", 18 },
        new object[] { "Pain", "Cooperation", 16 },
        new object[] { "Pain", "Wisdom", 14 },
        new object[] { "Pain", "Ambition", 12 },
        new object[] { "Pain", "Friendliness", 10 },
        new object[] { "Pain", "Discernment", 8 },
        new object[] { "Pain", "Knowledge", 6 },
        new object[] { "Pain", "Humility", 4 },
        new object[] { "Pain", "Analytical", 2 },
        new object[] { "Pain", "Justice", 1 },
        new object[] { "Pain", "Kindness", 0 },
        new object[] { "Pain", "Hope", 0 },
        new object[] { "Pain", "Peace", 0 },
        new object[] { "Pain", "Faith", 0 },
        new object[] { "Pain", "Counsel", 0 }
    };

        // Example usage: Passing data to generate the bar graph
        //new Program().MainBarGraph(page, values);

        new Program().VirtuesBarGraph(page, values);



        Table table = new Table()
        {
            ColumnWidths = "65",
            Margin = new MarginInfo(15, 0, 0, 0),
            BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#315496")),
            DefaultCellBorder = new BorderInfo(BorderSide.All, 1.0f, Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#315496"))),
        };
        page.Paragraphs.Add(table);

        Aspose.Pdf.Row TableHeaderRow = table.Rows.Add();

        Aspose.Pdf.Cell headerCell = TableHeaderRow.Cells.Add();

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

        headerCell.Paragraphs.Add(styledText);

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

        foreach (string[] row in rowData)
        {
            Aspose.Pdf.Row tableRow = virtuesTable1.Rows.Add(); 

            foreach (string cellData in row)
            {
                Aspose.Pdf.Cell cell = tableRow.Cells.Add();
                
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

        Aspose.Pdf.Row virtuesTableRow = virtuesTable.Rows.Add();

        Aspose.Pdf.Cell virtuesTableCell1 = virtuesTableRow.Cells.Add();

        virtuesTableCell1.Paragraphs.Add(virtuesTable1);

        Aspose.Pdf.Cell virtuesTableCell2 = virtuesTableRow.Cells.Add();





        Table virtuesTableBox1 = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(0, 0, 0, 0),
            Margin = new MarginInfo(0, 10, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        Aspose.Pdf.Row rowBox1 = virtuesTableBox1.Rows.Add();

        Aspose.Pdf.Cell cellBox1 = rowBox1.Cells.Add();
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

        Aspose.Pdf.Row rowBox2 = virtuesTableBox2.Rows.Add();

        Aspose.Pdf.Cell cellBox2 = rowBox2.Cells.Add();
        cellBox2.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));


        TextFragment styledTextBox2 = new TextFragment();

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

        Aspose.Pdf.Row rowBox3 = virtuesTableBox3.Rows.Add();

        Aspose.Pdf.Cell cellBox3 = rowBox3.Cells.Add();
        cellBox3.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));

        TextFragment styledTextBox3 = new TextFragment();

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

        Aspose.Pdf.Row TableOfContexHeaderRow = TableOfContext.Rows.Add();
        foreach (string headerText in TableOfContextHeaderTexts)
        {
            TableOfContexHeaderRow.Cells.Add(headerText);
        }

        foreach (Aspose.Pdf.Cell cell in TableOfContexHeaderRow.Cells)
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

        Aspose.Pdf.Row row = layoutTable.Rows.Add();

        Aspose.Pdf.Cell graphCell = row.Cells.Add();

        //BarGraph(graphCell);
        //Program p = new Program();
        //p.BarGraph(graphCell, 71.4f, 71.4f, 64.5f, "POWER", "Mentor", "Binder", "Principal");
        //p.BarGraph(graphCell, 63.7f, 62.3f, 55.6f, "PUSH", "Charmer", "Guardian", "Dominion");
        //p.BarGraph(graphCell, 51.6f, 39.9f, 37.7f, "PAIN", "Harmonizer", "Visualizer", "Angel");


        //AddBarGraph(graphCell);
        //AddBarGraph(graphCell);
        //AddBarGraph(graphCell);

        Program p = new Program();
        p.AddBarGraphDimensions(graphCell, "POWER", 71.4f, 71.4f, 67.3f, "Mentor", "Binder", "Principal");
        p.AddBarGraphDimensions(graphCell, "PUSH", 63.7f, 62.3f, 55.6f, "Charmer", "BinderGuardian", "Dominion");
        p.AddBarGraphDimensions(graphCell, "PAIN", 51.6f, 39.9f, 37.7f, "Harmonizer", "Visualizer", "Angel");

        

        //Second Column: Add Text content
        Aspose.Pdf.Cell textCell = row.Cells.Add();
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
        Aspose.Pdf.Row TableOfContexHeaderRow21 = TableOfContext2.Rows.Add();

        // Create a new cell
        Aspose.Pdf.Cell headerCell21 = TableOfContexHeaderRow21.Cells.Add();
        Aspose.Pdf.Cell headerCell22 = TableOfContexHeaderRow21.Cells.Add();

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
        Aspose.Pdf.Row TableOfContexHeaderRow31 = TableOfContext3.Rows.Add();

        // Create a new cell
        Aspose.Pdf.Cell headerCell31 = TableOfContexHeaderRow31.Cells.Add();
        Aspose.Pdf.Cell headerCell32 = TableOfContexHeaderRow31.Cells.Add();

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
        Aspose.Pdf.Row TableOfContexHeaderRow41 = TableOfContext4.Rows.Add();

        // Create a new cell
        Aspose.Pdf.Cell headerCell41 = TableOfContexHeaderRow41.Cells.Add();
        Aspose.Pdf.Cell headerCell42 = TableOfContexHeaderRow41.Cells.Add();

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

    static void PageWithSplitContentRightSideContentTable(Aspose.Pdf.Cell textCell, string textPart1, string textPart2)
    {
        Table TableOfContext1 = new Table
        {
            ColumnWidths = "360",
            Margin = new MarginInfo(5, 2, 5, 5),
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };


        textCell.Paragraphs.Add(TableOfContext1);

        Aspose.Pdf.Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        Aspose.Pdf.Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

        TextFragment styledText = new TextFragment();

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
        Aspose.Pdf.Row TableOfContexHeaderRow1 = TableOfContext1.Rows.Add();

        // Create a new cell
        Aspose.Pdf.Cell headerCell = TableOfContexHeaderRow1.Cells.Add();

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

    static void AddVirtuesBlockToRow(Aspose.Pdf.Cell targetCell, string text)
    {
        Table virtuesTable = new Table
        {
            ColumnWidths = "100%",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
            DefaultCellPadding = new MarginInfo(10, 10, 10, 10),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        Aspose.Pdf.Row row = virtuesTable.Rows.Add();

        Aspose.Pdf.Cell cell = row.Cells.Add();
        cell.BackgroundColor = Color.FromRgb(System.Drawing.Color.LightBlue);

        TextFragment styledText = new TextFragment();

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

        cell.Paragraphs.Add(styledText);

        cell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        cell.Margin = new MarginInfo(15, 15, 15, 15);

        targetCell.Paragraphs.Add(virtuesTable);
    }

    public void AddBarGraphDimensions(Cell maingraph, string title, float p1, float p2, float p3, string p1Name, string p2Name, string p3Name)
    {
        Table table = new Table()
        {
            ColumnWidths = "180"
        };

        float maxHight = Math.Max(p1, Math.Max(p2, p3));

        table.DefaultCellBorder = new BorderInfo(BorderSide.None);

        BorderInfo borderInfo = new BorderInfo
        {
            Top = new GraphInfo
            {
                LineWidth = 0.1f, // Set the top border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#E6EAEB"))
            },
            Left = new GraphInfo
            {
                LineWidth = 0.1f, // Set the left border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#E6EAEB"))
            },
            Right = new GraphInfo
            {
                LineWidth = 3, // Set the right border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#A6A6A8"))
            },
            Bottom = new GraphInfo
            {
                LineWidth = 3, // Set the bottom border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#A6A6A8"))
            }
        };

        table.Border = borderInfo;
        table.Margin = new MarginInfo
        {
            Top = 0,
            Bottom = 10,
            Left = 0,
            Right = 0
        };


        maingraph.Paragraphs.Add(table);

        Row row = table.Rows.Add();
        //row.DefaultCellPadding = new MarginInfo(0, 0, 0, 0);

        Cell headerCell = row.Cells.Add();

        TextFragment styledText = new TextFragment();

        TextSegment segment = new TextSegment(title)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 10,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        styledText.Segments.Add(segment);

        headerCell.Paragraphs.Add(styledText);

        // Style the cell itself (background, padding, etc.)
        //headerCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
        headerCell.Margin = new MarginInfo(50, 0, 0, 0);

        Row row2 = table.Rows.Add();

        Cell graphCell = row2.Cells.Add();


        Table nestedTable = new Table()
        {
            ColumnWidths = "20 80 80",
            DefaultCellBorder = new BorderInfo(BorderSide.None)
        };

        graphCell.Paragraphs.Add(nestedTable);

        Row nesterTabelRow = nestedTable.Rows.Add();

        Cell percentageCell = nesterTabelRow.Cells.Add();

        //float minVal = 100 - maxHight;
        //List<int> numbers = new List<int>();
        //float[] indicate = { (float)Math.Round(minVal, 1) };
        //int current = (int)Math.Ceiling(minVal);

        //int rem = current % 10;
        //current += Math.Abs(rem - 10);

        //while (current <= 100)
        //{
        //    numbers.Add(current);
        //    current += 10; 
        //}


        TextFragment percentageCellStyledText = new TextFragment();

        TextSegment percentageCellSegment = new TextSegment("100%\n80%\n60%\n40%\n20%\n0%")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 6,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        percentageCellStyledText.Segments.Add(percentageCellSegment);

        percentageCell.Paragraphs.Add(percentageCellStyledText);
        percentageCell.Margin = new MarginInfo(4, 0, 0, 0);




        Cell dimensionalCell = nesterTabelRow.Cells.Add();

        Table dimensionalCellTable = new Table()
        {
            ColumnWidths = "80",
            DefaultCellBorder = new BorderInfo(BorderSide.None)
        };

        dimensionalCell.Paragraphs.Add(dimensionalCellTable);

        Row dimensionalCellTableCell1 = dimensionalCellTable.Rows.Add();
        Row dimensionalCellTableCell2 = dimensionalCellTable.Rows.Add();

        Cell dimensionalPercentageCell = dimensionalCellTableCell1.Cells.Add();
        Cell dimensionalPercentageCell2 = dimensionalCellTableCell2.Cells.Add();

        TextFragment dimensionalPercentageCellStyledText = new TextFragment();

        TextSegment dimensionalPercentageCellSegment1 = new TextSegment($"{p1}%\b\b\b")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 5,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        TextSegment dimensionalPercentageCellSegment2 = new TextSegment($"{p2}%\b\b\b")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4473C5")),
                FontSize = 5,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        TextSegment dimensionalPercentageCellSegment3 = new TextSegment($"{p3}%")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3C8E7")),
                FontSize = 5,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        dimensionalPercentageCellStyledText.Segments.Add(dimensionalPercentageCellSegment1);
        dimensionalPercentageCellStyledText.Segments.Add(dimensionalPercentageCellSegment2);
        dimensionalPercentageCellStyledText.Segments.Add(dimensionalPercentageCellSegment3);

        dimensionalPercentageCellStyledText.Margin = new MarginInfo(5, 0, 0, 0);

        dimensionalPercentageCell.Paragraphs.Add(dimensionalPercentageCellStyledText);


        
        Graph graph = new Graph(70.0, 51.0);
         
        Aspose.Pdf.Drawing.Rectangle p1Bar = new Aspose.Pdf.Drawing.Rectangle(0, 0, 20, (float)(p1 * 0.50))
        {
            GraphInfo = new GraphInfo
            {
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#2F5596")),
                FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#2F5596")),
            }
        };
        graph.Shapes.Add(p1Bar);

        Aspose.Pdf.Drawing.Rectangle p2Bar = new Aspose.Pdf.Drawing.Rectangle(25, 0, 20, (float)(p2 * 0.50))
        {
            GraphInfo = new GraphInfo
            {
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4473C5")),
                FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4473C5"))
            }
        };
        graph.Shapes.Add(p2Bar);

        Aspose.Pdf.Drawing.Rectangle p3Bar = new Aspose.Pdf.Drawing.Rectangle(50, 0, 20, (float)(p3 * 0.50))
        {
            GraphInfo = new GraphInfo
            {
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3C8E7")),
                FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3C8E7"))
            }
        };
        graph.Shapes.Add(p3Bar);

        dimensionalPercentageCell2.Paragraphs.Add(graph);





        Cell seriesCell = nesterTabelRow.Cells.Add();

        Table seriesTable = new Table()
        {
            ColumnWidths = "5 60",
            DefaultCellBorder = new BorderInfo(BorderSide.None)
        };

        seriesCell.Paragraphs.Add(seriesTable);

        Row seriesRow1 = seriesTable.Rows.Add();

        Cell seriesColorCell1 = seriesRow1.Cells.Add();
        Cell seriesNameCell1 = seriesRow1.Cells.Add();

        SeriesNameCellMethod(seriesColorCell1, seriesNameCell1, p1Name, "#2F5596");

        Row seriesRow2 = seriesTable.Rows.Add();

        Cell seriesColorCell2 = seriesRow2.Cells.Add();
        Cell seriesNameCell2 = seriesRow2.Cells.Add();

        SeriesNameCellMethod(seriesColorCell2, seriesNameCell2, p2Name, "#4473C5");

        Row seriesRow3 = seriesTable.Rows.Add();

        Cell seriesColorCell3 = seriesRow3.Cells.Add();
        Cell seriesNameCell3 = seriesRow3.Cells.Add();

        SeriesNameCellMethod(seriesColorCell3, seriesNameCell3, p3Name, "#B4C7E8");

        
        Row empltyRow = table.Rows.Add();

        Cell emptyCell = empltyRow.Cells.Add();

        TextFragment emptyCellStyledText = new TextFragment();

        TextSegment emptyCellSegment = new TextSegment("\b")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 3,
                FontStyle = FontStyles.Bold,
                LineSpacing = 5,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        emptyCellStyledText.Segments.Add(emptyCellSegment);

        emptyCell.Paragraphs.Add(emptyCellStyledText);

        maingraph.Margin = new MarginInfo(30, 0, 0, 0);


    }

    public void SeriesNameCellMethod(Cell seriesColorCell,Cell seriesNameCell, string seriesName, string color)
    {
        Graph graphSeries1 = new Graph(3.0, 3.0);

        Rectangle graphSeries1Color1 = new Aspose.Pdf.Drawing.Rectangle(0, 0, 3, 3)
        {
            GraphInfo = new GraphInfo
            {
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(color)),
                FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(color))
            }
        };

        graphSeries1.Shapes.Add(graphSeries1Color1);

        TextFragment seriesNameCellStyledText = new TextFragment();

        TextSegment seriesNameCellSegment = new TextSegment(seriesName)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#776B75")),
                FontSize = 5,
                FontStyle = FontStyles.Regular,
                Font = FontRepository.FindFont("Arial"),
                HorizontalAlignment = HorizontalAlignment.Center
            }
        };


        seriesNameCellStyledText.Margin = new Aspose.Pdf.MarginInfo{ Top = 0, Bottom = 4, Left = 0, Right = 0};

        seriesNameCellStyledText.Segments.Add(seriesNameCellSegment);

        seriesColorCell.Paragraphs.Add(graphSeries1);
        seriesColorCell.Margin = new MarginInfo(0, 3, 0, 0);
        seriesNameCell.Paragraphs.Add(seriesNameCellStyledText);
        seriesNameCell.Margin = new MarginInfo(0, 0, 0, 2);
    }

    public void VirtuesBarGraph(Page page, object[][] values)
    {
        Table table = new Table()
        {
            ColumnWidths = "500",
            Margin = new MarginInfo(20, 1, 20, 3),
        };
        page.Paragraphs.Add(table);

        Row row1 = table.Rows.Add();

        Cell title = row1.Cells.Add();

        Table TitleTable = new Table()
        {
            ColumnWidths ="50, 150, 150, 150",
            DefaultCellBorder = new BorderInfo(BorderSide.None)
        };

        title.Paragraphs.Add(TitleTable);

        Row TitleTableRow = TitleTable.Rows.Add();

        Cell TitleTableRowCell1 = TitleTableRow.Cells.Add();
        Cell TitleTableRowCell2 = TitleTableRow.Cells.Add();
        Cell TitleTableRowCell3 = TitleTableRow.Cells.Add();
        Cell TitleTableRowCell4 = TitleTableRow.Cells.Add();

        TextFragment titleTableRowCellStyledText = new TextFragment();

        TextSegment titleTableRowCellSegment = new TextSegment("\b")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 9,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial"),
                LineSpacing = 2,
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        titleTableRowCellStyledText.Segments.Add(titleTableRowCellSegment);
        titleTableRowCellStyledText.HorizontalAlignment = HorizontalAlignment.Center;
        titleTableRowCellStyledText.VerticalAlignment = VerticalAlignment.Center;

        TitleTableRowCell1.Paragraphs.Add(titleTableRowCellStyledText);

        VirtueTableTitle(TitleTableRowCell2, "POWER");
        VirtueTableTitle(TitleTableRowCell3, "PUSH");
        VirtueTableTitle(TitleTableRowCell4, "PAIN");



        Row row2 = table.Rows.Add();

        VirtueTableBarGraph(row2, values);










    }

    public void VirtueTableTitle(Cell title, string titleOfTable)
    {

        Table titleTable = new Table()
        {
            ColumnWidths = "140",
            
        };

        title.Paragraphs.Add(titleTable);

        

        BorderInfo borderInfo = new BorderInfo
        {
            Top = new GraphInfo
            {
                LineWidth = 0.1f, // Set the top border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#F4F4F2"))
            },
            Left = new GraphInfo
            {
                LineWidth = 0.2f, // Set the left border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#9497A0"))
            },
            Right = new GraphInfo
            {
                LineWidth = 0.5f, // Set the right border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#747C89"))
            },
            Bottom = new GraphInfo
            {
                LineWidth = 0.3f, // Set the bottom border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#747A90"))
            }
        };

        titleTable.Border = borderInfo;


        Row titleTableRow = titleTable.Rows.Add();

        Cell titleTableRowCell = titleTableRow.Cells.Add();

        BorderInfo titleTableRowCellBorderInfo = new BorderInfo
        {
            Top = new GraphInfo
            {
                LineWidth = 4.7f, // Set the top border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#CEDAF0"))
            },
            Left = new GraphInfo
            {
                LineWidth = 4.8f, // Set the left border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#BEC9E5"))
            },
            Right = new GraphInfo
            {
                LineWidth = 4.5f, // Set the right border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3BFD5"))
            },
            Bottom = new GraphInfo
            {
                LineWidth = 4.2f, // Set the bottom border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#A8B8DA"))
            }
        };

        titleTableRowCell.Border = titleTableRowCellBorderInfo;


        TextFragment titleTableRowCellStyledText = new TextFragment();

        TextSegment titleTableRowCellSegment = new TextSegment(titleOfTable)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 9,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial"),
                LineSpacing = 2,
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        titleTableRowCellStyledText.Segments.Add(titleTableRowCellSegment);
        titleTableRowCellStyledText.HorizontalAlignment = HorizontalAlignment.Center;
        titleTableRowCellStyledText.VerticalAlignment = VerticalAlignment.Center;

        titleTableRowCell.Paragraphs.Add(titleTableRowCellStyledText);
        titleTableRowCell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B7CBF0"));


    }


    public void VirtueTableBarGraph(Row row, Object[] values)
    {
        Table MainTable = new Table()
        {
            ColumnWidths = "500",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };

        Cell rowCell = row.Cells.Add();
        rowCell.Paragraphs.Add(MainTable);



        Table table = new Table()
        {
            ColumnWidths = "50 450",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };

        Cell rowCell = row.Cells.Add();
        rowCell.Paragraphs.Add(table);

        Row tableRow = table.Rows.Add();

        Cell tableRowCell1 = tableRow.Cells.Add();

        Cell tableRowCell2 = tableRow.Cells.Add();

        var lastValuesWithPercent = values.Select(row => $"{((object[])row)[((object[])row).Length - 1]}%").ToArray();

        var lastValues = values.Select(row => ((object[])row)[((object[])row).Length - 1]).ToArray();

        var result = string.Join("\b", lastValuesWithPercent);

        TextFragment percentageText = new TextFragment();

        TextSegment titleTableRowCellSegment = new TextSegment("\n100%\n80%\n60%\n40%\n20%\n0%")
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 9,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial"),
                LineSpacing = 2,
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        percentageText.Segments.Add(titleTableRowCellSegment);
        percentageText.HorizontalAlignment = HorizontalAlignment.Center;
        //percentageText.VerticalAlignment = VerticalAlignment.Center;

        tableRowCell1.Paragraphs.Add(percentageText);

        VirtueTableBarGraphDiagram(tableRowCell2, lastValues, result);





    }


    public void VirtueTableBarGraphDiagram(Cell tableRowCell2, Object[] values, string lastValuesWithPercent)
    {
        Table table = new Table()
        {
            ColumnWidths = "450",
            DefaultCellBorder = new BorderInfo(BorderSide.None),
        };

        tableRowCell2.Paragraphs.Add(table);

        Row row1 = table.Rows.Add();

        Cell cell1 = row1.Cells.Add();

        TextFragment percentageText = new TextFragment();

        TextSegment titleTableRowCellSegment = new TextSegment(lastValuesWithPercent)
        {
            TextState = new TextState
            {
                ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
                FontSize = 3,
                FontStyle = FontStyles.Bold,
                Font = FontRepository.FindFont("Arial"),
                LineSpacing = 2,
                HorizontalAlignment = HorizontalAlignment.Center,
            }
        };
        percentageText.Segments.Add(titleTableRowCellSegment);
        percentageText.HorizontalAlignment = HorizontalAlignment.Center;

        cell1.Paragraphs.Add(percentageText);


        Row row2 = table.Rows.Add();
        Cell cell2 = row2.Cells.Add();

        var numericValues = values.Select(v => Convert.ToInt32(v)).ToArray();

        var graph = new Graph(450.0, 100.0);

        cell2.Paragraphs.Add(graph);

        BorderInfo borderInfo = new BorderInfo
        {
            //Top = new GraphInfo
            //{
            //    LineWidth = 0.1f, // Set the top border width
            //    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#E6EAEB"))
            //},
            //Left = new GraphInfo
            //{
            //    LineWidth = 0.1f, // Set the left border width
            //    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#E6EAEB"))
            //},
            //Right = new GraphInfo
            //{
            //    LineWidth = 3, // Set the right border width
            //    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#A6A6A8"))
            //},
            Bottom = new GraphInfo
            {
                LineWidth = 3, // Set the bottom border width
                Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#BFBEC6"))
            }
        };

        table.Border = borderInfo;

        float xPosition = 5; 
        float spacing = 3; 

        foreach (var value in numericValues)
        {
            
            Rectangle bar = new Rectangle(xPosition, 0, 5, (float)(value * 0.50))
            {
                GraphInfo = new GraphInfo
                {
                    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3C8E7")),
                    FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#B3C8E7")),
                }
            };
            graph.Shapes.Add(bar);

           
            // Move xPosition for the next bar
            xPosition += 5 + spacing; // Adjust spacing between bars
        }






    }






























    //private static void AddBarGraph(Aspose.Pdf.Cell cell)
    //{
    //    // Define the table size
    //    float tableWidth = 150;
    //    float tableHeight = 120;

    //    // Create a table to hold the graph
    //    Table table = new Table
    //    {
    //        ColumnWidths = tableWidth.ToString(), // Table width matches the specified width
    //        DefaultCellBorder = new BorderInfo(BorderSide.None, 0.1f),
    //        DefaultCellPadding = new MarginInfo(2, 2, 2, 2) // Smaller padding to save space
    //    };

    //    // Add a row
    //    Row row = table.Rows.Add();

    //    // Add a cell for the bar graph
    //    Aspose.Pdf.Cell graphCell = row.Cells.Add();

    //    // Add the title "POWER"
    //    TextFragment title = new TextFragment("POWER")
    //    {
    //        TextState =
    //    {
    //        FontSize = 10, // Smaller font size for title
    //        ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#4472c8")),
    //        FontStyle = FontStyles.Bold
    //    },
    //        HorizontalAlignment = HorizontalAlignment.Center
    //    };
    //    graphCell.Paragraphs.Add(title);

    //    // Define bar values and labels
    //    float[] values = { 71.4f, 71.4f, 67.3f }; // Bar heights (percentage)
    //    string[] labels = { "Mentor", "Binder", "Principal" }; // Bar labels
    //    Aspose.Pdf.Color[] colors =
    //    {
    //    Aspose.Pdf.Color.FromRgb(System.Drawing.Color.DarkBlue),    // Mentor
    //    Aspose.Pdf.Color.FromRgb(System.Drawing.Color.MediumBlue),  // Binder
    //    Aspose.Pdf.Color.FromRgb(System.Drawing.Color.LightBlue)    // Principal
    //};

    //    // Define chart dimensions relative to the table
    //    float chartWidth = tableWidth;     // Match table width
    //    float chartHeight = tableHeight - 40; // Subtract height for title and labels
    //    float barWidth = chartWidth / (values.Length * 2); // Scale bar width dynamically
    //    float spaceBetweenBars = barWidth; // Equal space between bars
    //    float baselineY = 20;              // Baseline for bars

    //    // Create a graph for the chart
    //    Graph graph = new Graph(chartWidth, chartHeight)
    //    {
    //        Left = 50,
    //        Top = 100
    //    };

    //    // Draw bars
    //    for (int i = 0; i < values.Length; i++)
    //    {
    //        float barHeight = values[i] / 100 * (chartHeight - 40); // Scale bar height dynamically

    //        // Add a rectangle for the bar
    //        Rectangle rect = new Rectangle(
    //            i * (barWidth + 1), baselineY,
    //            barWidth, barHeight
    //        )
    //        {
    //            GraphInfo =
    //        {
    //            FillColor = colors[i],
    //            Color = colors[i]
    //        }
    //        };
    //        graph.Shapes.Add(rect);

    //        // Add percentage text above each bar
    //        TextFragment percentage = new TextFragment(values[i] + "%")
    //        {
    //            TextState = { FontSize = 8 }, // Smaller font size
    //            HorizontalAlignment = HorizontalAlignment.Center
    //        };
    //        percentage.Position = new Position(
    //            i * (barWidth + spaceBetweenBars) + barWidth / 2, // Center percentage above the bar
    //            baselineY + barHeight + 2 // Place percentage close to the bar top
    //        );
    //        graphCell.Paragraphs.Add(percentage);

    //        // Add label below each bar
    //        TextFragment label = new TextFragment(labels[i])
    //        {
    //            TextState = { FontSize = 8 }, // Smaller font size
    //            HorizontalAlignment = HorizontalAlignment.Center
    //        };
    //        label.Position = new Position(
    //            i * (barWidth + spaceBetweenBars) + barWidth / 2, // Center label under the bar
    //            baselineY - 10 // Place label below the baseline
    //        );
    //        graphCell.Paragraphs.Add(label);
    //    }

    //    // Add the graph to the graph cell
    //    graphCell.Paragraphs.Add(graph);

    //    // Add the table to the main cell
    //    cell.Paragraphs.Add(table);
    //}



}







































//public void BarGraph(Aspose.Pdf.Cell leftCell, float B2, float C2, float D2, string title, string series1, string series2, string series3)
//{
//    if (leftCell == null)
//    {
//        throw new ArgumentNullException(nameof(leftCell), "The leftCell argument cannot be null.");
//    }

//    CultureInfo culture = new CultureInfo("en-US");
//    Thread.CurrentThread.CurrentCulture = culture;
//    Thread.CurrentThread.CurrentUICulture = culture;

//    var workbook = new Aspose.Cells.Workbook();
//    var worksheet = workbook.Worksheets[0];

//    workbook.DefaultStyle.Font.Name = "Arial";

//    //worksheet.Cells["A1"].Value = "Category";
//    //worksheet.Cells["B1"].Value = "Mentor";
//    //worksheet.Cells["C1"].Value = "Binder";
//    //worksheet.Cells["D1"].Value = "Principal";

//    worksheet.Cells["B2"].Value = B2;
//    worksheet.Cells["C2"].Value = C2;
//    worksheet.Cells["D2"].Value = D2;


//    var styleCell1 = worksheet.Cells["B2"].GetStyle();
//    var styleCell2 = worksheet.Cells["C2"].GetStyle();
//    var styleCell3 = worksheet.Cells["D2"].GetStyle();

//    styleCell1.Pattern = BackgroundType.Solid;
//    styleCell2.Pattern = BackgroundType.Solid;
//    styleCell3.Pattern = BackgroundType.Solid;

//    styleCell1.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#2F5596");
//    styleCell2.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#4473C5");
//    styleCell3.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#B4C7E7");

//    worksheet.Cells["B2"].SetStyle(styleCell1);
//    worksheet.Cells["C2"].SetStyle(styleCell2);
//    worksheet.Cells["D2"].SetStyle(styleCell3);

//    //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#2F5596"));
//    //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#4473C5"));
//    //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#B4C7E8"));

//    int chartIndex = worksheet.Charts.Add(Aspose.Cells.Charts.ChartType.Column, 5, 0, 15, 5);
//    var chart = worksheet.Charts[chartIndex];

//    chart.Title.Text = title;
//    //chart.Title.Font.Color = Aspose.Cells.Drawing.ColorHelper.FromOleColor(unchecked((int)0xFFFFA500));

//    chart.Title.Font.Color = Aspose.Cells.Drawing.ColorHelper.FromOleColor(System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("#4472c8")));
//    chart.Title.Font.Name = "Arial";
//    chart.NSeries.Add("B2:D2", true);
//    chart.NSeries[0].Name = series1;
//    chart.NSeries[1].Name = series2;
//    chart.NSeries[2].Name = series3;


//    chart.PlotArea.Area.FillFormat.SolidFill.Color = Aspose.Cells.Drawing.ColorHelper.FromOleColor(0xFFFFFF);

//    chart.PlotArea.Area.Transparency = 1.0;
//    //chart.getChartArea().getBorder().setVisible(false);
//    chart.PlotArea.Border.Weight = 0;
//    chart.PlotArea.Border.Transparency = 1;


//    //chart.ValueAxis.IsVisible = false;
//    chart.CategoryAxis.IsVisible = false;;

//    var imagePath = title+"chart.png";
//    chart.ToImage(imagePath, new Aspose.Cells.Rendering.ImageOrPrintOptions
//    {
//        ImageType = Aspose.Cells.Drawing.ImageType.Png
//    });

//    if (leftCell != null)
//    {
//        Aspose.Pdf.Image chartImage = new Aspose.Pdf.Image
//        {
//            File = imagePath
//        };

//        chartImage.FixWidth = 150;
//        chartImage.FixHeight = 110;

//        chartImage.Margin = new Aspose.Pdf.MarginInfo
//        {
//            Left = 20,
//            Right = 5,
//            Top = 5,
//            Bottom = 5
//        };


//        leftCell.Paragraphs.Add(chartImage);

//        //Aspose.Pdf.Drawing.Graph shadowGraph = new Aspose.Pdf.Drawing.Graph((float)(chartImage.FixWidth + 10), (float)(chartImage.FixHeight + 10));

//        //// Define the shadow rectangle (slightly offset to the right and bottom of the image)
//        //Aspose.Pdf.Drawing.Rectangle shadowRect = new Aspose.Pdf.Drawing.Rectangle(
//        //    5,  // X-offset
//        //    5,  // Y-offset
//        //    (float)chartImage.FixWidth,   // Width
//        //    (float)chartImage.FixHeight   // Height
//        //);

//        //shadowRect.GraphInfo = new Aspose.Pdf.GraphInfo
//        //{
//        //    Color = Aspose.Pdf.Color.Gray, // Shadow outline color
//        //    FillColor = Aspose.Pdf.Color.FromArgb(1, 235, 237, 249) // Light shadow with 50% transparency
//        //};

//        //shadowGraph.Shapes.Add(shadowRect);
//        //leftCell.Paragraphs.Add(shadowGraph);
//    }
//    else
//    {
//        Console.WriteLine("Error: leftCell is null.");
//    }
//}

//public void MainBarGraph(Page page, params object[][] values)
//{
//    if (page == null)
//    {
//        throw new ArgumentNullException(nameof(page), "The leftCell argument cannot be null.");
//    }

//    // Set culture for formatting
//    CultureInfo culture = new CultureInfo("en-US");
//    Thread.CurrentThread.CurrentCulture = culture;
//    Thread.CurrentThread.CurrentUICulture = culture;

//    // Create an Excel workbook
//    var workbook = new Workbook();
//    var worksheet = workbook.Worksheets[0];

//    workbook.DefaultStyle.Font.Name = "Arial";

//    // Dynamically fill worksheet cells with the input values
//    for (int i = 0; i < values.Length; i++)
//    {
//        // Set values for columns dynamically
//        var va = values[i][2];
//        string columnLetter = ((char)(((int)'B') + i)).ToString();
//        worksheet.Cells[$"{columnLetter}2"].Value = values[i][2];  // Set data values in cells B2, C2, D2, etc.
//        worksheet.Cells[$"{columnLetter}3"].Value = values[i][1];
//    }

//    // Set cell styles for each column
//    for (int i = 0; i < values.Length; i++)
//    {
//        string columnLetter = ((char)('B' + i)).ToString();
//        var style = worksheet.Cells[$"{columnLetter}2"].GetStyle();
//        style.Pattern = BackgroundType.Solid;
//        style.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#2F5596"); // You can dynamically assign colors
//        worksheet.Cells[$"{columnLetter}2"].SetStyle(style);
//    }

//    // Set chart styles for cells B2, C2, D2
//    SetCellStyle(worksheet, "B2", "#2F5596");
//    SetCellStyle(worksheet, "C2", "#4473C5");
//    SetCellStyle(worksheet, "D2", "#B4C7E7");

//    // Create a chart and add series dynamically
//    int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
//    var chart = worksheet.Charts[chartIndex];

//    // Set chart title based on categories from image
//    //chart.Title.Text = "Virtues Chart - Power, Push, and Pain";
//    //chart.Title.Font.Color = ColorHelper.FromOleColor(0x2F5596);
//    //chart.Title.Font.Name = "Arial";

//    // Add the series to the chart dynamically
//    int seriesIndex = chart.NSeries.Add($"B2:{(char)('B' + values.Length - 1)}2", true);
//    chart.NSeries.CategoryData = $"B3:{(char)(((int)'B') + values.Length - 1)}3";

//    //chart.CategoryAxis.TickLabels.AlignmentType = TextOrientationType.NoRotation; // This rotates the labels vertically
//    chart.CategoryAxis.TickLabels.Font.Size = 8;

//    // Loop through each series and set its name (category titles like Power, Push, Pain)
//    for (int i = 0; i < values.Length; i++)
//    {
//        chart.NSeries[i].Name = "";//values[i][1].ToString();  // These names should be like "Power," "Push," etc.
//    }



//    // Set chart plot area background
//    chart.PlotArea.Area.FillFormat.SolidFill.Color = ColorHelper.FromOleColor(0xFFFFFF);

//    // Save the chart as an image
//    var imagePath = "Mainchart.png";
//    chart.ToImage(imagePath, new ImageOrPrintOptions
//    {
//        ImageType = ImageType.Png
//    });

//    // If leftCell is not null, embed the chart image in the PDF
//    if (page != null)
//    {
//        Aspose.Pdf.Image chartImage = new Aspose.Pdf.Image
//        {
//            File = imagePath
//        };

//        chartImage.FixWidth = 550;
//        chartImage.FixHeight = 130;

//        chartImage.Margin = new MarginInfo
//        {
//            Left = 10,
//            Right = 5,
//            Top = 5,
//            Bottom = 5
//        };

//        page.Paragraphs.Add(chartImage);
//    }
//    else
//    {
//        Console.WriteLine("Error: leftCell is null.");
//    }
//}

//private void SetCellStyle(Worksheet worksheet, string cell, string color)
//{
//    var style = worksheet.Cells[cell].GetStyle();
//    style.Pattern = BackgroundType.Solid;
//    style.ForegroundColor = System.Drawing.ColorTranslator.FromHtml(color);
//    worksheet.Cells[cell].SetStyle(style);
//}





