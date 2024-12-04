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
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using System.Globalization;


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
        Program p = new Program();
        p.BarGraph(graphCell, 71.4f, 71.4f, 64.5f, "POWER", "Mentor", "Binder", "Principal");
        p.BarGraph(graphCell, 63.7f, 62.3f, 55.6f, "PUSH", "Charmer", "Guardian", "Dominion");
        p.BarGraph(graphCell, 51.6f, 39.9f, 37.7f, "PAIN", "Harmonizer", "Visualizer", "Angel");

        

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


    public void BarGraph(Aspose.Pdf.Cell leftCell, float B2, float C2, float D2, string title, string series1, string series2, string series3)
    {
        if (leftCell == null)
        {
            throw new ArgumentNullException(nameof(leftCell), "The leftCell argument cannot be null.");
        }

        CultureInfo culture = new CultureInfo("en-US");
        Thread.CurrentThread.CurrentCulture = culture;
        Thread.CurrentThread.CurrentUICulture = culture;

        var workbook = new Aspose.Cells.Workbook();
        var worksheet = workbook.Worksheets[0];

        workbook.DefaultStyle.Font.Name = "Arial";

        //worksheet.Cells["A1"].Value = "Category";
        //worksheet.Cells["B1"].Value = "Mentor";
        //worksheet.Cells["C1"].Value = "Binder";
        //worksheet.Cells["D1"].Value = "Principal";

        worksheet.Cells["B2"].Value = B2;
        worksheet.Cells["C2"].Value = C2;
        worksheet.Cells["D2"].Value = D2;


        var styleCell1 = worksheet.Cells["B2"].GetStyle();
        var styleCell2 = worksheet.Cells["C2"].GetStyle();
        var styleCell3 = worksheet.Cells["D2"].GetStyle();

        styleCell1.Pattern = BackgroundType.Solid;
        styleCell2.Pattern = BackgroundType.Solid;
        styleCell3.Pattern = BackgroundType.Solid;

        styleCell1.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#2F5596");
        styleCell2.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#4473C5");
        styleCell3.ForegroundColor = System.Drawing.ColorTranslator.FromHtml("#B4C7E7");

        worksheet.Cells["B2"].SetStyle(styleCell1);
        worksheet.Cells["C2"].SetStyle(styleCell2);
        worksheet.Cells["D2"].SetStyle(styleCell3);

        //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#2F5596"));
        //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#4473C5"));
        //worksheet.Cells["B2"].SetStyle(System.Drawing.ColorTranslator.FromHtml("#B4C7E8"));

        int chartIndex = worksheet.Charts.Add(Aspose.Cells.Charts.ChartType.Column, 5, 0, 15, 5);
        var chart = worksheet.Charts[chartIndex];

        chart.Title.Text = title;
        chart.Title.Font.Color = Aspose.Cells.Drawing.ColorHelper.FromOleColor(0x2F5596);
        chart.Title.Font.Name = "Arial";
        chart.NSeries.Add("B2:D2", true);
        chart.NSeries[0].Name = series1;
        chart.NSeries[1].Name = series2;
        chart.NSeries[2].Name = series3;


        chart.PlotArea.Area.FillFormat.SolidFill.Color = Aspose.Cells.Drawing.ColorHelper.FromOleColor(0xFFFFFF);

        //var seriesColor = Aspose.Cells.Drawing.ColorHelper.FromOleColor(0xFFFFFF);

        //chart.NSeries[1].Area.FillFormat.SolidFill.Color = seriesColor;
        //chart.NSeries[2].Area.FillFormat.SolidFill.Color = seriesColor;


        //chart.ValueAxis.IsVisible = false;
        chart.CategoryAxis.IsVisible = false;;

        var imagePath = title+"chart.png";
        chart.ToImage(imagePath, new Aspose.Cells.Rendering.ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png
        });

        if (leftCell != null)
        {
            Aspose.Pdf.Image chartImage = new Aspose.Pdf.Image
            {
                File = imagePath
            };

            chartImage.FixWidth = 150;
            chartImage.FixHeight = 110;

            chartImage.Margin = new Aspose.Pdf.MarginInfo
            {
                Left = 20,
                Right = 5,
                Top = 5,
                Bottom = 5
            };

           leftCell.Paragraphs.Add(chartImage);

            //Aspose.Pdf.Drawing.Graph shadowGraph = new Aspose.Pdf.Drawing.Graph((float)(chartImage.FixWidth + 10), (float)(chartImage.FixHeight + 10));

            //// Define the shadow rectangle (slightly offset to the right and bottom of the image)
            //Aspose.Pdf.Drawing.Rectangle shadowRect = new Aspose.Pdf.Drawing.Rectangle(
            //    5,  // X-offset
            //    5,  // Y-offset
            //    (float)chartImage.FixWidth,   // Width
            //    (float)chartImage.FixHeight   // Height
            //);

            //shadowRect.GraphInfo = new Aspose.Pdf.GraphInfo
            //{
            //    Color = Aspose.Pdf.Color.Gray, // Shadow outline color
            //    FillColor = Aspose.Pdf.Color.FromArgb(1, 235, 237, 249) // Light shadow with 50% transparency
            //};

            //shadowGraph.Shapes.Add(shadowRect);
            //leftCell.Paragraphs.Add(shadowGraph);
        }
        else
        {
            Console.WriteLine("Error: leftCell is null.");
        }
    }



}
