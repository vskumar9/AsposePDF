using Aspose.Imaging.Xmp.Types.Complex.Dimensions;
using Aspose.Pdf;
using Aspose.Pdf.Drawing;
using Aspose.Pdf.Text;
using Azure;
using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Drawing.Printing;

namespace Aspose
{
    internal class OldCode
    {
        static void main(string[] args)
        {
            string connectionString = "data source=PTSQLTESTDB01;database=TalentAqu;integrated security=true;trustservercertificate = true;";

            Document pdfDocument = new Document();
            Page page = pdfDocument.Pages.Add();
            page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            page.PageInfo.Margin.Left = 10;
            page.PageInfo.Margin.Right = 10;
            page.PageInfo.Margin.Top = 10;
            page.PageInfo.Margin.Bottom = 10;

            MainPage(page);
            SecondPage(page, connectionString);

            ThirdPage(page);






            TextFragment header = new TextFragment("SAMPLE - NOT TO BE SHARED WITHOUT WRITTEN CONSENT OF LIMENEAL SOLUTIONS - FZCO")
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = { Top = 30 },
                TextState =
            {
                Font = FontRepository.FindFont("Arial-Bold"),
                FontSize = 12
            }
            };

            page.Paragraphs.Add(header);

            Table table = new Table
            {
                ColumnWidths = "100 100 100 250",
                Margin = new MarginInfo(10, 10, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                RepeatingRowsCount = 2,
                IsBroken = true
            };

            page.Paragraphs.Add(table);

            string[] headerTexts = { "Virtues", "Emphasis", "Focus", "Role requirement Vs Individuals' expression" };

            Row headerRow = table.Rows.Add();
            foreach (string headerText in headerTexts)
            {
                headerRow.Cells.Add(headerText);
            }

            foreach (Cell cell in headerRow.Cells)
            {
                cell.BackgroundColor = Color.LightBlue;
                cell.DefaultCellTextState = new TextState
                {
                    ForegroundColor = Color.Blue,
                    FontSize = 12
                };
            }

            Row percentageRow = table.Rows.Add();
            percentageRow.DefaultCellPadding = new MarginInfo(0, 0, 0, 10);
            percentageRow.Cells.Add("");
            percentageRow.Cells.Add("");
            percentageRow.Cells.Add("");
            Cell progressBarCell = percentageRow.Cells.Add();

            Table table1 = new Table
            {
                ColumnWidths = "150 100",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                RepeatingRowsCount = 1,
                IsBroken = true
            };

            Row row1 = table1.Rows.Add();
            Row row2 = table1.Rows.Add();
            Cell textColumn = row2.Cells.Add();
            TextFragment textFragment = new TextFragment("0%  20%  40%  60%  80%  100%");
            textFragment.TextState.FontSize = 10;
            textColumn.Paragraphs.Add(textFragment);

            progressBarCell.Paragraphs.Add(table1);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query1 = "SELECT * FROM TalentTable";

                using (SqlCommand command = new SqlCommand(query1, connection))
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string virtue = reader["Virtue"].ToString()!;
                        string emphasis = reader["Emphasis"].ToString()!;
                        string focus = reader["Focus"].ToString()!;
                        int requirement = Convert.ToInt32(reader["Requirement"]);
                        int expression = Convert.ToInt32(reader["Expression"]);
                        AddRowWithProgressBar(table, virtue!, emphasis!, focus!, requirement, expression);
                    }
                }
            }


            TextFragment header1 = new TextFragment("SAMPLE - NOT TO BE SHARED WITHOUT WRITTEN CONSENT OF LIMENEAL SOLUTIONS - FZCO")
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = { Top = 20 },
            };
            header1.TextState.Font = FontRepository.FindFont("Arial-Bold");
            header1.TextState.FontSize = 12;
            page.Paragraphs.Add(header1);

            page.Paragraphs.Add(new TextFragment("\n"));

            AddTextByPostion(page, "Effectiveness Score", 120, 620, 12);
            AddTextByPostion(page, "Fulfilment Score", 390, 620, 12);

            DrawCircularProgressBarWithText(98, 50, 180, page);
            DrawCircularProgressBarWithText(98, 330, 180, page);

            AddTextByPostion(page, "98%", 150, 545, 12);
            AddTextByPostion(page, "95%", 430, 545, 12);

            page.Paragraphs.Add(new TextFragment("\n\n\n\n\n"));

            Table mainTable = new Table
            {
                ColumnWidths = "200 100 200",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                Margin = new MarginInfo(50, 0, 0, 0)
            };

            page.Paragraphs.Add(mainTable);

            Row row = mainTable.Rows.Add();

            AddCardContent(row, "YOU WOULD BE MOST EFFECTIVE", "This means, the degree of expression of almost all the virtues required to successfully deliver in this role align with the degree expressed by you.", 5);

            Cell spacerCell = row.Cells.Add("");
            spacerCell.Border = new BorderInfo(BorderSide.None);


            AddCardContent(row, "YOU ARE HIGHLY FULFILLED", "In this state, there is fluidity between your body and mind, where you are totally absorbed in the tasks involved, beyond the point of distraction with increased commitment, motivation, concentration and performance.", 5);

            //pdfDocument.Save(@"D:\Aspose_PdfAspose_PDF.pdf");
        }

        static void AddTextByPostion(Page page, string text, int x, int y, int fontSize)
        {
            TextFragment text1 = new TextFragment(text)
            {
                Position = new Position(x, y),
            };

            text1.TextState.FontSize = fontSize;
            text1.TextState.Font = FontRepository.FindFont("Arial");
            text1.TextState.ForegroundColor = Color.Blue;

            page.Paragraphs.Add(text1);
        }

        static void AddRowWithProgressBar(Table table, string virtues, string emphasis, string focus, int requirement, int expression)
        {
            Row row = table.Rows.Add();

            row.DefaultCellPadding = new MarginInfo(0, 0, 0, 10);

            RowText(row, virtues);
            RowText(row, emphasis);
            RowText(row, focus);

            Cell progressBarCell = row.Cells.Add();

            Table nestedTable = new Table
            {
                ColumnWidths = "150 100",
                DefaultCellBorder = new BorderInfo(BorderSide.None)
            };

            Row progressBarRow = nestedTable.Rows.Add();

            Graph graph = new Graph(200.0, 30.0);

            Aspose.Pdf.Drawing.Rectangle requirementBar = new Aspose.Pdf.Drawing.Rectangle(0, 0, (float)(requirement * 1.25), 10)
            {
                GraphInfo = new GraphInfo
                {
                    Color = Color.Blue,
                    FillColor = Color.Blue
                }
            };
            graph.Shapes.Add(requirementBar);

            Aspose.Pdf.Drawing.Rectangle expressionBar = new Aspose.Pdf.Drawing.Rectangle(0, 15, (float)(expression * 1.25), 10)
            {
                GraphInfo = new GraphInfo
                {
                    Color = Color.Orange,
                    FillColor = Color.Orange
                }
            };
            graph.Shapes.Add(expressionBar);

            Cell progressBarColumn = progressBarRow.Cells.Add();
            progressBarColumn.Paragraphs.Add(graph);

            Cell textColumn = progressBarRow.Cells.Add();
            TextFragment textFragment = new TextFragment($"{expression}% Expression\n{requirement}% Requirement");
            textFragment.TextState.FontSize = 11;
            textFragment.TextState.ForegroundColor = Color.Blue;
            textFragment.TextState.LineSpacing = 5;
            textColumn.Paragraphs.Add(textFragment);

            progressBarCell.Paragraphs.Add(nestedTable);
        }

        static void RowText(Row row, string text)
        {
            Cell textColumn0 = row.Cells.Add();
            TextFragment textFragment0 = new TextFragment(text);
            textFragment0.TextState.FontSize = 11;
            textFragment0.TextState.ForegroundColor = Color.Blue;
            textColumn0.Paragraphs.Add(textFragment0);
        }

        static void DrawCircularProgressBarWithText(int percentage, float x, float y, Page page)
        {
            Graph graph = new Graph(200.0, 200.0);
            graph.Left = x;
            graph.Top = y;

            Circle outerCircle = new Circle(100, 100, 45);
            outerCircle.GraphInfo = new GraphInfo
            {
                LineWidth = 6,
                Color = Color.FromRgb(1.0f, 0.3f, 0.0f)
            };

            graph.Shapes.Add(outerCircle);

            Circle innerCircle = new Circle(100, 100, 30);
            innerCircle.GraphInfo = new GraphInfo
            {
                LineWidth = 2,
                Color = Color.LightGray,
                FillColor = Color.LightGray
            };
            graph.Shapes.Add(innerCircle);

            page.Paragraphs.Add(graph);
        }

        static void AddCardContent(Row mainRow, string title, string body, float margin)
        {
            Cell cardCell = mainRow.Cells.Add("");
            cardCell.BackgroundColor = Color.LightBlue;
            cardCell.Border = new BorderInfo(BorderSide.All, 1f, Color.Blue);
            cardCell.Margin = new MarginInfo(margin, 0, margin, 0);

            TextFragment titleFragment = new TextFragment(title);
            titleFragment.TextState.FontSize = 10;
            titleFragment.TextState.Font = FontRepository.FindFont("Arial-Bold");
            titleFragment.TextState.ForegroundColor = Color.Blue;
            titleFragment.HorizontalAlignment = HorizontalAlignment.Left;
            titleFragment.VerticalAlignment = VerticalAlignment.Top;
            cardCell.Paragraphs.Add(titleFragment);

            TextFragment bodyFragment = new TextFragment(body);
            bodyFragment.TextState.FontSize = 10;
            bodyFragment.TextState.Font = FontRepository.FindFont("Arial");
            bodyFragment.TextState.ForegroundColor = Color.Blue;
            bodyFragment.HorizontalAlignment = HorizontalAlignment.Left;
            cardCell.Paragraphs.Add(bodyFragment);
        }

        static void MainPage(Page page)
        {

            Image leftImage = new Image
            {
                File = @"D:\Aspose_Pdf\Aspose\Aspose\asset\MainPageCards.jpeg",
                FixWidth = 400,
                FixHeight = 100,
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = new MarginInfo { Top = 50 },
            };
            page.Paragraphs.Add(leftImage);

            Image rightImage = new Image
            {
                File = @"D:\Aspose_Pdf\Aspose\Aspose\asset\MainPageBanner.jpeg",
                FixWidth = 170,
                FixHeight = 170,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new MarginInfo { Top = -100 },
            };
            page.Paragraphs.Add(rightImage);

            TextFragment header = new TextFragment("LIMENEAL TALENT\nFULFILMENT\nREPORT")
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                Margin = { Top = 50, Left = 45 },
                TextState =
                {
                    Font = FontRepository.FindFont("Times New Roman"),
                    FontSize = 40,
                    ForegroundColor = Color.Blue,
                    FontStyle = FontStyles.Bold,
                    LineSpacing = 7
                }
            };
            page.Paragraphs.Add(header);

            TextFragment header1 = new TextFragment("SAMPLE - NOT TO BE SHARED WITHOUT WRITTEN CONSENT OF LIMENEAL SOLUTIONS - FZCO")
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = { Top = 20 },
            };
            header1.TextState.Font = FontRepository.FindFont("Arial-Bold");
            header1.TextState.FontSize = 12;
            page.Paragraphs.Add(header1);

            Image CenterImage = new Image
            {
                File = @"D:\Aspose_Pdf\Aspose\Aspose\asset\MainPageImage.jpeg",
                FixWidth = 500,
                FixHeight = 340,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new MarginInfo { Top = 10 },
            };

            page.Paragraphs.Add(CenterImage);

            TextFragment rights = new TextFragment("Limeneal Wheel Dimensions, Virtues and Traits are Copyrighted. All Rights Reserved.")
            {
                Margin = { Top = 4, Bottom = 10, Left = 25 },
            };
            rights.TextState.Font = FontRepository.FindFont("Times New Roman");
            rights.TextState.FontSize = 6;
            page.Paragraphs.Add(rights);


        }

        static void SecondPage(Page page, string connectionString)
        {
            TextFragment header = new TextFragment()
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = { Top = 30 }
            };

            TextSegment mainText = new TextSegment("9 DIMENSIONS OF THE LIMENEAL WHEEL")
            {
                TextState = new TextState
                {
                    Font = FontRepository.FindFont("Arial-Bold"),
                    FontSize = 18,
                    ForegroundColor = Color.Blue
                }
            };

            TextSegment registeredSymbol = new TextSegment("®")
            {
                TextState = new TextState
                {
                    Font = FontRepository.FindFont("Arial-Bold"),
                    FontSize = 8,
                    ForegroundColor = Color.Blue,
                },
            };

            registeredSymbol.Position = new Position(header.Position.XIndent + 290, header.Position.YIndent + 10);

            header.Segments.Add(mainText);
            header.Segments.Add(registeredSymbol);

            page.Paragraphs.Add(header);


            TextFragment header1 = new TextFragment("SAMPLE - NOT TO BE SHARED WITHOUT WRITTEN CONSENT OF LIMENEAL SOLUTIONS - FZCO")
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = { Top = 10 },
            };
            header1.TextState.Font = FontRepository.FindFont("Arial-Bold");
            header1.TextState.FontSize = 12;
            page.Paragraphs.Add(header1);

            Image leftImage = new Image
            {
                File = @"D:\Aspose_Pdf\Aspose\Aspose\asset\SecondPageWheel.jpeg",
                FixWidth = 180,
                FixHeight = 160,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new MarginInfo { Top = 10 },
            };
            page.Paragraphs.Add(leftImage);

            TextFragment intro = new TextFragment("INTRODUCTION")
            {

                Margin = { Top = 10 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial-Bold"),
                    FontSize = 14,
                    ForegroundColor = Color.Blue,
                    HorizontalAlignment = HorizontalAlignment.Center
                }
            };
            page.Paragraphs.Add(intro);


            TextFragment content1 = new TextFragment("Limeneal Wheel® is a Human Fulfilment Model, built on the science and philosophy of virtues to discover your Purpose, Passion, and Potential, manifested through 36 virtues and classified under 9 unique dimensions.")
            {

                Margin = { Top = 3, Left = 30 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial"),
                    FontSize = 13,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 5
                }
            };
            page.Paragraphs.Add(content1);

            TextFragment content2 = new TextFragment("Limeneal Wheel® is an intellectual proprietary work created and developed by Vivian Selvathurai Alfred, significantly influenced by the work on virtues by the 12th-century philosopher, Thomas Aquinas, and inspiration from ancient philosophers - Augustine of Hippo, Plato, Aristotle, and Thiruvalluvar, the 6th century philosopher from South India.")
            {

                Margin = { Top = 2, Left = 30 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial"),
                    FontSize = 13,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 5
                }
            };
            page.Paragraphs.Add(content2);

            TextFragment content3 = new TextFragment("The word Limeneal is derived from the Latin word, “Limen” meaning “Gateway” and “Neal” in Sanskrit meaning “Champion”. Together, they indicate “Gateway of a Champion”.")
            {

                Margin = { Top = 2, Left = 30 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial"),
                    FontSize = 13,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 5
                }
            };
            page.Paragraphs.Add(content3);

            TextFragment content4 = new TextFragment("The Limeneal Wheel® virtue mapper tool involves a proprietary virtues benchmarking process that compares degree of virtues required for a career or a job role with those expressed by individuals. Such comparison will then objectively demonstrate, an individual's inclination, effectiveness and fulfillment score.")
            {

                Margin = { Top = 2, Left = 30 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial"),
                    FontSize = 13,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 5
                }
            };
            page.Paragraphs.Add(content4);

            TextFragment content5 = new TextFragment("“Where the roads of your Purpose, Passion and Potential meet, your True Calling is found.”")
            {

                Margin = { Top = 2, Left = 30 },
                TextState =
                {
                    Font = FontRepository.FindFont("Arial-Bold"),
                    FontSize = 12,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 5
                }
            };
            page.Paragraphs.Add(content5);

            Table table = new Table
            {
                ColumnWidths = "140 140 250",
                Margin = new MarginInfo(30, 10, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.Box, .5f, Color.Blue),
            };

            page.Paragraphs.Add(table);

            string[] Text_1 = { "PURPOSE\n\nPASSION\n\nPOTENTIAL", "Is your guiding star\n\nIs what lights your fire\n\nIs what you can become", "You need all 3 in right proportions vs. aspired career/role.\nWe help you to assess this so that not only you follow your purpose but have sufficient fire and strengths to sustain it." };

            Row Row_1 = table.Rows.Add();
            foreach (string headerText in Text_1)
            {
                Cell cell = Row_1.Cells.Add();

                cell.Margin = new MarginInfo(5, 0, 0, 2);

                TextFragment fragment = new TextFragment(headerText)
                {
                    TextState =
                {
                    FontSize = 12,
                    ForegroundColor = Color.Blue
                },
                };

                cell.Paragraphs.Add(fragment);
            }



            string query = @"SELECT Dimension, ProductExpertScore, SalesSpecialistScore FROM BenchmarkScores";

            // Fetch data from the database
            DataTable scoresData = FetchDataFromDatabase(connectionString, query);

            // Generate PDF with dynamic bar chart
            GeneratePdf(scoresData, page);

        }

        static void ThirdPage(Page page)
        {
            Table TableOfContext = new Table
            {
                ColumnWidths = "550",
                Margin = new MarginInfo(10, 10, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
            };

            page.Paragraphs.Add(TableOfContext);

            string[] TableOfContextHeaderTexts = { "TABLE OF CONTENTS" };

            Row TableOfContexHeaderRow = TableOfContext.Rows.Add();
            foreach (string headerText in TableOfContextHeaderTexts)
            {
                TableOfContexHeaderRow.Cells.Add(headerText);
            }

            foreach (Cell cell in TableOfContexHeaderRow.Cells)
            {
                cell.BackgroundColor = Color.LightBlue;
                cell.DefaultCellTextState = new TextState
                {
                    ForegroundColor = Color.Blue,
                    FontSize = 12,
                    FontStyle = FontStyles.Bold,
                };
                cell.Paragraphs[0].HorizontalAlignment = HorizontalAlignment.Center;
                cell.Paragraphs[0].Margin = new MarginInfo(0, 5, 0, 5);
            }

            Table table = new Table
            {
                ColumnWidths = "225 225", // Left column (text and image), Right column (page buttons)
                Margin = new MarginInfo(30, 10, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None)
            };

            // Add the table to the page
            page.Paragraphs.Add(table);

            // Left Column: Text and Image
            Row contentRow = table.Rows.Add();

            // Add text content to the left column
            Cell textCell = contentRow.Cells.Add();
            TextFragment textFragment = new TextFragment("LIMITED TALENT DRIVERS")
            {
                Margin = { Top = 50, Left = 30 },
                TextState =
            {
                FontSize = 12,
                FontStyle = FontStyles.Bold,
                ForegroundColor = Color.Blue,
                LineSpacing = 3
            }
            };

            TextFragment textFragment1 = new TextFragment("Top 3 Dimensions of an individual indicating Orientation, Purpose, Passion, Attitude, and preferred work environment.Identifies preferred communication style of the candidate.")
            {
                Margin = { Left = 30 },
                TextState =
            {
                FontSize = 10,
                ForegroundColor = Color.Blue,
                LineSpacing = 3
            }
            };


            textCell.Paragraphs.Add(textFragment);
            textCell.Paragraphs.Add(textFragment1);

            // Add image to the left column (if you have an image to add)
            //Cell imageCell = contentRow.Cells.Add();
            //Image image = new Image
            //{
            //    File = "path-to-your-image.png"  // Replace with your image file path
            //};
            //imageCell.Paragraphs.Add(image);

            // Right Column: Page Navigation
            Cell navCell = contentRow.Cells.Add();

            // Add "Pages" label
            TextFragment pagesLabel = new TextFragment("Pages")
            {
                TextState =
            {
                FontSize = 12,
                ForegroundColor = Color.Blue
            }
            };
            navCell.Paragraphs.Add(pagesLabel);


            //TextFragment page4Button = new TextFragment("4")
            //{
            //    TextState = { ForegroundColor = Color.White },
            //    BackgroundColor = Color.DarkBlue
            //};
            //TextFragment page5Button = new TextFragment("5")
            //{
            //    TextState = { ForegroundColor = Color.White },
            //    BackgroundColor = Color.DarkBlue
            //};

            // Add page navigation buttons (as TextFragments)
            //navCell.Paragraphs.Add(page4Button);
            //navCell.Paragraphs.Add(new TextFragment("to"));
            //navCell.Paragraphs.Add(page5Button);

        }


        static DataTable FetchDataFromDatabase(string connectionString, string query)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        dataTable.Load(reader);
                    }
                }
            }

            return dataTable;
        }

        static void GeneratePdf(DataTable scoresData, Page page)
        {
            Table table = new Table
            {
                ColumnWidths = "150 380",
                Margin = new MarginInfo(30, 10, 10, 2),
                DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5f, Aspose.Pdf.Color.Blue)
            };

            //table.DefaultCellPadding = new MarginInfo(0, 0, 5, 0);


            Row contentRow = table.Rows.Add();

            Cell descriptionCell = contentRow.Cells.Add();
            descriptionCell.Paragraphs.Add(new TextFragment("Illustrative benchmark scores of two similar roles in an Electric Vehicle (EV) Showroom presented alongside.\n\n\n\nVirtues are required to a different degree in each role to achieve effectiveness and fulfillment.")
            {
                TextState =
                {
                    FontSize = 12,
                    ForegroundColor = Color.Blue,
                    LineSpacing = 3,
                },
                Margin = new MarginInfo(5, 5, 5, 5)
            });


            // Right Column: Graph
            Cell graphCell = contentRow.Cells.Add();
            //Graph graph = CreateCompositeGraph(scoresData, page); // Create the full bar graph
            //graphCell.Paragraphs.Add(graph);

            // Add the table to the page
            page.Paragraphs.Add(table);
        }



        //static Graph CreateCompositeGraph(DataTable scoresData, Page page)
        //{
        //    double maxValue = 50; // Max value for scaling
        //    float barWidth = 10;   // Width of each bar
        //    float spacing = 50;    // Space between sets of bars
        //    float xOffset = 10;    // X-offset for the graph
        //    float yOffset = 10;    // Y-offset from bottom
        //    float graphWidth = 500; // Overall graph width
        //    float graphHeight = 150; // Overall graph height

        //    Graph graph = new Graph(graphWidth, graphHeight);

        //    for (int i = 0; i < scoresData.Rows.Count; i++)
        //    {
        //        DataRow row = scoresData.Rows[i];

        //        string dimension = row["Dimension"]?.ToString() ?? throw new NullReferenceException("The 'Dimension' column contains a null value.");
        //        double productExpertScore = Convert.ToDouble(row["ProductExpertScore"]);
        //        double salesSpecialistScore = Convert.ToDouble(row["SalesSpecialistScore"]);

        //        float productExpertHeight = (float)((productExpertScore / maxValue) * graphHeight);
        //        float salesSpecialistHeight = (float)((salesSpecialistScore / maxValue) * graphHeight);

        //        // Calculate positions for bars
        //        float productExpertX = xOffset + (i * spacing);
        //        float salesSpecialistX = productExpertX + barWidth + 10;

        //        // Create bars
        //        Aspose.Pdf.Drawing.Rectangle productExpertBar = new Aspose.Pdf.Drawing.Rectangle(
        //            productExpertX, yOffset, productExpertX + barWidth, yOffset + productExpertHeight)
        //        {
        //            GraphInfo = { FillColor = Aspose.Pdf.Color.Blue }
        //        };
        //        Aspose.Pdf.Drawing.Rectangle salesSpecialistBar = new Aspose.Pdf.Drawing.Rectangle(
        //            salesSpecialistX, yOffset, salesSpecialistX + barWidth, yOffset + salesSpecialistHeight)
        //        {
        //            GraphInfo = { FillColor = Aspose.Pdf.Color.Green }
        //        };

        //        // Add bars to the graph
        //        graph.Shapes.Add(productExpertBar);
        //        graph.Shapes.Add(salesSpecialistBar);

        //        // Add dimension labels
        //        TextFragment textFragment = new TextFragment(dimension)
        //        {
        //            Position = new Aspose.Pdf.Text.Position((productExpertX + salesSpecialistX) / 2, yOffset - 15),
        //            TextState =
        //        {
        //            FontSize = 10,
        //            ForegroundColor = Aspose.Pdf.Color.Black,
        //            FontStyle = Aspose.Pdf.Text.FontStyles.Bold
        //        }
        //        };
        //        // Add text to the page (you can add it directly to the page or graph depending on your structure)
        //        page.Paragraphs.Add(textFragment);
        //    }

        //    // Add legend to the graph
        //    AddLegend(graph, graphWidth, graphHeight, page);

        //    return graph;
        //}


        //static void AddLegend(Graph graph, float graphWidth, float graphHeight, Page page)
        //{
        //    // Add legend for Product Expert
        //    Aspose.Pdf.Drawing.Rectangle blueBox = new Aspose.Pdf.Drawing.Rectangle(graphWidth - 200, graphHeight - 20, graphWidth - 180, graphHeight - 10)
        //    {
        //        GraphInfo = { FillColor = Aspose.Pdf.Color.Blue }
        //    };
        //    graph.Shapes.Add(blueBox);

        //    // Add text for Product Expert
        //    TextFragment productExpertText = new TextFragment("Product Expert")
        //    {
        //        Position = new Aspose.Pdf.Text.Position(graphWidth - 170, graphHeight - 15),
        //        TextState = { FontSize = 10 }
        //    };
        //    // Add the text fragment to the page
        //    page.Paragraphs.Add(productExpertText);

        //    // Add legend for Sales Specialist
        //    Aspose.Pdf.Drawing.Rectangle greenBox = new Aspose.Pdf.Drawing.Rectangle(graphWidth - 200, graphHeight - 40, graphWidth - 180, graphHeight - 30)
        //    {
        //        GraphInfo = { FillColor = Aspose.Pdf.Color.Green }
        //    };
        //    graph.Shapes.Add(greenBox);

        //    // Add text for Sales Specialist
        //    TextFragment salesSpecialistText = new TextFragment("Sales Specialist")
        //    {
        //        Position = new Aspose.Pdf.Text.Position(graphWidth - 170, graphHeight - 35),
        //        TextState = { FontSize = 10 }
        //    };
        //    // Add the text fragment to the page
        //    page.Paragraphs.Add(salesSpecialistText);
        //}




    }
}
