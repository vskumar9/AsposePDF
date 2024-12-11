using Aspose.Pdf;
using Aspose.Pdf.Drawing;
using Aspose.Pdf.Text;
using Rectangle = Aspose.Pdf.Drawing.Rectangle;

namespace Aspose
{
    internal class Level3
    {
        public void LW_Level3()
        {
            Document pdfDocument = new Document();
            Page page = pdfDocument.Pages.Add();
            page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            page.PageInfo.Margin.Left = 10;
            page.PageInfo.Margin.Right = 10;
            page.PageInfo.Margin.Top = 10;
            page.PageInfo.Margin.Bottom = 10;

            Page1(page);

            pdfDocument.Save(@"D:\LW_Level3.pdf");

        }


        public void Page1(Page page)
        {
            PageHeaderContent(page);

            Top5Virtues(page);

            VirtuesRecommendation(page);

            string pageEndContent = "Where the achieved score is equal to or greater than 7, Limeneal recommends to consider advanced levels of courses/workshops/trainings. For all other cases, basic to intermediate levels of courses/workshops/trainings are recommended.";
            Table TableOfContext = new Table
            {
                ColumnWidths = "550",
                Margin = new MarginInfo(10, 0, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                BackgroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE")),
            };

            page.Paragraphs.Add(TableOfContext);

            Row TableOfContexHeaderRow1 = TableOfContext.Rows.Add();

            Cell cell = TableOfContexHeaderRow1.Cells.Add();

            TextFragment textFragment1 = new TextFragment();

            TextSegment segment1 = new TextSegment(pageEndContent)
            {
                TextState = new TextState
                {
                    ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9.5f,
                    FontStyle = FontStyles.Bold,
                    LineSpacing = 5,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment1.Segments.Add(segment1);
            cell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
            //cell.Alignment = HorizontalAlignment.Center;
            cell.Margin = new MarginInfo(10, 10, 10, 3);

            cell.Paragraphs.Add(textFragment1);

        }



        static void PageHeaderContent(Page page)
        {
            Table TableOfContext = new Table
            {
                ColumnWidths = "550",
                Margin = new MarginInfo(10, 0, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                BackgroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE")),
            };

            page.Paragraphs.Add(TableOfContext);

            string text1 = "LIMENEAL TALENT RECOMMENDED INTERVENTIONS";
            string text2 = "Based on scores secured by Mr. John Smith vs. top 5 virtues required for the job role, organization may consider " +
                            "several interventions to improve the scores to create higher impact on the results. Few indicative interventions " +
                            "associated with the top 5 virtues are presented below for reference";
            

            Row TableOfContexHeaderRow1 = TableOfContext.Rows.Add();

            Cell cell = TableOfContexHeaderRow1.Cells.Add();

            TextFragment textFragment1 = new TextFragment();

            TextSegment segment1 = new TextSegment(text1)
            {
                TextState = new TextState
                {
                    ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 14,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial")
                }
            };
            

            textFragment1.Segments.Add(segment1);
            cell.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
            cell.Alignment = HorizontalAlignment.Center;
            cell.Margin = new MarginInfo(0, 15, 0, 3);

            cell.Paragraphs.Add(textFragment1);


            Row TableOfContexHeaderRow2 = TableOfContext.Rows.Add();

            Cell cell2 = TableOfContexHeaderRow2.Cells.Add();

            TextFragment textFragment2 = new TextFragment();

            TextSegment segment2 = new TextSegment(text2)
            {
                TextState = new TextState
                {
                    ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Regular,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 5
                }
            };

            textFragment2.Segments.Add(segment2);
            cell2.Margin = new MarginInfo(10, 10, 10, 3);
            cell2.Paragraphs.Add(textFragment2);




        }

        static void Top5Virtues(Page page)
        {
           Table table = new Table()
           {
               ColumnWidths = "220 330",
               Margin = new MarginInfo(10, 0, 10, 25),
               DefaultCellBorder = new BorderInfo(BorderSide.None),

           };

            page.Paragraphs.Add(table);

            Row row = table.Rows.Add();

            Cell leftCell = row.Cells.Add();
            Cell rightCell = row.Cells.Add();

            Top5VirtuesLeftSideCell(leftCell);
            Top5VirtuesRightSideCell(rightCell);


        }

        static void Top5VirtuesLeftSideCell(Cell leftCell)
        {
            Table leftTable = new Table()
            {
                ColumnWidths = "100 50 50",
                Margin = new MarginInfo(10, 0, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
            };
            leftCell.Paragraphs.Add(leftTable);

            Row leftTableRow1 = leftTable.Rows.Add();
            Cell leftTableRow1Cell1 = leftTableRow1.Cells.Add();
            Cell leftTableRow1Cell2 = leftTableRow1.Cells.Add();
            Cell leftTableRow1Cell3 = leftTableRow1.Cells.Add();
            leftTableRow1Cell1.Alignment = HorizontalAlignment.Center;
            leftTableRow1Cell2.Alignment = HorizontalAlignment.Center;
            leftTableRow1Cell3.Alignment = HorizontalAlignment.Center;

            string row1cell1s1 = "\nTop 5 Virtues\n";
            string row1cell1s2 = "required by the job role";
            string row1cell2 = "Required\nScore";
            string row1cell3 = "Achieved\nScore";



            TextFragment textFragment1 = new TextFragment();

            TextSegment segment1 = new TextSegment(row1cell1s1)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 3
                }
            };

            TextSegment segment2 = new TextSegment(row1cell1s2)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 8,
                    FontStyle = FontStyles.Regular,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 3
                }
            };


            textFragment1.Segments.Add(segment1);
            textFragment1.Segments.Add(segment2);
            //leftTableRow1Cell1.Margin = new MarginInfo(10, 10, 10, 3);
            leftTableRow1Cell1.Paragraphs.Add(textFragment1);



            TextFragment textFragment2 = new TextFragment();

            TextSegment segment3 = new TextSegment(row1cell2)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 3
                }
            };

            textFragment2.Segments.Add(segment3);
            leftTableRow1Cell2.Paragraphs.Add(textFragment2);



            TextFragment textFragment3 = new TextFragment();

            TextSegment segment4 = new TextSegment(row1cell3)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 3
                }
            };

            textFragment3.Segments.Add(segment4);
            leftTableRow1Cell3.Paragraphs.Add(textFragment3);





            Row leftTableRow2 = leftTable.Rows.Add();
            Cell leftTableRow2Cell1 = leftTableRow2.Cells.Add();
            Cell leftTableRow2Cell2 = leftTableRow2.Cells.Add();
            Cell leftTableRow2Cell3 = leftTableRow2.Cells.Add();

            leftTableRow2Cell1.Margin = new MarginInfo { Top = 5, Left = 25 };
            leftTableRow2Cell2.Alignment = HorizontalAlignment.Center;
            leftTableRow2Cell3.Alignment = HorizontalAlignment.Center;

            string row2Cell1 = "Friendliness\nTrustworthiness\nPrudence\nAmbition\nCharisma";
            string row2Cell2 = "9.0\n9.0\n9.0\n8.8\n8.5";
            string row2Cell3 = "6.3\n4.7\n4.3\n6.7\n8.5";

            TextFragment textFragment4 = new TextFragment();

            TextSegment segment5 = new TextSegment(row2Cell1)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 10
                }
            };

            textFragment4.Segments.Add(segment5);
            leftTableRow2Cell1.Paragraphs.Add(textFragment4);


            TextFragment textFragment5 = new TextFragment();

            TextSegment segment6 = new TextSegment(row2Cell2)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 10
                }
            };

            textFragment5.Segments.Add(segment6);
            leftTableRow2Cell2.Paragraphs.Add(textFragment5);

            TextFragment textFragment6 = new TextFragment();

            TextSegment segment7 = new TextSegment(row2Cell3)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 9,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 10
                }
            };

            textFragment6.Segments.Add(segment7);
            leftTableRow2Cell3.Paragraphs.Add(textFragment6);

        }
        static void Top5VirtuesRightSideCell(Cell rightCell)
        {
            Table table = new Table()
            {
                ColumnWidths = "330",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
            };

            rightCell.Paragraphs.Add(table);

            Row row1 = table.Rows.Add();
            Cell cell1 = row1.Cells.Add();
            cell1.Alignment = HorizontalAlignment.Center;

            Table HeadingTable = new Table()
            {
                ColumnWidths = "100 100",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                Margin = new MarginInfo { Left = 60 }
            };

            cell1.Paragraphs.Add(HeadingTable);

            BorderInfo borderInfo = new BorderInfo
            {
                Top = new GraphInfo
                {
                    LineWidth = 1f, // Set the top border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#d9d9d9"))
                },
                Left = new GraphInfo
                {
                    LineWidth = 1f, // Set the left border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#d9d9d9"))
                },
                Right = new GraphInfo
                {
                    LineWidth = 1f, // Set the right border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#d9d9d9"))
                },
                Bottom = new GraphInfo
                {
                    LineWidth = 1f, // Set the bottom border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#d9d9d9"))
                }
            };

            HeadingTable.Border = borderInfo;
            
            

            Row HeadingTableRow = HeadingTable.Rows.Add();
            Cell HeadingTableRowCell1 = HeadingTableRow.Cells.Add();
            Cell HeadingTableRowCell2 = HeadingTableRow.Cells.Add();

            new Level3().Top5VirtuesRightSideCellGraphHeading(HeadingTableRowCell1, "#dae3f3", "Achieved Score");
            new Level3().Top5VirtuesRightSideCellGraphHeading(HeadingTableRowCell2, "#2f5597", "Required Score");


            Row row2 = table.Rows.Add();
            Cell cell2 = row2.Cells.Add();

            new Level3().Top5VirtuesRightSideCellGraph(cell2);



        }


        public void Top5VirtuesRightSideCellGraphHeading(Cell cell, string color, string text)
        {
            Table table = new Table()
            {
                ColumnWidths = "15 85",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
            };

            cell.Paragraphs.Add(table);

            Row row = table.Rows.Add();
            Cell cell1 = row.Cells.Add();
            Cell cell2 = row.Cells.Add();

            Graph graph = new Graph(15.0, 15.0);
            cell1.Paragraphs.Add(graph);

            Rectangle rectangle = new Rectangle(15, 5, 5, 5)
            {
                GraphInfo = new GraphInfo
                {
                    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(color)),
                    FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml(color)),
                }
            };
            graph.Shapes.Add(rectangle);

            TextFragment textFragment = new TextFragment();

            TextSegment segment = new TextSegment(text)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 7,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment.Segments.Add(segment);
            //cell.BackgroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
            //cell.Margin = new MarginInfo(10, 3, 0, 3);
            cell2.Alignment = HorizontalAlignment.Center;

            cell2.Paragraphs.Add(textFragment);

        }


        public void Top5VirtuesRightSideCellGraph(Cell cell)
        {
            Table table = new Table()
            {
                ColumnWidths = "30 300",
                DefaultCellBorder = new BorderInfo(BorderSide.None),

            };
            cell.Paragraphs.Add(table);

            Row row1 = table.Rows.Add();

            Cell cell1 = row1.Cells.Add();
            Cell cell2 = row1.Cells.Add();

            cell1.Alignment = HorizontalAlignment.Center;

            string yAxis = "9.0\n7.0\n5.0\n3.0\n1.0";

            TextFragment yText = new TextFragment();

            TextSegment segment1 = new TextSegment(yAxis)
            {
                TextState = new TextState()
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#2f5597")),
                    FontSize = 10,
                    FontStyle = FontStyles.Regular,
                    LineSpacing = 5,
                    Font = FontRepository.FindFont("Arial")
                }
            };

            yText.Segments.Add(segment1);
            cell1.Paragraphs.Add(yText);


            Table table2 = new Table()
            {
                ColumnWidths = "60 60 60 60 60",
            };

            cell2.Paragraphs.Add(table2);

            BorderInfo borderInfo = new BorderInfo
            {
                Left = new GraphInfo
                {
                    LineWidth = 1f, // Set the left border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#000"))
                },
                Bottom = new GraphInfo
                {
                    LineWidth = 0.5f, // Set the bottom border width
                    Color = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#000"))
                }
            };

            table2.Border = borderInfo;

            Row table2Row1 = table2.Rows.Add();

            for (int i = 0; i < 5; i++)
            {
                Cell table2Row1Cell = table2Row1.Cells.Add();
                Top5VirtuesRightSideCellGraphRequiredScoreGraph(table2Row1Cell);

            }


        }

        public void Top5VirtuesRightSideCellGraphRequiredScoreGraph(Cell cell)
        {
            Graph graph = new Graph(60.0, 60);
            cell.Paragraphs.Add(graph);

            Rectangle rectangle = new Rectangle(15, 15, 20, 5)
            {
                GraphInfo = new GraphInfo
                {
                    Color = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#2f5597")),
                    FillColor = Aspose.Pdf.Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#2f5597")),
                }
            };
            graph.Shapes.Add(rectangle);

        }



        static void VirtuesRecommendation(Page page)
        {
            Table TableOfContext = new Table
            {
                ColumnWidths = "400",
                Margin = new MarginInfo(10, 0, 10, 10),
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                BackgroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE")),
            };

            page.Paragraphs.Add(TableOfContext);

            string text1 = "Recommended interventions for Top 5 Virtues to enhance effectiveness";
            
            Row TableOfContexHeaderRow1 = TableOfContext.Rows.Add();

            Cell cell = TableOfContexHeaderRow1.Cells.Add();

            TextFragment textFragment1 = new TextFragment();

            TextSegment segment1 = new TextSegment(text1)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment1.Segments.Add(segment1);
            cell.BackgroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#DCE1EE"));
            cell.Margin = new MarginInfo(10, 3, 0, 3);

            cell.Paragraphs.Add(textFragment1);

            //Row recommandationRow = TableOfContext.Rows.Add();

            //Cell recommandationCell = recommandationRow.Cells.Add();

            Table recommandationTable = new Table()
            {
                ColumnWidths = "80 420",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
                Margin = new MarginInfo { Top = 10, Left = 30 }
            };

            page.Paragraphs.Add(recommandationTable);

            Row row1 = recommandationTable.Rows.Add();
            Cell rowCell1 = row1.Cells.Add();
            Cell rowCell2 = row1.Cells.Add();

            //rowCell1.Alignment = HorizontalAlignment.Center;
            //rowCell1.Margin = new MarginInfo { Left = 10 };

            rowCell2.Margin = new MarginInfo { Left = 40 };

            string vir = "VIRTUES";
            string recom = "RECOMMENDATION";


            TextFragment textFragment2 = new TextFragment();

            TextSegment segment2 = new TextSegment(vir)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment2.Segments.Add(segment2);

            rowCell1.Paragraphs.Add(textFragment2);

            TextFragment textFragment3 = new TextFragment();

            TextSegment segment3 = new TextSegment(recom)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Bold,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment3.Segments.Add(segment3);

            rowCell2.Paragraphs.Add(textFragment3);

            string[] recommandationArray1 = { "Friendliness", "Trustworthiness", "Prudence", "Ambition", "Charisma" };
            string[] recommandationArray2 = { 
                                              "Social Skills Training\nParticipating in team-building exercises.\nActive listening exercises.",
                                              "Empathy related Training\nEmotional Regulation Techniques\nRapport Building Exercises",
                                              "Training programs on decision-making and problem solving skills\nPrudence and Leadership training program.\nCritical thinking and descision analysis",
                                              "Workshop on 7 Habits of Highly Effective People\nAchievement Motivation Training\nGoal-Achievement Coaching",
                                              "Public speaking training: Improve articulation and confidence.\nCommunication skills training\nBody language training" 
                                             };

            for(int i = 0; i < recommandationArray1.Length; i++)
            {
                Row row2 = recommandationTable.Rows.Add();
                new Level3().VirtuesRecommendationRows(row2, recommandationArray1[i], recommandationArray2[i]);
            }





        }

        public void VirtuesRecommendationRows(Row row, string left, string right)
        {
            Cell cell1 = row.Cells.Add();
            Cell cell2 = row.Cells.Add();

            //cell1.Margin = new MarginInfo { Top = 30 };


            TextFragment textFragment1 = new TextFragment();

            TextSegment segment1 = new TextSegment(left)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Regular,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment1.Segments.Add(segment1);
            textFragment1.VerticalAlignment = VerticalAlignment.Top;

            cell1.Paragraphs.Add(textFragment1);


            Table table = new Table()
            {
                ColumnWidths = "30 390",
                DefaultCellBorder = new BorderInfo(BorderSide.None),
            };

            cell2.Paragraphs.Add(table);

            Row tableRow = table.Rows.Add();
            Cell tableRowCell1 = tableRow.Cells.Add();
            Cell tableRowCell2 = tableRow.Cells.Add();

            tableRowCell1.Margin = new MarginInfo { Top = 20, Left = 20 };
            tableRowCell2.Margin = new MarginInfo { Top = 10, Left = 10 };

            string num = "1\n2\n3\n";

            TextFragment textFragment2 = new TextFragment();

            TextSegment segment2 = new TextSegment(num)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Bold,
                    LineSpacing = 5,
                    Font = FontRepository.FindFont("Arial")
                }
            };


            textFragment2.Segments.Add(segment2);

            tableRowCell1.Paragraphs.Add(textFragment2);


            TextFragment textFragment3 = new TextFragment();

            TextSegment segment3 = new TextSegment(right)
            {
                TextState = new TextState
                {
                    ForegroundColor = Color.FromRgb(System.Drawing.ColorTranslator.FromHtml("#344d8f")),
                    FontSize = 10,
                    FontStyle = FontStyles.Regular,
                    Font = FontRepository.FindFont("Arial"),
                    LineSpacing = 5
                }
            };


            textFragment3.Segments.Add(segment3);

            tableRowCell2.Paragraphs.Add(textFragment3);



        }




    }
}
