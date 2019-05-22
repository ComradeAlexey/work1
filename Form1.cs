using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using Series = Xceed.Words.NET.Series;

namespace DiagramView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            var saveFileDialog1 = new SaveFileDialog
            {
                Filter = "Word Document (.docx ,.doc)|*.docx;*.doc|All files (*.*)|*.*",
                FilterIndex = 1
            };

            var openFileDialog = new OpenFileDialog
            {
                Filter = "Word Document (.docx ,.doc)|*.docx;*.doc|All files (*.*)|*.*",
                FilterIndex = 1,
                Multiselect = false
            };
            

                try
                {
                    if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                        return;
                    // получаем выбранный файл
                    string filename = saveFileDialog1.FileName;
                    //создаем новый word document
                    DocX document = DocX.Create(filename);

                    document.InsertParagraph("F*CK");

                    document.InsertChart(CreatePieChart());

                    document.Save();
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message, "Ошибка");
                }            
        }
        private static Xceed.Words.NET.Chart CreatePieChart()
        {
            // создаём круговую диаграмму
            PieChart pieChart = new PieChart();
            // добавляем легенду слева от диаграммы
            pieChart.AddLegend(ChartLegendPosition.Left, false);

            // создаём набор данных и добавляем на диаграмму
            pieChart.AddSeries(TestData.GetSeriesFirst());

            return pieChart;
        }
    }
    class TestData
    {
        public string name { get; set; }
        public int value { get; set; }

        private static List<TestData> GetTestDataFirst()
        {
            List<TestData> testDataFirst = new List<TestData>
            {
                new TestData() { name = "1", value = 1 },
                new TestData() { name = "10", value = 10 },
                new TestData() { name = "5", value = 5 },
                new TestData() { name = "8", value = 8 },
                new TestData() { name = "5", value = 5 }
            };

            return testDataFirst;
        }

        private static List<TestData> GetTestDataSecond()
        {
            List<TestData> testDataSecond = new List<TestData>
            {
                new TestData() { name = "12", value = 12 },
                new TestData() { name = "3", value = 3 },
                new TestData() { name = "8", value = 8 },
                new TestData() { name = "15", value = 15 },
                new TestData() { name = "1", value = 1 }
            };

            return testDataSecond;
        }

        public static Series GetSeriesFirst()
        {
            // создаём набор данных
            Series seriesFirst = new Series("First");
            // заполняем данными
            seriesFirst.Bind(GetTestDataFirst(), "name", "value");
            return seriesFirst;
        }

        public static Series GetSeriesSecond()
        {
            // создаём набор данных
            Series seriesSecond = new Series("Second");
            // заполняем данными
            seriesSecond.Bind(GetTestDataSecond(), "name", "value");
            return seriesSecond;
        }
    }
}
