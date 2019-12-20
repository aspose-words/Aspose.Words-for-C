#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace
{
    void WorkWithChartDataLabel(System::String const &outputDataDir)
    {
        // ExStart:WorkWithChartDataLabel
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Bar, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Get first series.
        System::SharedPtr<ChartSeries> series0 = shape->get_Chart()->get_Series()->idx_get(0);
        System::SharedPtr<ChartDataLabelCollection> dataLabelCollection = series0->get_DataLabels();

        // Add data label to the first and second point of the first series.
        System::SharedPtr<ChartDataLabel> chartDataLabel00 = dataLabelCollection->Add(0);
        System::SharedPtr<ChartDataLabel> chartDataLabel01 = dataLabelCollection->Add(1);

        // Set properties.
        chartDataLabel00->set_ShowLegendKey(true);
        // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
        // Positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
        // Corresponding data point.
        chartDataLabel00->set_ShowLeaderLines(true);
        chartDataLabel00->set_ShowCategoryName(false);
        chartDataLabel00->set_ShowPercentage(false);
        chartDataLabel00->set_ShowSeriesName(true);
        chartDataLabel00->set_ShowValue(true);
        chartDataLabel00->set_Separator(u"/");
        chartDataLabel01->set_ShowValue(true);

        System::String outputPath = outputDataDir + u"WorkWithChartDataLabels.WorkWithChartDataLabel.docx";
        doc->Save(outputPath);
        // ExEnd:WorkWithChartDataLabel
        std::cout << "Simple bar chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void DefaultOptionsForDataLabels(System::String const &outputDataDir)
    {
        // ExStart:DefaultOptionsForDataLabels
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Pie, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Clear();

        System::SharedPtr<ChartSeries> series = chart->get_Series()->Add(u"Series 1", System::MakeArray<System::String>({u"Category1", u"Category2", u"Category3"}), System::MakeArray<double>({2.7, 3.2, 0.8}));

        System::SharedPtr<ChartDataLabelCollection> labels = series->get_DataLabels();
        labels->set_ShowPercentage(true);
        labels->set_ShowValue(true);
        labels->set_ShowLeaderLines(false);
        labels->set_Separator(u" - ");

        System::String outputPath = outputDataDir + u"WorkWithChartDataLabels.DefaultOptionsForDataLabels.docx";
        doc->Save(outputPath);
        // ExEnd:DefaultOptionsForDataLabels
        std::cout << "Default options for data labels of chart series created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void WorkWithChartDataLabels()
{
    std::cout << "WorkWithChartDataLabels example started." << std::endl;
    // The path to the output documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    WorkWithChartDataLabel(outputDataDir);
    DefaultOptionsForDataLabels(outputDataDir);
    std::cout << "WorkWithChartDataLabels example finished." << std::endl << std::endl;
}

