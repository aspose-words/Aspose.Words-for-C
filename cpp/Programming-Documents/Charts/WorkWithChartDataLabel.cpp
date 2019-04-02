#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartDataLabel.h>
#include <Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void WorkWithChartDataLabel()
{
    std::cout << "WorkWithChartDataLabel example started." << std::endl;
    // ExStart:WorkWithChartDataLabel
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
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

    System::String outputPath = dataDir + GetOutputFilePath(u"WorkWithChartDataLabel.docx");
    doc->Save(outputPath);
    // ExEnd:WorkWithChartDataLabel
    std::cout << "Simple bar chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkWithChartDataLabel example finished." << std::endl << std::endl;
}