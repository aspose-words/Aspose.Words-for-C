#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void InsertBubbleChart()
{
    std::cout << "InsertBubbleChart example started." << std::endl;
    // ExStart:InsertBubbleChart
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert Bubble chart.
    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Bubble, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Use this overload to add series to any type of Bubble charts.
    chart->get_Series()->Add(u"AW Series 1",
                             System::MakeArray<double>({0.7, 1.8, 2.6}),
                             System::MakeArray<double>({2.7, 3.2, 0.8}),
                             System::MakeArray<double>({10, 4, 8}));
    System::String outputPath = outputDataDir + u"InsertBubbleChart.docx";
    doc->Save(outputPath);
    // ExEnd:InsertBubbleChart
    std::cout << "Bubble chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertBubbleChart example finished." << std::endl << std::endl;
}