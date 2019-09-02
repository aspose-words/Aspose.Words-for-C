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

void InsertScatterChart()
{
    std::cout << "InsertScatterChart example started." << std::endl;
    // ExStart:InsertScatterChart
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert Scatter chart.
    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Scatter, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Use this overload to add series to any type of Scatter charts.
    chart->get_Series()->Add(u"AW Series 1", System::MakeArray<double>({0.7, 1.8, 2.6}), System::MakeArray<double>({2.7, 3.2, 0.8}));

    System::String outputPath = outputDataDir + u"InsertScatterChart.docx";
    doc->Save(outputPath);
    // ExEnd:InsertScatterChart
    std::cout << "Scatter chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertScatterChart example finished." << std::endl << std::endl;
}