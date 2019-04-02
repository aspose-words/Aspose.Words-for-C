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

void InsertAreaChart()
{
    std::cout << "InsertAreaChart example started." << std::endl;
    // ExStart:InsertAreaChart
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    // Insert Area chart.
    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Area, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Use this overload to add series to any type of Area, Radar and Stock charts.
    chart->get_Series()->Add(u"AW Series 1",
                             System::MakeArray<System::DateTime>({System::DateTime(2002, 5, 1), System::DateTime(2002, 6, 1), System::DateTime(2002, 7, 1), System::DateTime(2002, 8, 1), System::DateTime(2002, 9, 1)}),
                             System::MakeArray<double>({32, 32, 28, 12, 15}));
    System::String outputPath = dataDir + GetOutputFilePath(u"InsertAreaChart.docx");
    doc->Save(outputPath);
    // ExEnd:InsertAreaChart
    std::cout << "Scatter chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertAreaChart example finished." << std::endl << std::endl;
}