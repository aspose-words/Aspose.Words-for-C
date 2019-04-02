#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartMarker.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Charts/MarkerSymbol.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void WorkWithSingleChartSeries()
{
    std::cout << "WorkWithSingleChartSeries example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();
    // ExStart:WorkWithSingleChartSeries
    // Get first series.
    System::SharedPtr<ChartSeries> series0 = shape->get_Chart()->get_Series()->idx_get(0);

    // Get second series.
    System::SharedPtr<ChartSeries> series1 = shape->get_Chart()->get_Series()->idx_get(1);

    // Change first series name.
    series0->set_Name(u"My Name1");

    // Change second series name.
    series1->set_Name(u"My Name2");

    // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
    series0->set_Smooth(true);
    series1->set_Smooth(true);
    // ExEnd:WorkWithSingleChartSeries
    // ExStart:ChartDataPoint 
    // Specifies whether by default the parent element shall inverts its colors if the value is negative.
    series0->set_InvertIfNegative(true);

    // Set default marker symbol and size.
    series0->get_Marker()->set_Symbol(MarkerSymbol::Circle);
    series0->get_Marker()->set_Size(15);

    series1->get_Marker()->set_Symbol(MarkerSymbol::Star);
    series1->get_Marker()->set_Size(10);
    // ExEnd:ChartDataPoint 
    System::String outputPath = dataDir + GetOutputFilePath(u"WorkWithSingleChartSeries.docx");
    doc->Save(outputPath);
    std::cout << "Chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkWithSingleChartSeries example finished." << std::endl << std::endl;
}