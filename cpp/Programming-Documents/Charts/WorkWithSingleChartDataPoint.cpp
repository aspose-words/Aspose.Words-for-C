#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartDataPoint.h>
#include <Model/Drawing/Charts/ChartDataPointCollection.h>
#include <Model/Drawing/Charts/ChartMarker.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Charts/MarkerSymbol.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void WorkWithSingleChartDataPoint()
{
    std::cout << "WorkWithSingleChartDataPoint example started." << std::endl;
    // ExStart:WorkWithSingleChartDataPoint
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Get first series.
    System::SharedPtr<ChartSeries> series0 = shape->get_Chart()->get_Series()->idx_get(0);
    // Get second series.
    System::SharedPtr<ChartSeries> series1 = shape->get_Chart()->get_Series()->idx_get(1);
    System::SharedPtr<ChartDataPointCollection> dataPointCollection = series0->get_DataPoints();

    // Add data point to the first and second point of the first series.
    System::SharedPtr<ChartDataPoint> dataPoint00 = dataPointCollection->Add(0);
    System::SharedPtr<ChartDataPoint> dataPoint01 = dataPointCollection->Add(1);

    // Set explosion.
    dataPoint00->set_Explosion(50);

    // Set marker symbol and size.
    dataPoint00->get_Marker()->set_Symbol(MarkerSymbol::Circle);
    dataPoint00->get_Marker()->set_Size(15);

    dataPoint01->get_Marker()->set_Symbol(MarkerSymbol::Diamond);
    dataPoint01->get_Marker()->set_Size(20);

    // Add data point to the third point of the second series.
    System::SharedPtr<ChartDataPoint> dataPoint12 = series1->get_DataPoints()->Add(2);
    dataPoint12->set_InvertIfNegative(true);
    dataPoint12->get_Marker()->set_Symbol(Aspose::Words::Drawing::Charts::MarkerSymbol::Star);
    dataPoint12->get_Marker()->set_Size(20);
    System::String outputPath = dataDir + GetOutputFilePath(u"WorkWithSingleChartDataPoint.docx");
    doc->Save(outputPath);
    // ExEnd:WorkWithSingleChartDataPoint
    std::cout << "Single line chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkWithSingleChartDataPoint example finished." << std::endl << std::endl;
}