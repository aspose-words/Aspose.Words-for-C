#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPoint.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPointCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartMarker.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/MarkerSymbol.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void WorkWithSingleChartDataPoint()
{
    std::cout << "WorkWithSingleChartDataPoint example started." << std::endl;
    // ExStart:WorkWithSingleChartDataPoint
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Get first series.
    System::SharedPtr<ChartSeries> series0 = shape->get_Chart()->get_Series()->idx_get(0);
    // Get second series.
    System::SharedPtr<ChartSeries> series1 = shape->get_Chart()->get_Series()->idx_get(1);
    System::SharedPtr<ChartDataPointCollection> dataPointCollection = series0->get_DataPoints();

    // Set explosion.
    dataPointCollection->idx_get(0)->set_Explosion(50);

    // Set marker symbol and size.
    dataPointCollection->idx_get(0)->get_Marker()->set_Symbol(MarkerSymbol::Circle);
    dataPointCollection->idx_get(0)->get_Marker()->set_Size(15);

    dataPointCollection->idx_get(1)->get_Marker()->set_Symbol(MarkerSymbol::Diamond);
    dataPointCollection->idx_get(1)->get_Marker()->set_Size(20);

    // Add data point to the third point of the second series.
    series1->get_DataPoints()->idx_get(2)->set_InvertIfNegative(true);
    series1->get_DataPoints()->idx_get(2)->get_Marker()->set_Symbol(MarkerSymbol::Star);
    series1->get_DataPoints()->idx_get(2)->get_Marker()->set_Size(20);
    System::String outputPath = outputDataDir + u"WorkWithSingleChartDataPoint.docx";
    doc->Save(outputPath);
    // ExEnd:WorkWithSingleChartDataPoint
    std::cout << "Single line chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "WorkWithSingleChartDataPoint example finished." << std::endl << std::endl;
}