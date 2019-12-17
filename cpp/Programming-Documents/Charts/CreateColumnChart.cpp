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

namespace
{
    void InsertSimpleColumnChart(System::String const &outputDataDir)
    {
        // ExStart:InsertSimpleColumnChart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Add chart with default data. You can specify different chart types and sizes.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        // Chart property of Shape contains all chart related options.
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // ExStart:ChartSeriesCollection 
        // Get chart series collection.
        System::SharedPtr<ChartSeriesCollection> seriesColl = chart->get_Series();
        // Check series count.
        std::cout << seriesColl->get_Count() << std::endl;
        // ExEnd:ChartSeriesCollection 

        // Delete default generated series.
        seriesColl->Clear();

        // Create category names array, in this example we have two categories.
        System::ArrayPtr<System::String> categories = System::MakeArray<System::String>({u"AW Category 1", u"AW Category 2"});

        // Adding new series. Please note, data arrays must not be empty and arrays must be the same size.
        seriesColl->Add(u"AW Series 1", categories, System::MakeArray<double>({1, 2}));
        seriesColl->Add(u"AW Series 2", categories, System::MakeArray<double>({3, 4}));
        seriesColl->Add(u"AW Series 3", categories, System::MakeArray<double>({5, 6}));
        seriesColl->Add(u"AW Series 4", categories, System::MakeArray<double>({7, 8}));
        seriesColl->Add(u"AW Series 5", categories, System::MakeArray<double>({9, 10}));

        System::String outputPath = outputDataDir + u"CreateColumnChart.InsertSimpleColumnChart.doc";
        doc->Save(outputPath);
        // ExEnd:InsertSimpleColumnChart
        std::cout << "Simple column chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void InsertColumnChart(System::String const &outputDataDir)
    {
        // ExStart:InsertColumnChart
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert Column chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Use this overload to add series to any type of Bar, Column, Line and Surface charts.
        chart->get_Series()->Add(u"AW Series 1", System::MakeArray<System::String>({u"AW Category 1", u"AW Category 2"}), System::MakeArray<double>({1, 2}));

        System::String outputPath = outputDataDir + u"CreateColumnChart.InsertColumnChart.doc";
        doc->Save(outputPath);
        // ExEnd:InsertColumnChart
        std::cout << "Column chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void CreateColumnChart()
{
    std::cout << "CreateColumnChart example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    InsertSimpleColumnChart(outputDataDir);
    InsertColumnChart(outputDataDir);
    std::cout << "CreateColumnChart example finished." << std::endl << std::endl;
}