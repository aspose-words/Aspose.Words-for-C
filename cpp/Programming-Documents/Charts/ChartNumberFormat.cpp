#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartDataLabel.h>
#include <Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Model/Drawing/Charts/ChartNumberFormat.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartTitle.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace
{
    void FormatNumberOfDataLabel(System::String const &dataDir)
    {
        //ExStart:FormatNumberofDataLabel
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Add chart with default data.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Title()->set_Text(u"Data Labels With Different Number Format");

        // Delete default generated series.
        chart->get_Series()->Clear();

        // Add new series
        System::SharedPtr<ChartSeries> series0 = chart->get_Series()->Add(u"AW Series 0", System::MakeArray<System::String>({u"AW0", u"AW1", u"AW2"}), System::MakeArray<double>({2.5, 1.5, 3.5}));

        // Add DataLabel to the first point of the first series.
        System::SharedPtr<ChartDataLabel> chartDataLabel0 = series0->get_DataLabels()->Add(0);
        chartDataLabel0->set_ShowValue(true);

        // Set currency format code.
        chartDataLabel0->get_NumberFormat()->set_FormatCode(u"\"$\"#,##0.00");

        System::SharedPtr<ChartDataLabel> chartDataLabel1 = series0->get_DataLabels()->Add(1);
        chartDataLabel1->set_ShowValue(true);

        // Set date format code.
        chartDataLabel1->get_NumberFormat()->set_FormatCode(u"d/mm/yyyy");

        System::SharedPtr<ChartDataLabel> chartDataLabel2 = series0->get_DataLabels()->Add(2);
        chartDataLabel2->set_ShowValue(true);

        // Set percentage format code.
        chartDataLabel2->get_NumberFormat()->set_FormatCode(u"0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        chartDataLabel2->get_NumberFormat()->set_IsLinkedToSource(true);

        System::String outputPath = dataDir + GetOutputFilePath(u"ChartNumberFormat.docx");
        doc->Save(outputPath);
        //ExEnd:FormatNumberofDataLabel
        std::cout << "Simple line chart created with formatted data lablel successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void ChartNumberFormat()
{
    std::cout << "ChartNumberFormat example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
    FormatNumberOfDataLabel(dataDir);
    std::cout << "ChartNumberFormat example finished." << std::endl << std::endl;
}