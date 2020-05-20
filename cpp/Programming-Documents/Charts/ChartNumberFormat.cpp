#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartNumberFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace
{
    void FormatNumberOfDataLabel(System::String const &outputDataDir)
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
        System::SharedPtr<ChartSeries> series1 = chart->get_Series()->Add(u"AW Series 1", System::MakeArray<System::String>({u"AW0", u"AW1", u"AW2"}), System::MakeArray<double>({2.5, 1.5, 3.5}));

        series1->set_HasDataLabels(true);
        series1->get_DataLabels()->set_ShowValue(true);
        series1->get_DataLabels()->idx_get(0)->get_NumberFormat()->set_FormatCode(u"\"$\"#,##0.00");
        series1->get_DataLabels()->idx_get(1)->get_NumberFormat()->set_FormatCode(u"dd/mm/yyyy");
        series1->get_DataLabels()->idx_get(2)->get_NumberFormat()->set_FormatCode(u"0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        series1->get_DataLabels()->idx_get(2)->get_NumberFormat()->set_IsLinkedToSource(true);

        System::String outputPath = outputDataDir + u"ChartNumberFormat.docx";
        doc->Save(outputPath);
        //ExEnd:FormatNumberofDataLabel
        std::cout << "Simple line chart created with formatted data lablel successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void ChartNumberFormat()
{
    std::cout << "ChartNumberFormat example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    FormatNumberOfDataLabel(outputDataDir);
    std::cout << "ChartNumberFormat example finished." << std::endl << std::endl;
}