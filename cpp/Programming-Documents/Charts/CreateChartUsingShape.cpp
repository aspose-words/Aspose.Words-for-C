#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartLegend.h>
#include <Model/Drawing/Charts/LegendPosition.h>
#include <Model/Drawing/Charts/ChartTitle.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

void CreateChartUsingShape()
{
    std::cout << "CreateChartUsingShape example started." << std::endl;
    // ExStart:CreateChartUsingShape
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithCharts();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

    System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);
    System::SharedPtr<Chart> chart = shape->get_Chart();

    // Determines whether the title shall be shown for this chart. Default is true.
    chart->get_Title()->set_Show(true);
    // Setting chart Title.
    chart->get_Title()->set_Text(u"Sample Line Chart Title");
    // Determines whether other chart elements shall be allowed to overlap title.
    chart->get_Title()->set_Overlay(false);

    // Please note if null or empty value is specified as title text, auto generated title will be shown.

    // Determines how legend shall be shown for this chart.
    chart->get_Legend()->set_Position(LegendPosition::Left);
    chart->get_Legend()->set_Overlay(true);
    System::String outputPath = dataDir + GetOutputFilePath(u"CreateChartUsingShape.docx");
    doc->Save(outputPath);
    // ExEnd:CreateChartUsingShape
    std::cout << "Simple line chart created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CreateChartUsingShape example finished." << std::endl << std::endl;
}