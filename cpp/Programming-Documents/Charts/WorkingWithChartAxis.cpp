#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Drawing/Charts/AxisBound.h>
#include <Model/Drawing/Charts/AxisCategoryType.h>
#include <Model/Drawing/Charts/AxisCrosses.h>
#include <Model/Drawing/Charts/AxisDisplayUnit.h>
#include <Model/Drawing/Charts/AxisScaling.h>
#include <Model/Drawing/Charts/AxisTickLabelPosition.h>
#include <Model/Drawing/Charts/AxisTickMark.h>
#include <Model/Drawing/Charts/Chart.h>
#include <Model/Drawing/Charts/ChartAxis.h>
#include <Model/Drawing/Charts/ChartNumberFormat.h>
#include <Model/Drawing/Charts/ChartSeries.h>
#include <Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Model/Drawing/Charts/ChartType.h>
#include <Model/Drawing/Shape.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace
{
    void DefineXYAxisProperties(System::String const &outputDataDir)
    {
        //ExStart:DefineXYAxisProperties
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Area, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        System::ArrayPtr<System::DateTime> dates = System::MakeArray<System::DateTime>({System::DateTime(2002, 1, 1),
                                                                                        System::DateTime(2002, 6, 1),
                                                                                        System::DateTime(2002, 7, 1),
                                                                                        System::DateTime(2002, 8, 1),
                                                                                        System::DateTime(2002, 9, 1)});
        chart->get_Series()->Add(u"AW Series 1", dates, System::MakeArray<double>({640, 320, 280, 120, 150}));

        System::SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        System::SharedPtr<ChartAxis> yAxis = chart->get_AxisY();

        // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
        xAxis->set_CategoryType(AxisCategoryType::Category);

        // Define X axis properties.
        xAxis->set_Crosses(AxisCrosses::Custom);
        xAxis->set_CrossesAt(3);
        // measured in display units of the Y axis (hundreds)
        xAxis->set_ReverseOrder(true);
        xAxis->set_MajorTickMark(AxisTickMark::Cross);
        xAxis->set_MinorTickMark(AxisTickMark::Outside);
        xAxis->set_TickLabelOffset(200);

        // Define Y axis properties.
        yAxis->set_TickLabelPosition(AxisTickLabelPosition::High);
        yAxis->set_MajorUnit(100);
        yAxis->set_MinorUnit(50);
        yAxis->get_DisplayUnit()->set_Unit(AxisBuiltInUnit::Hundreds);
        yAxis->get_Scaling()->set_Minimum(System::MakeObject<AxisBound>(100));
        yAxis->get_Scaling()->set_Maximum(System::MakeObject<AxisBound>(700));

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.DefineXYAxisProperties.docx";
        doc->Save(outputPath);
        //ExEnd:DefineXYAxisProperties
        std::cout << "Properties of X and Y axis are set successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetDateTimeValuesToAxis(System::String const &outputDataDir)
    {
        // ExStart:SetDateTimeValuesToAxis
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        System::ArrayPtr<System::DateTime> dates = System::MakeArray<System::DateTime>({System::DateTime(2017, 11, 6),
                                                                                        System::DateTime(2017, 11, 9),
                                                                                        System::DateTime(2017, 11, 15),
                                                                                        System::DateTime(2017, 11, 21),
                                                                                        System::DateTime(2017, 11, 25),
                                                                                        System::DateTime(2017, 11, 29)});
        chart->get_Series()->Add(u"AW Series 1", dates, System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2, 5.3}));

        // Set X axis bounds.
        System::SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        xAxis->get_Scaling()->set_Minimum(System::MakeObject<AxisBound>((System::DateTime(2017, 11, 5)).ToOADate()));
        xAxis->get_Scaling()->set_Maximum(System::MakeObject<AxisBound>((System::DateTime(2017, 12, 3)).ToOADate()));

        // Set major units to a week and minor units to a day.
        xAxis->set_MajorUnit(7);
        xAxis->set_MinorUnit(1);
        xAxis->set_MajorTickMark(AxisTickMark::Cross);
        xAxis->set_MinorTickMark(AxisTickMark::Outside);

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.SetDateTimeValuesToAxis.docx";
        doc->Save(outputPath);
        // ExEnd:SetDateTimeValuesToAxis
        std::cout << "DateTime values are set for chart axis successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetNumberFormatForAxis(System::String const &outputDataDir)
    {
        // ExStart:SetNumberFormatForAxis
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        chart->get_Series()->Add(u"AW Series 1",
                                 System::MakeArray<System::String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 System::MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

        // Set number format.
        chart->get_AxisY()->get_NumberFormat()->set_FormatCode(u"#,##0");

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.SetNumberFormatForAxis.docx";
        doc->Save(outputPath);
        // ExEnd:SetNumberFormatForAxis
        std::cout << "Set number format for axis successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetBoundsOfAxis(System::String const &outputDataDir)
    {
        // ExStart:SetboundsOfAxis
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        chart->get_Series()->Add(u"AW Series 1",
                                 System::MakeArray<System::String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        chart->get_AxisY()->get_Scaling()->set_Minimum(System::MakeObject<AxisBound>(0));
        chart->get_AxisY()->get_Scaling()->set_Maximum(System::MakeObject<AxisBound>(6));

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.SetBoundsOfAxis.docx";
        doc->Save(outputPath);
        // ExEnd:SetboundsOfAxis
        std::cout << "Set Bounds of chart axis successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetIntervalUnitBetweenLabelsOnAxis(System::String const &outputDataDir)
    {
        // ExStart:SetIntervalUnitBetweenLabelsOnAxis
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        chart->get_Series()->Add(u"AW Series 1",
                                 System::MakeArray<System::String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        chart->get_AxisX()->set_TickLabelSpacing(2);

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.SetIntervalUnitBetweenLabelsOnAxis.docx";
        doc->Save(outputPath);
        // ExEnd:SetIntervalUnitBetweenLabelsOnAxis
        std::cout << "Set interval unit between labels on an axis successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void HideChartAxis(System::String const &outputDataDir)
    {
        // ExStart:HideChartAxis
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Insert chart.
        System::SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        System::SharedPtr<Chart> chart = shape->get_Chart();

        // Clear demo data.
        chart->get_Series()->Clear();

        // Fill data.
        chart->get_Series()->Add(u"AW Series 1",
                                 System::MakeArray<System::String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        // Hide the Y axis.
        chart->get_AxisY()->set_Hidden(true);

        System::String outputPath = outputDataDir + u"WorkingWithChartAxis.HideChartAxis.docx";
        doc->Save(outputPath);
        // ExEnd:HideChartAxis
        std::cout << "Y Axis of chart has hidden successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void TickMultiLineLabelAlignment(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:TickMultiLineLabelAlignment
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.docx");
        System::SharedPtr<Shape> shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));
        System::SharedPtr<ChartAxis> axis = shape->get_Chart()->get_AxisX();

        //This property has effect only for multi-line labels.
        axis->set_TickLabelAlignment(ParagraphAlignment::Right);

        doc->Save(outputDataDir + u"WorkingWithChartAxis.TickMultiLineLabelAlignment.docx");
        // ExEnd:TickMultiLineLabelAlignment
    }
}

void WorkingWithChartAxis()
{
    std::cout << "WorkingWithChartAxis example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithCharts();
    System::String outputDataDir = GetOutputDataDir_WorkingWithCharts();
    DefineXYAxisProperties(outputDataDir);
    SetDateTimeValuesToAxis(outputDataDir);
    SetNumberFormatForAxis(outputDataDir);
    SetBoundsOfAxis(outputDataDir);
    SetIntervalUnitBetweenLabelsOnAxis(outputDataDir);
    HideChartAxis(outputDataDir);
    TickMultiLineLabelAlignment(inputDataDir, outputDataDir);
    std::cout << "WorkingWithChartAxis example finished." << std::endl << std::endl;
}