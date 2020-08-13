// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCharts.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/primitive_types.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/console.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/MarkerSymbol.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/LegendPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartNumberFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartMarker.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartLegend.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPointCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPoint.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxisType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxis.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTimeUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickMark.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickLabelPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisScaling.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisScaleType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisDisplayUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCrosses.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCategoryType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBuiltInUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBound.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
namespace ApiExamples {

void ExCharts::ApplyDataLabels(SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int labelsCount, String numberFormat, String separator)
{
    for (int i = 0; i < labelsCount; i++)
    {
        series->set_HasDataLabels(true);
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsVisible());

        // Edit the appearance of the new data label
        series->get_DataLabels()->idx_get(i)->set_ShowCategoryName(true);
        series->get_DataLabels()->idx_get(i)->set_ShowSeriesName(true);
        series->get_DataLabels()->idx_get(i)->set_ShowValue(true);
        series->get_DataLabels()->idx_get(i)->set_ShowLeaderLines(true);
        series->get_DataLabels()->idx_get(i)->set_ShowLegendKey(true);
        series->get_DataLabels()->idx_get(i)->set_ShowPercentage(false);
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_ShowDataLabelsRange());

        // Apply number format and separator
        series->get_DataLabels()->idx_get(i)->get_NumberFormat()->set_FormatCode(numberFormat);
        series->get_DataLabels()->idx_get(i)->set_Separator(separator);

        // The label automatically becomes visible
        ASSERT_TRUE(series->get_DataLabels()->idx_get(i)->get_IsVisible());
    }
}

void ExCharts::ApplyDataPoints(SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int dataPointsCount, Aspose::Words::Drawing::Charts::MarkerSymbol markerSymbol, int dataPointSize)
{
    for (int i = 0; i < dataPointsCount; i++)
    {
        SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> point = series->get_DataPoints()->Add(i);
        point->get_Marker()->set_Symbol(markerSymbol);
        point->get_Marker()->set_Size(dataPointSize);

        ASSERT_EQ(i, point->get_Index());
    }
}

SharedPtr<Aspose::Words::Drawing::Charts::Chart> ExCharts::AppendChart(SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::Drawing::Charts::ChartType chartType, double width, double height)
{
    SharedPtr<Shape> chartShape = builder->InsertChart(chartType, width, height);
    SharedPtr<Chart> chart = chartShape->get_Chart();
    chart->get_Series()->Clear();

    EXPECT_EQ(0, chart->get_Series()->get_Count());

    return chart;
}

namespace gtest_test
{

class ExCharts : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExCharts> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExCharts>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExCharts> ExCharts::s_instance;

} // namespace gtest_test

void ExCharts::ChartTitle()
{
    //ExStart
    //ExFor:Chart
    //ExFor:Chart.Title
    //ExFor:ChartTitle
    //ExFor:ChartTitle.Overlay
    //ExFor:ChartTitle.Show
    //ExFor:ChartTitle.Text
    //ExSummary:Shows how to insert a chart and change its title.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Use a document builder to insert a bar chart
    SharedPtr<Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bar, 400, 300);

    // Get the chart object from the containing shape
    SharedPtr<Chart> chart = chartShape->get_Chart();

    // Set the title text, which appears at the top center of the chart and modify its appearance
    SharedPtr<Aspose::Words::Drawing::Charts::ChartTitle> title = chart->get_Title();
    title->set_Text(u"MyChart");
    title->set_Overlay(true);
    title->set_Show(true);

    doc->Save(ArtifactsDir + u"Charts.ChartTitle.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartTitle.docx");
    chartShape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::NonPrimitive, chartShape->get_ShapeType());
    ASSERT_TRUE(chartShape->get_HasChart());

    title = chartShape->get_Chart()->get_Title();

    ASSERT_EQ(u"MyChart", title->get_Text());
    ASSERT_TRUE(title->get_Overlay());
    ASSERT_TRUE(title->get_Show());
}

namespace gtest_test
{

TEST_F(ExCharts, ChartTitle)
{
    s_instance->ChartTitle();
}

} // namespace gtest_test

void ExCharts::DefineNumberFormatForDataLabels()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.NumberFormat
    //ExFor:ChartNumberFormat.FormatCode
    //ExSummary:Shows how to set number format for the data labels of the entire series.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add chart with default data
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    // Delete default generated series
    shape->get_Chart()->get_Series()->Clear();

    SharedPtr<ChartSeries> series = shape->get_Chart()->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel"}), MakeArray<double>({2.5, 1.5, 3.5}));

    SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    // Display chart values in the data labels, by default it is false
    dataLabels->set_ShowValue(true);
    // Set currency format for the data labels of the entire series
    dataLabels->get_NumberFormat()->set_FormatCode(u"\"$\"#,##0.00");

    doc->Save(ArtifactsDir + u"Charts.DefineNumberFormatForDataLabels.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.DefineNumberFormatForDataLabels.docx");
    series = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

    ASSERT_EQ(String::Empty, series->get_DataLabels()->get_NumberFormat()->get_FormatCode());
}

namespace gtest_test
{

TEST_F(ExCharts, DefineNumberFormatForDataLabels)
{
    s_instance->DefineNumberFormatForDataLabels();
}

} // namespace gtest_test

void ExCharts::DataArraysWrongSize()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add chart with default data
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    seriesColl->Clear();

    // Create category names array, second category will be null
    ArrayPtr<String> categories = MakeArray<String>({u"Cat1", nullptr, u"Cat3", u"Cat4", u"Cat5", nullptr});

    // Adding new series with empty (double.NaN) values
    seriesColl->Add(u"AW Series 1", categories, MakeArray<double>({1, 2, std::numeric_limits<double>::quiet_NaN(), 4, 5, 6}));
    seriesColl->Add(u"AW Series 2", categories, MakeArray<double>({2, 3, std::numeric_limits<double>::quiet_NaN(), 5, 6, 7}));

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_0 = [&seriesColl, &categories]()
    {
        seriesColl->Add(u"AW Series 3", categories, MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), 4, 5, std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN()}));
    };

    ASSERT_THROW(_local_func_0(), System::ArgumentException);

    // CSPORTCPP: [WARNING] Using local variables. Make sure that local function ptr does not leave the current scope.
    std::function<void()> _local_func_1 = [&seriesColl, &categories]()
    {
        seriesColl->Add(u"AW Series 4", categories, MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN()}));
    };

    ASSERT_THROW(_local_func_1(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExCharts, DataArraysWrongSize)
{
    s_instance->DataArraysWrongSize();
}

} // namespace gtest_test

void ExCharts::EmptyValuesInChartData()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add chart with default data
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    seriesColl->Clear();

    // Create category names array, second category will be null
    ArrayPtr<String> categories = MakeArray<String>({u"Cat1", nullptr, u"Cat3", u"Cat4", u"Cat5", nullptr});

    // Adding new series with empty (double.NaN) values
    seriesColl->Add(u"AW Series 1", categories, MakeArray<double>({1, 2, std::numeric_limits<double>::quiet_NaN(), 4, 5, 6}));
    seriesColl->Add(u"AW Series 2", categories, MakeArray<double>({2, 3, std::numeric_limits<double>::quiet_NaN(), 5, 6, 7}));
    seriesColl->Add(u"AW Series 3", categories, MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), 4, 5, std::numeric_limits<double>::quiet_NaN(), 7, 8}));
    seriesColl->Add(u"AW Series 4", categories, MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), 9}));

    doc->Save(ArtifactsDir + u"Charts.EmptyValuesInChartData.docx");
}

namespace gtest_test
{

TEST_F(ExCharts, EmptyValuesInChartData)
{
    s_instance->EmptyValuesInChartData();
}

} // namespace gtest_test

void ExCharts::AxisProperties()
{
    //ExStart
    //ExFor:ChartAxis
    //ExFor:ChartAxis.CategoryType
    //ExFor:ChartAxis.Crosses
    //ExFor:ChartAxis.ReverseOrder
    //ExFor:ChartAxis.MajorTickMark
    //ExFor:ChartAxis.MinorTickMark
    //ExFor:ChartAxis.MajorUnit
    //ExFor:ChartAxis.MinorUnit
    //ExFor:ChartAxis.TickLabelOffset
    //ExFor:ChartAxis.TickLabelPosition
    //ExFor:ChartAxis.TickLabelSpacingIsAuto
    //ExFor:ChartAxis.TickMarkSpacing
    //ExFor:Charts.AxisCategoryType
    //ExFor:Charts.AxisCrosses
    //ExFor:Charts.Chart.AxisX
    //ExFor:Charts.Chart.AxisY
    //ExFor:Charts.Chart.AxisZ
    //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart.
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Clear demo data.
    chart->get_Series()->Clear();
    chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({640, 320, 280, 120, 150}));

    // Get chart axes
    SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
    SharedPtr<ChartAxis> yAxis = chart->get_AxisY();

    // For 2D charts like the one we made, the Z axis is null
    ASSERT_TRUE(chart->get_AxisZ() == nullptr);

    // Set X-axis options
    xAxis->set_CategoryType(Aspose::Words::Drawing::Charts::AxisCategoryType::Category);
    xAxis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Minimum);
    xAxis->set_ReverseOrder(false);
    xAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    xAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    xAxis->set_MajorUnit(10.0);
    xAxis->set_MinorUnit(15.0);
    xAxis->set_TickLabelOffset(50);
    xAxis->set_TickLabelPosition(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::Low);
    xAxis->set_TickLabelSpacingIsAuto(false);
    xAxis->set_TickMarkSpacing(1);

    // Set Y-axis options
    yAxis->set_CategoryType(Aspose::Words::Drawing::Charts::AxisCategoryType::Automatic);
    yAxis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Maximum);
    yAxis->set_ReverseOrder(true);
    yAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    yAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    yAxis->set_MajorUnit(100.0);
    yAxis->set_MinorUnit(20.0);
    yAxis->set_TickLabelPosition(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::NextToAxis);

    doc->Save(ArtifactsDir + u"Charts.AxisProperties.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisProperties.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCategoryType::Category, chart->get_AxisX()->get_CategoryType());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCrosses::Minimum, chart->get_AxisX()->get_Crosses());
    ASSERT_FALSE(chart->get_AxisX()->get_ReverseOrder());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, chart->get_AxisX()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisX()->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(0.5, chart->get_AxisX()->get_MinorUnit());
    ASSERT_EQ(50, chart->get_AxisX()->get_TickLabelOffset());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::Low, chart->get_AxisX()->get_TickLabelPosition());
    ASSERT_FALSE(chart->get_AxisX()->get_TickLabelSpacingIsAuto());
    ASSERT_EQ(1, chart->get_AxisX()->get_TickMarkSpacing());

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCategoryType::Category, chart->get_AxisY()->get_CategoryType());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCrosses::Maximum, chart->get_AxisY()->get_Crosses());
    ASSERT_TRUE(chart->get_AxisY()->get_ReverseOrder());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, chart->get_AxisY()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisY()->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(20.0, chart->get_AxisY()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::NextToAxis, chart->get_AxisY()->get_TickLabelPosition());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisProperties)
{
    s_instance->AxisProperties();
}

} // namespace gtest_test

void ExCharts::DateTimeValues()
{
    //ExStart
    //ExFor:AxisBound
    //ExFor:AxisBound.#ctor(Double)
    //ExFor:AxisBound.#ctor(DateTime)
    //ExFor:AxisScaling.Minimum
    //ExFor:AxisScaling.Maximum
    //ExFor:ChartAxis.Scaling
    //ExFor:Charts.AxisTickMark
    //ExFor:Charts.AxisTickLabelPosition
    //ExFor:Charts.AxisTimeUnit
    //ExFor:Charts.ChartAxis.BaseTimeUnit
    //ExSummary:Shows how to insert chart with date/time values.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Clear demo data
    chart->get_Series()->Clear();

    // Fill data
    chart->get_Series()->Add(u"Aspose Test Series", MakeArray<System::DateTime>({System::DateTime(2017, 11, 6), System::DateTime(2017, 11, 9), System::DateTime(2017, 11, 15), System::DateTime(2017, 11, 21), System::DateTime(2017, 11, 25), System::DateTime(2017, 11, 29)}), MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2, 5.3}));

    SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
    SharedPtr<ChartAxis> yAxis = chart->get_AxisY();

    // Set X axis bounds
    xAxis->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 11, 5).ToOADate()));
    xAxis->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 12, 3)));

    // Set major units to a week and minor units to a day
    xAxis->set_BaseTimeUnit(Aspose::Words::Drawing::Charts::AxisTimeUnit::Days);
    xAxis->set_MajorUnit(7.0);
    xAxis->set_MinorUnit(1.0);
    xAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    xAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Outside);

    // Define Y axis properties
    yAxis->set_TickLabelPosition(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::High);
    yAxis->set_MajorUnit(100.0);
    yAxis->set_MinorUnit(50.0);
    yAxis->get_DisplayUnit()->set_Unit(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Hundreds);
    yAxis->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(100.0));
    yAxis->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(700.0));

    doc->Save(ArtifactsDir + u"Charts.DateTimeValues.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.DateTimeValues.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 11, 5).ToOADate()), chart->get_AxisX()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 12, 3)), chart->get_AxisX()->get_Scaling()->get_Maximum());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTimeUnit::Days, chart->get_AxisX()->get_BaseTimeUnit());
    ASPOSE_ASSERT_EQ(7.0, chart->get_AxisX()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisX()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Outside, chart->get_AxisX()->get_MinorTickMark());

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::High, chart->get_AxisY()->get_TickLabelPosition());
    ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(50.0, chart->get_AxisY()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Hundreds, chart->get_AxisY()->get_DisplayUnit()->get_Unit());
    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(100.0), chart->get_AxisY()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(700.0), chart->get_AxisY()->get_Scaling()->get_Maximum());
}

namespace gtest_test
{

TEST_F(ExCharts, DateTimeValues)
{
    s_instance->DateTimeValues();
}

} // namespace gtest_test

void ExCharts::HideChartAxis()
{
    //ExStart
    //ExFor:ChartAxis.Hidden
    //ExSummary:Shows how to hide chart axes.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Hide both the X and Y axes
    chart->get_AxisX()->set_Hidden(true);
    chart->get_AxisY()->set_Hidden(true);

    // Clear demo data
    chart->get_Series()->Clear();
    chart->get_Series()->Add(u"AW Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}), MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

    doc->Save(ArtifactsDir + u"Charts.HideChartAxis.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.HideChartAxis.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASSERT_TRUE(chart->get_AxisX()->get_Hidden());
    ASSERT_TRUE(chart->get_AxisY()->get_Hidden());
}

namespace gtest_test
{

TEST_F(ExCharts, HideChartAxis)
{
    s_instance->HideChartAxis();
}

} // namespace gtest_test

void ExCharts::SetNumberFormatToChartAxis()
{
    //ExStart
    //ExFor:ChartAxis.NumberFormat
    //ExFor:Charts.ChartNumberFormat
    //ExFor:ChartNumberFormat.FormatCode
    //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
    //ExSummary:Shows how to set formatting for chart values.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Clear demo data and replace it with a new custom chart series
    chart->get_Series()->Clear();
    chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0}));

    // Set number format
    chart->get_AxisY()->get_NumberFormat()->set_FormatCode(u"#,##0");

    // Set this to override the above value and draw the number format from the source cell
    ASSERT_FALSE(chart->get_AxisY()->get_NumberFormat()->get_IsLinkedToSource());

    doc->Save(ArtifactsDir + u"Charts.SetNumberFormatToChartAxis.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.SetNumberFormatToChartAxis.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASSERT_EQ(u"#,##0", chart->get_AxisY()->get_NumberFormat()->get_FormatCode());
}

namespace gtest_test
{

TEST_F(ExCharts, SetNumberFormatToChartAxis)
{
    s_instance->SetNumberFormatToChartAxis();
}

} // namespace gtest_test

void ExCharts::TestDisplayChartsWithConversion(Aspose::Words::Drawing::Charts::ChartType chartType)
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart
    SharedPtr<Shape> shape = builder->InsertChart(chartType, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Clear demo data
    chart->get_Series()->Clear();

    chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

    doc->Save(ArtifactsDir + u"Charts.TestDisplayChartsWithConversion.docx");
    doc->Save(ArtifactsDir + u"Charts.TestDisplayChartsWithConversion.pdf");
}

namespace gtest_test
{

using TestDisplayChartsWithConversion_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExCharts::TestDisplayChartsWithConversion)>::type;

struct TestDisplayChartsWithConversionVP : public ExCharts, public ApiExamples::ExCharts, public ::testing::WithParamInterface<TestDisplayChartsWithConversion_Args>
{
    static std::vector<TestDisplayChartsWithConversion_Args> TestCases()
    {
        return
        {
            std::make_tuple(Aspose::Words::Drawing::Charts::ChartType::Column),
            std::make_tuple(Aspose::Words::Drawing::Charts::ChartType::Line),
            std::make_tuple(Aspose::Words::Drawing::Charts::ChartType::Pie),
            std::make_tuple(Aspose::Words::Drawing::Charts::ChartType::Bar),
            std::make_tuple(Aspose::Words::Drawing::Charts::ChartType::Area),
        };
    }
};

TEST_P(TestDisplayChartsWithConversionVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TestDisplayChartsWithConversion(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExCharts, TestDisplayChartsWithConversionVP, ::testing::ValuesIn(TestDisplayChartsWithConversionVP::TestCases()));

} // namespace gtest_test

void ExCharts::Surface3DChart()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert chart
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Surface3D, 432, 252);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Clear demo data
    chart->get_Series()->Clear();

    chart->get_Series()->Add(u"Aspose Test Series 1", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

    chart->get_Series()->Add(u"Aspose Test Series 2", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({900000, 50000, 1100000, 400000, 2500000}));

    chart->get_Series()->Add(u"Aspose Test Series 3", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), MakeArray<double>({500000, 820000, 1500000, 400000, 100000}));

    doc->Save(ArtifactsDir + u"Charts.SurfaceChart.docx");
    doc->Save(ArtifactsDir + u"Charts.SurfaceChart.pdf");
}

namespace gtest_test
{

TEST_F(ExCharts, Surface3DChart)
{
    s_instance->Surface3DChart();
}

} // namespace gtest_test

void ExCharts::ChartDataLabelCollection()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.ShowBubbleSize
    //ExFor:ChartDataLabelCollection.ShowCategoryName
    //ExFor:ChartDataLabelCollection.ShowSeriesName
    //ExFor:ChartDataLabelCollection.Separator
    //ExFor:ChartDataLabelCollection.ShowLeaderLines
    //ExFor:ChartDataLabelCollection.ShowLegendKey
    //ExFor:ChartDataLabelCollection.ShowPercentage
    //ExFor:ChartDataLabelCollection.ShowValue
    //ExSummary:Shows how to set default values for the data labels.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert bubble chart
    SharedPtr<Chart> chart = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble, 432, 252)->get_Chart();
    // Clear demo data
    chart->get_Series()->Clear();

    SharedPtr<ChartSeries> bubbleChartSeries = chart->get_Series()->Add(u"Aspose Test Series", MakeArray<double>({2.9, 3.5, 1.1, 4, 4}), MakeArray<double>({1.9, 8.5, 2.1, 6, 1.5}), MakeArray<double>({9, 4.5, 2.5, 8, 5}));

    // Set default values for the bubble chart data labels
    SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> bubbleChartDataLabels = bubbleChartSeries->get_DataLabels();
    bubbleChartDataLabels->set_ShowBubbleSize(true);
    bubbleChartDataLabels->set_ShowCategoryName(true);
    bubbleChartDataLabels->set_ShowSeriesName(true);
    bubbleChartDataLabels->set_Separator(u" - ");

    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);

    // Insert pie chart
    SharedPtr<Shape> shapeWithPieChart = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, 432, 252);
    // Clear demo data
    shapeWithPieChart->get_Chart()->get_Series()->Clear();

    SharedPtr<ChartSeries> pieChartSeries = shapeWithPieChart->get_Chart()->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel"}), MakeArray<double>({2.7, 3.2, 0.8}));

    // Set default values for the pie chart data labels
    SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> pieChartDataLabels = pieChartSeries->get_DataLabels();
    pieChartDataLabels->set_ShowLeaderLines(true);
    pieChartDataLabels->set_ShowLegendKey(true);
    pieChartDataLabels->set_ShowPercentage(true);
    pieChartDataLabels->set_ShowValue(true);

    doc->Save(ArtifactsDir + u"Charts.ChartDataLabelCollection.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartDataLabelCollection.docx");
    bubbleChartDataLabels = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();

    ASSERT_FALSE(bubbleChartDataLabels->get_ShowBubbleSize());
    ASSERT_FALSE(bubbleChartDataLabels->get_ShowCategoryName());
    ASSERT_FALSE(bubbleChartDataLabels->get_ShowSeriesName());
    ASSERT_EQ(u",", bubbleChartDataLabels->get_Separator());

    pieChartDataLabels = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();

    ASSERT_FALSE(pieChartDataLabels->get_ShowLeaderLines());
    ASSERT_FALSE(pieChartDataLabels->get_ShowLegendKey());
    ASSERT_FALSE(pieChartDataLabels->get_ShowPercentage());
    ASSERT_FALSE(pieChartDataLabels->get_ShowValue());
}

namespace gtest_test
{

TEST_F(ExCharts, ChartDataLabelCollection)
{
    s_instance->ChartDataLabelCollection();
}

} // namespace gtest_test

void ExCharts::ChartDataLabels()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Use a document builder to insert a bar chart
    SharedPtr<Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 400, 300);

    // Get the chart object from the containing shape
    SharedPtr<Chart> chart = chartShape->get_Chart();

    // The chart already contains demo data comprised of 3 series each with 4 categories
    ASSERT_EQ(3, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());

    // Apply data labels to every series in the graph
    for (auto series : System::IterateOver(chart->get_Series()))
    {
        ApplyDataLabels(series, 4, u"000.0", u", ");
        ASSERT_EQ(4, series->get_DataLabels()->get_Count());
    }

    // Get the enumerator for a data label collection
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ChartDataLabel>>> enumerator = chart->get_Series()->idx_get(0)->get_DataLabels()->GetEnumerator();
        // And use it to go over all the data labels in one series and change their separator
        while (enumerator->MoveNext())
        {
            ASSERT_EQ(u", ", enumerator->get_Current()->get_Separator());
            enumerator->get_Current()->set_Separator(u" & ");
        }
    }

    // If the chart looks too busy, we can remove data labels one by one
    chart->get_Series()->idx_get(1)->get_DataLabels()->idx_get(2)->ClearFormat();

    // We can also clear an entire data label collection for one whole series
    chart->get_Series()->idx_get(2)->get_DataLabels()->ClearFormat();

    doc->Save(ArtifactsDir + u"Charts.ChartDataLabels.docx");
}

namespace gtest_test
{

TEST_F(ExCharts, ChartDataLabels)
{
    s_instance->ChartDataLabels();
}

} // namespace gtest_test

void ExCharts::ChartDataPoint()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Add a line chart, which will have default data that we will use
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 500, 350);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Apply diamond-shaped data points to the line of the first series
    for (auto series : System::IterateOver(chart->get_Series()))
    {
        ApplyDataPoints(series, 4, Aspose::Words::Drawing::Charts::MarkerSymbol::Diamond, 15);
    }

    // We can further decorate a series line by smoothing it
    chart->get_Series()->idx_get(0)->set_Smooth(true);

    // Get the enumerator for the data point collection from one series
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint>>> enumerator = chart->get_Series()->idx_get(0)->get_DataPoints()->GetEnumerator();
        // And use it to go over all the data labels in one series and change their separator
        while (enumerator->MoveNext())
        {
            ASSERT_FALSE(enumerator->get_Current()->get_InvertIfNegative());
        }
    }

    // If the chart looks too busy, we can remove data points one by one
    chart->get_Series()->idx_get(1)->get_DataPoints()->RemoveAt(2);

    // We can also clear an entire data point collection for one whole series
    chart->get_Series()->idx_get(2)->get_DataPoints()->Clear();

    doc->Save(ArtifactsDir + u"Charts.ChartDataPoint.docx");
}

namespace gtest_test
{

TEST_F(ExCharts, ChartDataPoint)
{
    s_instance->ChartDataPoint();
}

} // namespace gtest_test

void ExCharts::PieChartExplosion()
{
    //ExStart
    //ExFor:Charts.IChartDataPoint.Explosion
    //ExSummary:Shows how to manipulate the position of the portions of a pie chart.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, 500, 350);
    SharedPtr<Chart> chart = shape->get_Chart();

    // In a pie chart, the portions are the data points, which cannot have markers or sizes applied to them
    // However, we can set this variable to move any individual "slice" away from the center of the chart
    SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> cdp = chart->get_Series()->idx_get(0)->get_DataPoints()->Add(0);
    cdp->set_Explosion(10);

    cdp = chart->get_Series()->idx_get(0)->get_DataPoints()->Add(1);
    cdp->set_Explosion(40);

    doc->Save(ArtifactsDir + u"Charts.PieChartExplosion.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.PieChartExplosion.docx");
    SharedPtr<ChartSeries> series = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

    ASSERT_EQ(10, series->get_DataPoints()->idx_get(0)->get_Explosion());
    ASSERT_EQ(40, series->get_DataPoints()->idx_get(1)->get_Explosion());
}

namespace gtest_test
{

TEST_F(ExCharts, PieChartExplosion)
{
    s_instance->PieChartExplosion();
}

} // namespace gtest_test

void ExCharts::Bubble3D()
{
    //ExStart
    //ExFor:Charts.ChartDataLabel.ShowBubbleSize
    //ExFor:Charts.IChartDataPoint.Bubble3D
    //ExSummary:Shows how to use 3D effects with bubble charts.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a bubble chart with 3D effects on each bubble
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble3D, 500, 350);
    SharedPtr<Chart> chart = shape->get_Chart();

    ASSERT_TRUE(chart->get_Series()->idx_get(0)->get_Bubble3D());

    // Apply a data label to each bubble that displays the size of its bubble
    for (int i = 0; i < 3; i++)
    {
        chart->get_Series()->idx_get(0)->set_HasDataLabels(true);
        chart->get_Series()->idx_get(0)->get_DataLabels()->idx_get(i)->set_ShowBubbleSize(true);
    }

    doc->Save(ArtifactsDir + u"Charts.Bubble3D.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.Bubble3D.docx");
    SharedPtr<ChartSeries> series = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

    for (int i = 0; i < 3; i++)
    {
        ASSERT_TRUE(series->get_DataLabels()->idx_get(i)->get_ShowBubbleSize());
    }
}

namespace gtest_test
{

TEST_F(ExCharts, Bubble3D)
{
    s_instance->Bubble3D();
}

} // namespace gtest_test

void ExCharts::ChartSeriesCollection()
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // There are 4 ways of populating a chart's series collection
    // 1: Each series has a string array of categories, each with a corresponding data value
    // Some of the other possible applications are bar, column, line and surface charts
    SharedPtr<Chart> chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Column, 300, 300);

    // Create and name 3 categories with a string array
    ArrayPtr<String> categories = MakeArray<String>({u"Category 1", u"Category 2", u"Category 3"});

    // Create 2 series of data, each with one point for every category
    // This will generate a column graph with 3 clusters of 2 bars
    chart->get_Series()->Add(u"Series 1", categories, MakeArray<double>({76.6, 82.1, 91.6}));
    chart->get_Series()->Add(u"Series 2", categories, MakeArray<double>({64.2, 79.5, 94.0}));

    // Categories are distributed along the X-axis while values are distributed along the Y-axis
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Category, chart->get_AxisX()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisY()->get_Type());

    // 2: Each series will have a collection of dates with a corresponding value for each date
    // Area, radar and stock charts are some of the appropriate chart types for this
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Area, 300, 300);

    // Create a collection of dates to serve as categories
    ArrayPtr<System::DateTime> dates = MakeArray<System::DateTime>({System::DateTime(2014, 3, 31), System::DateTime(2017, 1, 23), System::DateTime(2017, 6, 18), System::DateTime(2019, 11, 22), System::DateTime(2020, 9, 7)});

    // Add one series with one point for each date
    // Our sporadic dates will be distributed along the X-axis in a linear fashion
    chart->get_Series()->Add(u"Series 1", dates, MakeArray<double>({15.8, 21.5, 22.9, 28.7, 33.1}));

    // 3: Each series will take two data arrays
    // Appropriate for scatter plots
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Scatter, 300, 300);

    // In each series, the first array contains the X-coordinates and the second contains respective Y-coordinates of points
    chart->get_Series()->Add(u"Series 1", MakeArray<double>({3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6}), MakeArray<double>({3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9}));
    chart->get_Series()->Add(u"Series 2", MakeArray<double>({2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3}), MakeArray<double>({7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6}));

    // Both axes are value axes in this case
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisX()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisY()->get_Type());

    // 4: Each series will be built from three data arrays, used for bubble charts
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Bubble, 300, 300);

    // The first two arrays contain X/Y coordinates like above and the third determines the thickness of each point
    chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.1, 5.0, 9.8}), MakeArray<double>({1.2, 4.9, 9.9}), MakeArray<double>({2.0, 4.0, 8.0}));

    doc->Save(ArtifactsDir + u"Charts.ChartSeriesCollection.docx");
}

namespace gtest_test
{

TEST_F(ExCharts, ChartSeriesCollection)
{
    s_instance->ChartSeriesCollection();
}

} // namespace gtest_test

void ExCharts::ChartSeriesCollectionModify()
{
    //ExStart
    //ExFor:ChartSeriesCollection
    //ExFor:ChartSeriesCollection.Clear
    //ExFor:ChartSeriesCollection.Count
    //ExFor:ChartSeriesCollection.GetEnumerator
    //ExFor:ChartSeriesCollection.Item(Int32)
    //ExFor:ChartSeriesCollection.RemoveAt(Int32)
    //ExSummary:Shows how to work with a chart's data collection.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Use a document builder to insert a bar chart
    SharedPtr<Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 400, 300);
    SharedPtr<Chart> chart = chartShape->get_Chart();

    // All charts come with demo data
    // This column chart currently has 3 series with 4 categories, which means 4 clusters, 3 columns in each
    SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> chartData = chart->get_Series();
    ASSERT_EQ(3, chartData->get_Count());
    //ExSkip

    // Iterate through the series with an enumerator and print their names
    {
        SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ChartSeries>>> enumerator = chart->get_Series()->GetEnumerator();
        // And use it to go over all the data labels in one series and change their separator
        while (enumerator->MoveNext())
        {
            System::Console::WriteLine(enumerator->get_Current()->get_Name());
        }
    }

    // We can add new data by adding a new series to the collection, with categories and data
    // We will match the existing category/series names in the demo data and add a 4th column to each column cluster
    ArrayPtr<String> categories = MakeArray<String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"});
    chart->get_Series()->Add(u"Series 4", categories, MakeArray<double>({4.4, 7.0, 3.5, 2.1}));
    ASSERT_EQ(4, chartData->get_Count());
    //ExSkip
    ASSERT_EQ(u"Series 4", chartData->idx_get(3)->get_Name());
    //ExSkip

    // We can remove series by index
    chartData->RemoveAt(2);
    ASSERT_EQ(3, chartData->get_Count());
    //ExSkip
    ASSERT_EQ(u"Series 4", chartData->idx_get(2)->get_Name());
    //ExSkip

    // We can also remove out all the series
    // This leaves us with an empty graph and is a convenient way of wiping out demo data
    chartData->Clear();
    ASSERT_EQ(0, chartData->get_Count());
    //ExSkip
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, ChartSeriesCollectionModify)
{
    s_instance->ChartSeriesCollectionModify();
}

} // namespace gtest_test

void ExCharts::AxisScaling()
{
    //ExStart
    //ExFor:AxisScaleType
    //ExFor:AxisScaling
    //ExFor:AxisScaling.LogBase
    //ExFor:AxisScaling.Type
    //ExSummary:Shows how to set up logarithmic axis scaling.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a scatter chart and clear its default data series
    SharedPtr<Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 300);
    SharedPtr<Chart> chart = chartShape->get_Chart();
    chart->get_Series()->Clear();

    // Insert a series with X/Y coordinates for 5 points
    chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.0, 2.0, 3.0, 4.0, 5.0}), MakeArray<double>({1.0, 20.0, 400.0, 8000.0, 160000.0}));

    // The scaling of the X axis is linear by default, which means it will display "0, 1, 2, 3..."
    // Linear axis scaling is suitable for our X-values, but our Y-values call for a logarithmic scale to be represented accurately on a graph
    // We can set the scaling of the Y-axis to Logarithmic with a base of 20
    // The Y-axis will now display "1, 20, 400, 8000...", which is ideal for accurate representation of this set of Y-values
    chart->get_AxisY()->get_Scaling()->set_Type(Aspose::Words::Drawing::Charts::AxisScaleType::Logarithmic);
    chart->get_AxisY()->get_Scaling()->set_LogBase(20.0);

    doc->Save(ArtifactsDir + u"Charts.AxisScaling.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisScaling.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisScaleType::Linear, chart->get_AxisX()->get_Scaling()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisScaleType::Logarithmic, chart->get_AxisY()->get_Scaling()->get_Type());
    ASPOSE_ASSERT_EQ(20.0, chart->get_AxisY()->get_Scaling()->get_LogBase());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisScaling)
{
    s_instance->AxisScaling();
}

} // namespace gtest_test

void ExCharts::AxisBound()
{
    //ExStart
    //ExFor:AxisBound.#ctor(Double)
    //ExFor:AxisBound.#ctor(DateTime)
    //ExFor:AxisBound.IsAuto
    //ExFor:AxisBound.Value
    //ExFor:AxisBound.ValueAsDate
    //ExSummary:Shows how to set custom axis bounds.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a scatter chart, remove default data and populate it with data from a ChartSeries
    SharedPtr<Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 300);
    SharedPtr<Chart> chart = chartShape->get_Chart();
    chart->get_Series()->Clear();
    chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.1, 5.4, 7.9, 3.5, 2.1, 9.7}), MakeArray<double>({2.1, 0.3, 0.6, 3.3, 1.4, 1.9}));

    // By default, the axis bounds are automatically defined so all the series data within the table is included
    ASSERT_TRUE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());

    // If we wish to set our own scale bounds, we need to replace them with new ones
    // Both the axis rulers will go from 0 to 10
    chart->get_AxisX()->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(0.0));
    chart->get_AxisX()->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(10.0));
    chart->get_AxisY()->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(0.0));
    chart->get_AxisY()->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(10.0));

    // These are custom and not defined automatically
    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());

    // Create a line graph
    chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 450, 300);
    chart = chartShape->get_Chart();
    chart->get_Series()->Clear();

    // Create a collection of dates, which will make up the X axis
    ArrayPtr<System::DateTime> dates = MakeArray<System::DateTime>({System::DateTime(1973, 5, 11), System::DateTime(1981, 2, 4), System::DateTime(1985, 9, 23), System::DateTime(1989, 6, 28), System::DateTime(1994, 12, 15)});

    // Assign a Y-value for each date
    chart->get_Series()->Add(u"Series 1", dates, MakeArray<double>({3.0, 4.7, 5.9, 7.1, 8.9}));

    // These particular bounds will cut off categories from before 1980 and from 1990 and onwards
    // This narrows the amount of categories and values in the viewport from 5 to 3
    // Note that the graph still contains the out-of-range data because we can see the line tend towards it
    chart->get_AxisX()->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1980, 1, 1)));
    chart->get_AxisX()->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1990, 1, 1)));

    doc->Save(ArtifactsDir + u"Charts.AxisBound.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisBound.docx");
    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();

    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(0.0, chart->get_AxisX()->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(10.0, chart->get_AxisX()->get_Scaling()->get_Maximum()->get_Value());

    ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(0.0, chart->get_AxisY()->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(10.0, chart->get_AxisY()->get_Scaling()->get_Maximum()->get_Value());

    chart = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true)))->get_Chart();

    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1980, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1990, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Maximum());

    ASSERT_TRUE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisBound)
{
    s_instance->AxisBound();
}

} // namespace gtest_test

void ExCharts::ChartLegend()
{
    //ExStart
    //ExFor:Chart.Legend
    //ExFor:ChartLegend
    //ExFor:ChartLegend.Overlay
    //ExFor:ChartLegend.Position
    //ExFor:LegendPosition
    //ExSummary:Shows how to edit the appearance of a chart's legend.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a line graph
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 450, 300);

    // Get its legend
    SharedPtr<Aspose::Words::Drawing::Charts::ChartLegend> legend = shape->get_Chart()->get_Legend();

    // By default, other elements of a chart will not overlap with its legend
    ASSERT_FALSE(legend->get_Overlay());

    // We can move its position by setting this attribute
    legend->set_Position(Aspose::Words::Drawing::Charts::LegendPosition::TopRight);

    doc->Save(ArtifactsDir + u"Charts.ChartLegend.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartLegend.docx");

    legend = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Legend();

    ASSERT_FALSE(legend->get_Overlay());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::LegendPosition::TopRight, legend->get_Position());
}

namespace gtest_test
{

TEST_F(ExCharts, ChartLegend)
{
    s_instance->ChartLegend();
}

} // namespace gtest_test

void ExCharts::AxisCross()
{
    //ExStart
    //ExFor:ChartAxis.AxisBetweenCategories
    //ExFor:ChartAxis.CrossesAt
    //ExSummary:Shows how to get a graph axis to cross at a custom location.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a column chart, which is populated by default values
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 450, 250);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Get the Y-axis to cross at a value of 3.0, making 3.0 the new Y-zero of our column chart
    // This effectively means that all the columns with Y-values about 3.0 will be above the Y-centre and point up,
    // while ones below 3.0 will point down
    SharedPtr<ChartAxis> axis = chart->get_AxisX();
    axis->set_AxisBetweenCategories(true);
    axis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Custom);
    axis->set_CrossesAt(3.0);

    doc->Save(ArtifactsDir + u"Charts.AxisCross.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisCross.docx");
    axis = (System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_AxisX();

    ASSERT_TRUE(axis->get_AxisBetweenCategories());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCrosses::Custom, axis->get_Crosses());
    ASPOSE_ASSERT_EQ(3.0, axis->get_CrossesAt());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisCross)
{
    s_instance->AxisCross();
}

} // namespace gtest_test

void ExCharts::ChartAxisDisplayUnit()
{
    //ExStart
    //ExFor:AxisBuiltInUnit
    //ExFor:ChartAxis.DisplayUnit
    //ExFor:ChartAxis.MajorUnitIsAuto
    //ExFor:ChartAxis.MajorUnitScale
    //ExFor:ChartAxis.MinorUnitIsAuto
    //ExFor:ChartAxis.MinorUnitScale
    //ExFor:ChartAxis.TickLabelSpacing
    //ExFor:ChartAxis.TickLabelAlignment
    //ExFor:AxisDisplayUnit
    //ExFor:AxisDisplayUnit.CustomUnit
    //ExFor:AxisDisplayUnit.Unit
    //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // Insert a scatter chart, which is populated by default values
    SharedPtr<Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 250);
    SharedPtr<Chart> chart = shape->get_Chart();

    // Set they Y axis to show major ticks every at every 10 units and minor ticks at every 1 units
    SharedPtr<ChartAxis> axis = chart->get_AxisY();
    axis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Outside);
    axis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Outside);

    axis->set_MajorUnit(10.0);
    axis->set_MinorUnit(1.0);

    // Stretch out the bounds of the axis out to show 3 major ticks and 27 minor ticks
    axis->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(-10.0));
    axis->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(20.0));

    // Do the same for the X-axis
    axis = chart->get_AxisX();
    axis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    axis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    axis->set_MajorUnit(10.0);
    axis->get_Scaling()->set_Minimum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(-10.0));
    axis->get_Scaling()->set_Maximum(MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(30.0));

    // We can also use this attribute to set minor tick spacing
    axis->set_TickLabelSpacing(2);
    // We can define text alignment when axis tick labels are multi-line
    // MS Word aligns them to the center by default
    axis->set_TickLabelAlignment(Aspose::Words::ParagraphAlignment::Right);

    // Get the axis to display values, but in millions
    axis->get_DisplayUnit()->set_Unit(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Millions);
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Millions, axis->get_DisplayUnit()->get_Unit());
    //ExSkip

    // Besides the built-in axis units we can choose from,
    // we can also set the axis to display values in some custom denomination, using the following attribute
    // The statement below is equivalent to the one above
    axis->get_DisplayUnit()->set_CustomUnit(1000000.0);
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
    //ExSkip

    doc->Save(ArtifactsDir + u"Charts.ChartAxisDisplayUnit.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartAxisDisplayUnit.docx");
    shape = System::DynamicCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));

    ASPOSE_ASSERT_EQ(450.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(250.0, shape->get_Height());

    axis = shape->get_Chart()->get_AxisX();

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, axis->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, axis->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
    ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(30.0, axis->get_Scaling()->get_Maximum()->get_Value());
    ASSERT_EQ(1, axis->get_TickLabelSpacing());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, axis->get_TickLabelAlignment());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
    ASPOSE_ASSERT_EQ(1000000.0, axis->get_DisplayUnit()->get_CustomUnit());

    axis = shape->get_Chart()->get_AxisY();

    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Outside, axis->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Outside, axis->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
    ASPOSE_ASSERT_EQ(1.0, axis->get_MinorUnit());
    ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(20.0, axis->get_Scaling()->get_Maximum()->get_Value());
}

namespace gtest_test
{

TEST_F(ExCharts, ChartAxisDisplayUnit)
{
    s_instance->ChartAxisDisplayUnit();
}

} // namespace gtest_test

} // namespace ApiExamples
