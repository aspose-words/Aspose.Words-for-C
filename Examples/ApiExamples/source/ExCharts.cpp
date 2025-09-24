// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCharts.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/compare.h>
#include <system/primitive_types.h>
#include <system/object_ext.h>
#include <system/math.h>
#include <system/linq/enumerable.h>
#include <system/globalization/number_format_info.h>
#include <system/globalization/culture_info.h>
#include <system/func.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/date_time.h>
#include <system/collections/ienumerator.h>
#include <system/array.h>
#include <iostream>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeTextOrientation.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/PresetTexture.h>
#include <Aspose.Words.Cpp/Model/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Model/Drawing/DashStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/LegendPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartYValueCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartYValue.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartXValueCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartXValue.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartStyle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesGroupCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesGroup.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartNumberFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartMultilevelValue.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartMarker.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartLegendEntryCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartLegendEntry.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartLegend.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataTable.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPointCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPoint.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelLocationMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxisType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxisTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxisCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxis.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/BubbleSizeCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTimeUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickMark.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickLabels.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickLabelPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisScaling.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisScaleType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisGroup.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisDisplayUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCrosses.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCategoryType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBuiltInUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBound.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2302797093u, ::Aspose::Words::ApiExamples::ExCharts, ThisTypeBaseTypesInfo);

void ExCharts::ApplyDataLabels(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t labelsCount, System::String numberFormat, System::String separator)
{
    series->set_HasDataLabels(true);
    series->set_Explosion(40);
    
    for (int32_t i = 0; i < labelsCount; i++)
    {
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsVisible());
        
        series->get_DataLabels()->idx_get(i)->set_ShowCategoryName(true);
        series->get_DataLabels()->idx_get(i)->set_ShowSeriesName(true);
        series->get_DataLabels()->idx_get(i)->set_ShowValue(true);
        series->get_DataLabels()->idx_get(i)->set_ShowLeaderLines(true);
        series->get_DataLabels()->idx_get(i)->set_ShowLegendKey(true);
        series->get_DataLabels()->idx_get(i)->set_ShowPercentage(false);
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsHidden());
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_ShowDataLabelsRange());
        
        series->get_DataLabels()->idx_get(i)->get_NumberFormat()->set_FormatCode(numberFormat);
        series->get_DataLabels()->idx_get(i)->set_Separator(separator);
        
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_ShowDataLabelsRange());
        ASSERT_TRUE(series->get_DataLabels()->idx_get(i)->get_IsVisible());
        ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsHidden());
    }
}

void ExCharts::ApplyDataPoints(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t dataPointsCount, Aspose::Words::Drawing::Charts::MarkerSymbol markerSymbol, int32_t dataPointSize)
{
    for (int32_t i = 0; i < dataPointsCount; i++)
    {
        System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> point = series->get_DataPoints()->idx_get(i);
        point->get_Marker()->set_Symbol(markerSymbol);
        point->get_Marker()->set_Size(dataPointSize);
        
        ASSERT_EQ(i, point->get_Index());
    }
}

System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> ExCharts::AppendChart(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::Drawing::Charts::ChartType chartType, double width, double height)
{
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(chartType, width, height);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    chart->get_Series()->Clear();
    EXPECT_EQ(0, chart->get_Series()->get_Count());
    //ExSkip
    
    return chart;
}


namespace gtest_test
{

class ExCharts : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExCharts> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExCharts>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExCharts> ExCharts::s_instance;

} // namespace gtest_test

void ExCharts::ChartTitle()
{
    //ExStart:ChartTitle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:Chart
    //ExFor:Chart.Title
    //ExFor:ChartTitle
    //ExFor:ChartTitle.Overlay
    //ExFor:ChartTitle.Show
    //ExFor:ChartTitle.Text
    //ExFor:ChartTitle.Font
    //ExSummary:Shows how to insert a chart and set a title.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a chart shape with a document builder and get its chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bar, 400, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    
    // Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartTitle> title = chart->get_Title();
    title->set_Text(u"My Chart");
    title->get_Font()->set_Size(15);
    title->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    
    // Set the "Show" property to "true" to make the title visible. 
    title->set_Show(true);
    
    // Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
    title->set_Overlay(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartTitle.docx");
    //ExEnd:ChartTitle
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.ChartTitle.docx");
    chartShape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASSERT_EQ(Aspose::Words::Drawing::ShapeType::NonPrimitive, chartShape->get_ShapeType());
    ASSERT_TRUE(chartShape->get_HasChart());
    
    title = chartShape->get_Chart()->get_Title();
    
    ASSERT_EQ(u"My Chart", title->get_Text());
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

void ExCharts::DataLabelNumberFormat()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.NumberFormat
    //ExFor:ChartDataLabelCollection.Font
    //ExFor:ChartNumberFormat.FormatCode
    //ExFor:ChartSeries.HasDataLabels
    //ExSummary:Shows how to enable and configure data labels for a chart series.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Add a line chart, then clear its demo data series to start with a clean chart,
    // and then set a title.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Series()->Clear();
    chart->get_Title()->set_Text(u"Monthly sales report");
    
    // Insert a custom chart series with months as categories for the X-axis,
    // and respective decimal amounts for the Y-axis.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Revenue", System::MakeArray<System::String>({u"January", u"February", u"March"}), System::MakeArray<double>({25.611, 21.439, 33.750}));
    
    // Enable data labels, and then apply a custom number format for values displayed in the data labels.
    // This format will treat displayed decimal values as millions of US Dollars.
    series->set_HasDataLabels(true);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    dataLabels->set_ShowValue(true);
    dataLabels->get_NumberFormat()->set_FormatCode(u"\"US$\" #,##0.000\"M\"");
    dataLabels->get_Font()->set_Size(12);
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataLabelNumberFormat.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.DataLabelNumberFormat.docx");
    series = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);
    
    ASSERT_TRUE(series->get_HasDataLabels());
    ASSERT_TRUE(series->get_DataLabels()->get_ShowValue());
    ASSERT_EQ(u"\"US$\" #,##0.000\"M\"", series->get_DataLabels()->get_NumberFormat()->get_FormatCode());
}

namespace gtest_test
{

TEST_F(ExCharts, DataLabelNumberFormat)
{
    s_instance->DataLabelNumberFormat();
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
    //ExFor:ChartAxis.Document
    //ExFor:ChartAxis.TickLabels
    //ExFor:ChartAxis.Format
    //ExFor:AxisTickLabels
    //ExFor:AxisTickLabels.Offset
    //ExFor:AxisTickLabels.Position
    //ExFor:AxisTickLabels.IsAutoSpacing
    //ExFor:AxisTickLabels.Alignment
    //ExFor:AxisTickLabels.Font
    //ExFor:AxisTickLabels.Spacing
    //ExFor:ChartAxis.TickMarkSpacing
    //ExFor:AxisCategoryType
    //ExFor:AxisCrosses
    //ExFor:Chart.AxisX
    //ExFor:Chart.AxisY
    //ExFor:Chart.AxisZ
    //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
    chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({640, 320, 280, 120, 150}));
    
    // Chart axes have various options that can change their appearance,
    // such as their direction, major/minor unit ticks, and tick marks.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> xAxis = chart->get_AxisX();
    xAxis->set_CategoryType(Aspose::Words::Drawing::Charts::AxisCategoryType::Category);
    xAxis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Minimum);
    xAxis->set_ReverseOrder(false);
    xAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    xAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    xAxis->set_MajorUnit(10.0);
    xAxis->set_MinorUnit(15.0);
    xAxis->get_TickLabels()->set_Offset(50);
    xAxis->get_TickLabels()->set_Position(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::Low);
    xAxis->get_TickLabels()->set_IsAutoSpacing(false);
    xAxis->set_TickMarkSpacing(1);
    
    ASPOSE_ASSERT_EQ(doc, xAxis->get_Document());
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> yAxis = chart->get_AxisY();
    yAxis->set_CategoryType(Aspose::Words::Drawing::Charts::AxisCategoryType::Automatic);
    yAxis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Maximum);
    yAxis->set_ReverseOrder(true);
    yAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    yAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    yAxis->set_MajorUnit(100.0);
    yAxis->set_MinorUnit(20.0);
    yAxis->get_TickLabels()->set_Position(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::NextToAxis);
    yAxis->get_TickLabels()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    yAxis->get_TickLabels()->get_Font()->set_Color(System::Drawing::Color::get_Red());
    yAxis->get_TickLabels()->set_Spacing(1);
    
    // Column charts do not have a Z-axis.
    ASSERT_TRUE(System::TestTools::IsNull(chart->get_AxisZ()));
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisProperties.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.AxisProperties.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCategoryType::Category, chart->get_AxisX()->get_CategoryType());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCrosses::Minimum, chart->get_AxisX()->get_Crosses());
    ASSERT_FALSE(chart->get_AxisX()->get_ReverseOrder());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, chart->get_AxisX()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisX()->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(0.5, chart->get_AxisX()->get_MinorUnit());
    ASSERT_EQ(50, chart->get_AxisX()->get_TickLabels()->get_Offset());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::Low, chart->get_AxisX()->get_TickLabels()->get_Position());
    ASSERT_FALSE(chart->get_AxisX()->get_TickLabels()->get_IsAutoSpacing());
    ASSERT_EQ(1, chart->get_AxisX()->get_TickMarkSpacing());
    ASSERT_TRUE(chart->get_AxisX()->get_Format()->get_IsDefined());
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCategoryType::Category, chart->get_AxisY()->get_CategoryType());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisCrosses::Maximum, chart->get_AxisY()->get_Crosses());
    ASSERT_TRUE(chart->get_AxisY()->get_ReverseOrder());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, chart->get_AxisY()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisY()->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(20.0, chart->get_AxisY()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::NextToAxis, chart->get_AxisY()->get_TickLabels()->get_Position());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Center, chart->get_AxisY()->get_TickLabels()->get_Alignment());
    ASSERT_EQ(System::Drawing::Color::get_Red().ToArgb(), chart->get_AxisY()->get_TickLabels()->get_Font()->get_Color().ToArgb());
    ASSERT_EQ(1, chart->get_AxisY()->get_TickLabels()->get_Spacing());
    ASSERT_TRUE(chart->get_AxisY()->get_Format()->get_IsDefined());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisProperties)
{
    s_instance->AxisProperties();
}

} // namespace gtest_test

void ExCharts::AxisCollection()
{
    //ExStart
    //ExFor:ChartAxisCollection
    //ExFor:Chart.Axes
    //ExSummary:Shows how to work with axes collection.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Hide the major grid lines on the primary and secondary Y axes.
    for (auto&& axis : System::IterateOver(chart->get_Axes()))
    {
        if (axis->get_Type() == Aspose::Words::Drawing::Charts::ChartAxisType::Value)
        {
            axis->set_HasMajorGridlines(false);
        }
    }
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisCollection.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, AxisCollection)
{
    s_instance->AxisCollection();
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
    //ExFor:AxisTickMark
    //ExFor:AxisTickLabelPosition
    //ExFor:AxisTimeUnit
    //ExFor:ChartAxis.BaseTimeUnit
    //ExFor:ChartAxis.HasMajorGridlines
    //ExFor:ChartAxis.HasMinorGridlines
    //ExSummary:Shows how to insert chart with date/time values.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
    chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<System::DateTime>({System::DateTime(2017, 11, 6), System::DateTime(2017, 11, 9), System::DateTime(2017, 11, 15), System::DateTime(2017, 11, 21), System::DateTime(2017, 11, 25), System::DateTime(2017, 11, 29)}), System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2, 5.3}));
    
    
    // Set lower and upper bounds for the X-axis.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> xAxis = chart->get_AxisX();
    xAxis->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 11, 5).ToOADate()));
    xAxis->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 12, 3)));
    
    // Set the major units of the X-axis to a week, and the minor units to a day.
    xAxis->set_BaseTimeUnit(Aspose::Words::Drawing::Charts::AxisTimeUnit::Days);
    xAxis->set_MajorUnit(7.0);
    xAxis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    xAxis->set_MinorUnit(1.0);
    xAxis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Outside);
    xAxis->set_HasMajorGridlines(true);
    xAxis->set_HasMinorGridlines(true);
    
    // Define Y-axis properties for decimal values.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> yAxis = chart->get_AxisY();
    yAxis->get_TickLabels()->set_Position(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::High);
    yAxis->set_MajorUnit(100.0);
    yAxis->set_MinorUnit(50.0);
    yAxis->get_DisplayUnit()->set_Unit(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Hundreds);
    yAxis->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(100.0));
    yAxis->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(700.0));
    yAxis->set_HasMajorGridlines(true);
    yAxis->set_HasMinorGridlines(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.DateTimeValues.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.DateTimeValues.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 11, 5).ToOADate()), chart->get_AxisX()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(2017, 12, 3)), chart->get_AxisX()->get_Scaling()->get_Maximum());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTimeUnit::Days, chart->get_AxisX()->get_BaseTimeUnit());
    ASPOSE_ASSERT_EQ(7.0, chart->get_AxisX()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, chart->get_AxisX()->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Outside, chart->get_AxisX()->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(true, chart->get_AxisX()->get_HasMajorGridlines());
    ASPOSE_ASSERT_EQ(true, chart->get_AxisX()->get_HasMinorGridlines());
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::High, chart->get_AxisY()->get_TickLabels()->get_Position());
    ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
    ASPOSE_ASSERT_EQ(50.0, chart->get_AxisY()->get_MinorUnit());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Hundreds, chart->get_AxisY()->get_DisplayUnit()->get_Unit());
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(100.0), chart->get_AxisY()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(700.0), chart->get_AxisY()->get_Scaling()->get_Maximum());
    ASPOSE_ASSERT_EQ(true, chart->get_AxisY()->get_HasMajorGridlines());
    ASPOSE_ASSERT_EQ(true, chart->get_AxisY()->get_HasMinorGridlines());
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
    chart->get_Series()->Add(u"AW Series 1", System::MakeArray<System::String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}), System::MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));
    
    // Hide the chart axes to simplify the appearance of the chart. 
    chart->get_AxisX()->set_Hidden(true);
    chart->get_AxisY()->set_Hidden(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.HideChartAxis.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.HideChartAxis.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
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
    //ExFor:ChartNumberFormat
    //ExFor:ChartNumberFormat.FormatCode
    //ExFor:ChartNumberFormat.IsLinkedToSource
    //ExSummary:Shows how to set formatting for chart values.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Add a custom series to the chart with categories for the X-axis,
    // and large respective numeric values for the Y-axis. 
    chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));
    
    // Set the number format of the Y-axis tick labels to not group digits with commas. 
    chart->get_AxisY()->get_NumberFormat()->set_FormatCode(u"#,##0");
    
    // This flag can override the above value and draw the number format from the source cell.
    ASSERT_FALSE(chart->get_AxisY()->get_NumberFormat()->get_IsLinkedToSource());
    
    doc->Save(get_ArtifactsDir() + u"Charts.SetNumberFormatToChartAxis.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.SetNumberFormatToChartAxis.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(chartType, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Series()->Clear();
    
    chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.TestDisplayChartsWithConversion.docx");
    doc->Save(get_ArtifactsDir() + u"Charts.TestDisplayChartsWithConversion.pdf");
}

namespace gtest_test
{

using ExCharts_TestDisplayChartsWithConversion_Args = System::MethodArgumentTuple<decltype(&Aspose::Words::ApiExamples::ExCharts::TestDisplayChartsWithConversion)>::type;

struct ExCharts_TestDisplayChartsWithConversion : public ExCharts, public Aspose::Words::ApiExamples::ExCharts, public ::testing::WithParamInterface<ExCharts_TestDisplayChartsWithConversion_Args>
{
    static std::vector<ParamType> TestCases()
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

TEST_P(ExCharts_TestDisplayChartsWithConversion, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TestDisplayChartsWithConversion(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExCharts_TestDisplayChartsWithConversion, ::testing::ValuesIn(ExCharts_TestDisplayChartsWithConversion::TestCases()));

} // namespace gtest_test

void ExCharts::Surface3DChart()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Surface3D, 500, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Series()->Clear();
    
    chart->get_Series()->Add(u"Aspose Test Series 1", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));
    
    chart->get_Series()->Add(u"Aspose Test Series 2", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({900000, 50000, 1100000, 400000, 2500000}));
    
    chart->get_Series()->Add(u"Aspose Test Series 3", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}), System::MakeArray<double>({500000, 820000, 1500000, 400000, 100000}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.SurfaceChart.docx");
    doc->Save(get_ArtifactsDir() + u"Charts.SurfaceChart.pdf");
}

namespace gtest_test
{

TEST_F(ExCharts, Surface3DChart)
{
    s_instance->Surface3DChart();
}

} // namespace gtest_test

void ExCharts::DataLabelsBubbleChart()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.Separator
    //ExFor:ChartDataLabelCollection.ShowBubbleSize
    //ExFor:ChartDataLabelCollection.ShowCategoryName
    //ExFor:ChartDataLabelCollection.ShowSeriesName
    //ExSummary:Shows how to work with data labels of a bubble chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble, 500, 300)->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Add a custom series with X/Y coordinates and diameter of each of the bubbles. 
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<double>({2.9, 3.5, 1.1, 4.0, 4.0}), System::MakeArray<double>({1.9, 8.5, 2.1, 6.0, 1.5}), System::MakeArray<double>({9.0, 4.5, 2.5, 8.0, 5.0}));
    
    // Enable data labels, and then modify their appearance.
    series->set_HasDataLabels(true);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    dataLabels->set_ShowBubbleSize(true);
    dataLabels->set_ShowCategoryName(true);
    dataLabels->set_ShowSeriesName(true);
    dataLabels->set_Separator(u" & ");
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataLabelsBubbleChart.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.DataLabelsBubbleChart.docx");
    dataLabels = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();
    
    ASSERT_TRUE(dataLabels->get_ShowBubbleSize());
    ASSERT_TRUE(dataLabels->get_ShowCategoryName());
    ASSERT_TRUE(dataLabels->get_ShowSeriesName());
    ASSERT_EQ(u" & ", dataLabels->get_Separator());
}

namespace gtest_test
{

TEST_F(ExCharts, DataLabelsBubbleChart)
{
    s_instance->DataLabelsBubbleChart();
}

} // namespace gtest_test

void ExCharts::DataLabelsPieChart()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.Separator
    //ExFor:ChartDataLabelCollection.ShowLeaderLines
    //ExFor:ChartDataLabelCollection.ShowLegendKey
    //ExFor:ChartDataLabelCollection.ShowPercentage
    //ExFor:ChartDataLabelCollection.ShowValue
    //ExSummary:Shows how to work with data labels of a pie chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, 500, 300)->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Insert a custom chart series with a category name for each of the sectors, and their frequency table.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Aspose Test Series", System::MakeArray<System::String>({u"Word", u"PDF", u"Excel"}), System::MakeArray<double>({2.7, 3.2, 0.8}));
    
    // Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
    series->set_HasDataLabels(true);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    dataLabels->set_ShowLeaderLines(true);
    dataLabels->set_ShowLegendKey(true);
    dataLabels->set_ShowPercentage(true);
    dataLabels->set_ShowValue(true);
    dataLabels->set_Separator(u"; ");
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataLabelsPieChart.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.DataLabelsPieChart.docx");
    dataLabels = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();
    
    ASSERT_TRUE(dataLabels->get_ShowLeaderLines());
    ASSERT_TRUE(dataLabels->get_ShowLegendKey());
    ASSERT_TRUE(dataLabels->get_ShowPercentage());
    ASSERT_TRUE(dataLabels->get_ShowValue());
    ASSERT_EQ(u"; ", dataLabels->get_Separator());
}

namespace gtest_test
{

TEST_F(ExCharts, DataLabelsPieChart)
{
    s_instance->DataLabelsPieChart();
}

} // namespace gtest_test

void ExCharts::DataLabels()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 400, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    
    ASSERT_EQ(3, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());
    
    // Apply data labels to every series in the chart.
    // These labels will appear next to each data point in the graph and display its value.
    for (auto&& series : System::IterateOver(chart->get_Series()))
    {
        ApplyDataLabels(series, 4, u"000.0", u", ");
        ASSERT_EQ(4, series->get_DataLabels()->get_Count());
    }
    
    // Change the separator string for every data label in a series.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabel>>> enumerator = chart->get_Series()->idx_get(0)->get_DataLabels()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            ASSERT_EQ(u", ", enumerator->get_Current()->get_Separator());
            enumerator->get_Current()->set_Separator(u" & ");
        }
    }
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabel> dataLabel = chart->get_Series()->idx_get(1)->get_DataLabels()->idx_get(2);
    dataLabel->get_Format()->get_Fill()->set_Color(System::Drawing::Color::get_Red());
    
    // For a cleaner looking graph, we can remove data labels individually.
    dataLabel->ClearFormat();
    
    // We can also strip an entire series of its data labels at once.
    chart->get_Series()->idx_get(2)->get_DataLabels()->ClearFormat();
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataLabels.docx");
}

namespace gtest_test
{

TEST_F(ExCharts, DataLabels)
{
    s_instance->DataLabels();
}

} // namespace gtest_test

void ExCharts::ChartDataPoint()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 500, 350);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(3, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());
    
    // Emphasize the chart's data points by making them appear as diamond shapes.
    for (auto&& series : System::IterateOver(chart->get_Series()))
    {
        ApplyDataPoints(series, 4, Aspose::Words::Drawing::Charts::MarkerSymbol::Diamond, 15);
    }
    
    // Smooth out the line that represents the first data series.
    chart->get_Series()->idx_get(0)->set_Smooth(true);
    
    // Verify that data points for the first series will not invert their colors if the value is negative.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint>>> enumerator = chart->get_Series()->idx_get(0)->get_DataPoints()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            ASSERT_FALSE(enumerator->get_Current()->get_InvertIfNegative());
        }
    }
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> dataPoint = chart->get_Series()->idx_get(1)->get_DataPoints()->idx_get(2);
    dataPoint->get_Format()->get_Fill()->set_Color(System::Drawing::Color::get_Red());
    
    // For a cleaner looking graph, we can clear format individually.
    dataPoint->ClearFormat();
    
    // We can also strip an entire series of data points at once.
    chart->get_Series()->idx_get(2)->get_DataPoints()->ClearFormat();
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartDataPoint.docx");
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
    //ExFor:IChartDataPoint.Explosion
    //ExFor:ChartDataPoint.Explosion
    //ExSummary:Shows how to move the slices of a pie chart away from the center.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pie, 500, 350);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(1, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Sales", chart->get_Series()->idx_get(0)->get_Name());
    
    // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
    // Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
    // Aspose.Words create data points automatically if them does not exist.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> dataPoint = chart->get_Series()->idx_get(0)->get_DataPoints()->idx_get(0);
    dataPoint->set_Explosion(10);
    
    // Displace the second portion by a greater distance.
    dataPoint = chart->get_Series()->idx_get(0)->get_DataPoints()->idx_get(1);
    dataPoint->set_Explosion(40);
    
    doc->Save(get_ArtifactsDir() + u"Charts.PieChartExplosion.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.PieChartExplosion.docx");
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);
    
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
    //ExFor:ChartDataLabel.ShowBubbleSize
    //ExFor:ChartDataLabel.Font
    //ExFor:IChartDataPoint.Bubble3D
    //ExFor:ChartSeries.Bubble3D
    //ExSummary:Shows how to use 3D effects with bubble charts.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble3D, 500, 350);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(1, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Y-Values", chart->get_Series()->idx_get(0)->get_Name());
    ASSERT_TRUE(chart->get_Series()->idx_get(0)->get_Bubble3D());
    
    // Apply a data label to each bubble that displays its diameter.
    for (int32_t i = 0; i < 3; i++)
    {
        chart->get_Series()->idx_get(0)->set_HasDataLabels(true);
        chart->get_Series()->idx_get(0)->get_DataLabels()->idx_get(i)->set_ShowBubbleSize(true);
        chart->get_Series()->idx_get(0)->get_DataLabels()->idx_get(i)->get_Font()->set_Size(12);
    }
    
    doc->Save(get_ArtifactsDir() + u"Charts.Bubble3D.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.Bubble3D.docx");
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);
    
    for (int32_t i = 0; i < 3; i++)
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // There are several ways of populating a chart's series collection.
    // Different series schemas are intended for different chart types.
    // 1 -  Column chart with columns grouped and banded along the X-axis by category:
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Column, 500, 300);
    
    System::ArrayPtr<System::String> categories = System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3"});
    
    // Insert two series of decimal values containing a value for each respective category.
    // This column chart will have three groups, each with two columns.
    chart->get_Series()->Add(u"Series 1", categories, System::MakeArray<double>({76.6, 82.1, 91.6}));
    chart->get_Series()->Add(u"Series 2", categories, System::MakeArray<double>({64.2, 79.5, 94.0}));
    
    // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Category, chart->get_AxisX()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisY()->get_Type());
    
    // 2 -  Area chart with dates distributed along the X-axis:
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Area, 500, 300);
    
    System::ArrayPtr<System::DateTime> dates = System::MakeArray<System::DateTime>({System::DateTime(2014, 3, 31), System::DateTime(2017, 1, 23), System::DateTime(2017, 6, 18), System::DateTime(2019, 11, 22), System::DateTime(2020, 9, 7)});
    
    // Insert a series with a decimal value for each respective date.
    // The dates will be distributed along a linear X-axis,
    // and the values added to this series will create data points.
    chart->get_Series()->Add(u"Series 1", dates, System::MakeArray<double>({15.8, 21.5, 22.9, 28.7, 33.1}));
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Category, chart->get_AxisX()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisY()->get_Type());
    
    // 3 -  2D scatter plot:
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Scatter, 500, 300);
    
    // Each series will need two decimal arrays of equal length.
    // The first array contains X-values, and the second contains corresponding Y-values
    // of data points on the chart's graph.
    chart->get_Series()->Add(u"Series 1", System::MakeArray<double>({3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6}), System::MakeArray<double>({3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9}));
    chart->get_Series()->Add(u"Series 2", System::MakeArray<double>({2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3}), System::MakeArray<double>({7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6}));
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisX()->get_Type());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartAxisType::Value, chart->get_AxisY()->get_Type());
    
    // 4 -  Bubble chart:
    chart = AppendChart(builder, Aspose::Words::Drawing::Charts::ChartType::Bubble, 500, 300);
    
    // Each series will need three decimal arrays of equal length.
    // The first array contains X-values, the second contains corresponding Y-values,
    // and the third contains diameters for each of the graph's data points.
    chart->get_Series()->Add(u"Series 1", System::MakeArray<double>({1.1, 5.0, 9.8}), System::MakeArray<double>({1.2, 4.9, 9.9}), System::MakeArray<double>({2.0, 4.0, 8.0}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartSeriesCollection.docx");
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
    //ExSummary:Shows how to add and remove series data in a chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a column chart that will contain three series of demo data by default.
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 400, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    
    // Each series has four decimal values: one for each of the four categories.
    // Four clusters of three columns will represent this data.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> chartData = chart->get_Series();
    
    ASSERT_EQ(3, chartData->get_Count());
    
    // Print the name of every series in the chart.
    {
        System::SharedPtr<System::Collections::Generic::IEnumerator<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries>>> enumerator = chart->get_Series()->GetEnumerator();
        while (enumerator->MoveNext())
        {
            std::cout << enumerator->get_Current()->get_Name() << std::endl;
        }
    }
    
    // These are the names of the categories in the chart.
    System::ArrayPtr<System::String> categories = System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"});
    
    // We can add a series with new values for existing categories.
    // This chart will now contain four clusters of four columns.
    chart->get_Series()->Add(u"Series 4", categories, System::MakeArray<double>({4.4, 7.0, 3.5, 2.1}));
    ASSERT_EQ(4, chartData->get_Count());
    //ExSkip
    ASSERT_EQ(u"Series 4", chartData->idx_get(3)->get_Name());
    //ExSkip
    
    // A chart series can also be removed by index, like this.
    // This will remove one of the three demo series that came with the chart.
    chartData->RemoveAt(2);
    
    ASSERT_FALSE(chartData->LINQ_Any(static_cast<System::Func<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries>, bool>>(static_cast<std::function<bool(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> s)>>([](System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> s) -> bool
    {
        return s->get_Name() == u"Series 3";
    }))));
    ASSERT_EQ(3, chartData->get_Count());
    //ExSkip
    ASSERT_EQ(u"Series 4", chartData->idx_get(2)->get_Name());
    //ExSkip
    
    // We can also clear all the chart's data at once with this method.
    // When creating a new chart, this is the way to wipe all the demo data
    // before we can begin working on a blank chart.
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
    //ExSummary:Shows how to apply logarithmic scaling to a chart axis.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Insert a series with X/Y coordinates for five points.
    chart->get_Series()->Add(u"Series 1", System::MakeArray<double>({1.0, 2.0, 3.0, 4.0, 5.0}), System::MakeArray<double>({1.0, 20.0, 400.0, 8000.0, 160000.0}));
    
    // The scaling of the X-axis is linear by default,
    // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
    // A linear axis is not ideal for our Y-values
    // since the points with the smaller Y-values will be harder to read.
    // A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
    // will spread the plotted points, allowing us to read their values on the chart more easily.
    chart->get_AxisY()->get_Scaling()->set_Type(Aspose::Words::Drawing::Charts::AxisScaleType::Logarithmic);
    chart->get_AxisY()->get_Scaling()->set_LogBase(20);
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisScaling.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.AxisScaling.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
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
    //ExFor:AxisBound.#ctor
    //ExFor:AxisBound.IsAuto
    //ExFor:AxisBound.Value
    //ExFor:AxisBound.ValueAsDate
    //ExSummary:Shows how to set custom axis bounds.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = chartShape->get_Chart();
    
    // Clear the chart's demo data series to start with a clean chart.
    chart->get_Series()->Clear();
    
    // Add a series with two decimal arrays. The first array contains the X-values,
    // and the second contains corresponding Y-values for points in the scatter chart.
    chart->get_Series()->Add(u"Series 1", System::MakeArray<double>({1.1, 5.4, 7.9, 3.5, 2.1, 9.7}), System::MakeArray<double>({2.1, 0.3, 0.6, 3.3, 1.4, 1.9}));
    
    // By default, default scaling is applied to the graph's X and Y-axes,
    // so that both their ranges are big enough to encompass every X and Y-value of every series.
    ASSERT_TRUE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    
    // We can define our own axis bounds.
    // In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
    chart->get_AxisX()->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(0.0));
    chart->get_AxisX()->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(10.0));
    chart->get_AxisY()->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(0.0));
    chart->get_AxisY()->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(10.0));
    
    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
    
    // Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
    chartShape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 450, 300);
    chart = chartShape->get_Chart();
    chart->get_Series()->Clear();
    
    System::ArrayPtr<System::DateTime> dates = System::MakeArray<System::DateTime>({System::DateTime(1973, 5, 11), System::DateTime(1981, 2, 4), System::DateTime(1985, 9, 23), System::DateTime(1989, 6, 28), System::DateTime(1994, 12, 15)});
    
    chart->get_Series()->Add(u"Series 1", dates, System::MakeArray<double>({3.0, 4.7, 5.9, 7.1, 8.9}));
    
    // We can set axis bounds in the form of dates as well, limiting the chart to a period.
    // Setting the range to 1980-1990 will omit the two of the series values
    // that are outside of the range from the graph.
    chart->get_AxisX()->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1980, 1, 1)));
    chart->get_AxisX()->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1990, 1, 1)));
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisBound.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.AxisBound.docx");
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(0.0, chart->get_AxisX()->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(10.0, chart->get_AxisX()->get_Scaling()->get_Maximum()->get_Value());
    
    ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(0.0, chart->get_AxisY()->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(10.0, chart->get_AxisY()->get_Scaling()->get_Maximum()->get_Value());
    
    chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 1, true)))->get_Chart();
    
    ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1980, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Minimum());
    ASPOSE_ASSERT_EQ(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(System::DateTime(1990, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Maximum());
    
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 450, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(3, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());
    
    // Move the chart's legend to the top right corner.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartLegend> legend = chart->get_Legend();
    legend->set_Position(Aspose::Words::Drawing::Charts::LegendPosition::TopRight);
    
    // Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
    legend->set_Overlay(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartLegend.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.ChartLegend.docx");
    
    legend = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_Legend();
    
    ASSERT_TRUE(legend->get_Overlay());
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 450, 250);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(3, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
    ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
    ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());
    
    // For column charts, the Y-axis crosses at zero by default,
    // which means that columns for all values below zero point down to represent negative values.
    // We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> axis = chart->get_AxisX();
    axis->set_Crosses(Aspose::Words::Drawing::Charts::AxisCrosses::Custom);
    axis->set_CrossesAt(3);
    axis->set_AxisBetweenCategories(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisCross.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.AxisCross.docx");
    axis = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart()->get_AxisX();
    
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

void ExCharts::AxisDisplayUnit()
{
    //ExStart
    //ExFor:AxisBuiltInUnit
    //ExFor:ChartAxis.DisplayUnit
    //ExFor:ChartAxis.MajorUnitIsAuto
    //ExFor:ChartAxis.MajorUnitScale
    //ExFor:ChartAxis.MinorUnitIsAuto
    //ExFor:ChartAxis.MinorUnitScale
    //ExFor:AxisDisplayUnit
    //ExFor:AxisDisplayUnit.CustomUnit
    //ExFor:AxisDisplayUnit.Unit
    //ExFor:AxisDisplayUnit.Document
    //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 450, 250);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    ASSERT_EQ(1, chart->get_Series()->get_Count());
    ASSERT_EQ(u"Y-Values", chart->get_Series()->idx_get(0)->get_Name());
    
    // Set the minor tick marks of the Y-axis to point away from the plot area,
    // and the major tick marks to cross the axis.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxis> axis = chart->get_AxisY();
    axis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Cross);
    axis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Outside);
    
    // Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
    axis->set_MajorUnit(10);
    axis->set_MinorUnit(1);
    
    // Set the Y-axis bounds to -10 and 20.
    // This Y-axis will now display 4 major tick marks and 27 minor tick marks.
    axis->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(-10.0));
    axis->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(20.0));
    
    // For the X-axis, set the major tick marks at every 10 units,
    // every minor tick mark at 2.5 units.
    axis = chart->get_AxisX();
    axis->set_MajorUnit(10);
    axis->set_MinorUnit(2.5);
    
    // Configure both types of tick marks to appear inside the graph plot area.
    axis->set_MajorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    axis->set_MinorTickMark(Aspose::Words::Drawing::Charts::AxisTickMark::Inside);
    
    // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
    axis->get_Scaling()->set_Minimum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(-10.0));
    axis->get_Scaling()->set_Maximum(System::MakeObject<Aspose::Words::Drawing::Charts::AxisBound>(30.0));
    axis->get_TickLabels()->set_Alignment(Aspose::Words::ParagraphAlignment::Right);
    
    ASSERT_EQ(1, axis->get_TickLabels()->get_Spacing());
    ASPOSE_ASSERT_EQ(doc, axis->get_DisplayUnit()->get_Document());
    
    // Set the tick labels to display their value in millions.
    axis->get_DisplayUnit()->set_Unit(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Millions);
    
    // We can set a more specific value by which tick labels will display their values.
    // This statement is equivalent to the one above.
    axis->get_DisplayUnit()->set_CustomUnit(1000000);
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
    //ExSkip
    
    doc->Save(get_ArtifactsDir() + u"Charts.AxisDisplayUnit.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.AxisDisplayUnit.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    
    ASPOSE_ASSERT_EQ(450.0, shape->get_Width());
    ASPOSE_ASSERT_EQ(250.0, shape->get_Height());
    
    axis = shape->get_Chart()->get_AxisX();
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, axis->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Inside, axis->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
    ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(30.0, axis->get_Scaling()->get_Maximum()->get_Value());
    ASSERT_EQ(1, axis->get_TickLabels()->get_Spacing());
    ASSERT_EQ(Aspose::Words::ParagraphAlignment::Right, axis->get_TickLabels()->get_Alignment());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
    ASPOSE_ASSERT_EQ(1000000.0, axis->get_DisplayUnit()->get_CustomUnit());
    
    axis = shape->get_Chart()->get_AxisY();
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Cross, axis->get_MajorTickMark());
    ASSERT_EQ(Aspose::Words::Drawing::Charts::AxisTickMark::Outside, axis->get_MinorTickMark());
    ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
    ASPOSE_ASSERT_EQ(1.0, axis->get_MinorUnit());
    ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
    ASPOSE_ASSERT_EQ(20.0, axis->get_Scaling()->get_Maximum()->get_Value());
}

namespace gtest_test
{

TEST_F(ExCharts, AxisDisplayUnit)
{
    s_instance->AxisDisplayUnit();
}

} // namespace gtest_test

void ExCharts::MarkerFormatting()
{
    //ExStart
    //ExFor:ChartDataPoint.Marker
    //ExFor:ChartMarker.Format
    //ExFor:ChartFormat.Fill
    //ExFor:ChartSeries.Marker
    //ExFor:ChartFormat.Stroke
    //ExFor:Stroke.ForeColor
    //ExFor:Stroke.BackColor
    //ExFor:Stroke.Visible
    //ExFor:Stroke.Transparency
    //ExFor:PresetTexture
    //ExFor:Fill.PresetTextured(PresetTexture)
    //ExSummary:Show how to set marker formatting.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Scatter, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"AW Series 1", System::MakeArray<double>({0.7, 1.8, 2.6, 3.9}), System::MakeArray<double>({2.7, 3.2, 0.8, 1.7}));
    
    // Set marker formatting.
    series->get_Marker()->set_Size(40);
    series->get_Marker()->set_Symbol(Aspose::Words::Drawing::Charts::MarkerSymbol::Square);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPointCollection> dataPoints = series->get_DataPoints();
    dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::Denim);
    dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
    dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Stroke()->set_BackColor(System::Drawing::Color::get_Red());
    dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::WaterDroplets);
    dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
    dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Stroke()->set_Visible(false);
    dataPoints->idx_get(2)->get_Marker()->get_Format()->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::GreenMarble);
    dataPoints->idx_get(2)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
    dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::Oak);
    dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
    dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Stroke()->set_Transparency(0.5);
    
    doc->Save(get_ArtifactsDir() + u"Charts.MarkerFormatting.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, MarkerFormatting)
{
    s_instance->MarkerFormatting();
}

} // namespace gtest_test

void ExCharts::SeriesColor()
{
    //ExStart
    //ExFor:ChartSeries.Format
    //ExSummary:Sows how to set series color.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    
    // Delete default generated series.
    seriesColl->Clear();
    
    // Create category names array.
    auto categories = System::MakeArray<System::String>({u"Category 1", u"Category 2"});
    
    // Adding new series. Value and category arrays must be the same size.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series1 = seriesColl->Add(u"Series 1", categories, System::MakeArray<double>({1, 2}));
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series2 = seriesColl->Add(u"Series 2", categories, System::MakeArray<double>({3, 4}));
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series3 = seriesColl->Add(u"Series 3", categories, System::MakeArray<double>({5, 6}));
    
    // Set series color.
    series1->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
    series2->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Yellow());
    series3->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Blue());
    
    doc->Save(get_ArtifactsDir() + u"Charts.SeriesColor.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, SeriesColor)
{
    s_instance->SeriesColor();
}

} // namespace gtest_test

void ExCharts::DataPointsFormatting()
{
    //ExStart
    //ExFor:ChartDataPoint.Format
    //ExSummary:Shows how to set individual formatting for categories of a column chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Adding new series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Series 1", System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"}), System::MakeArray<double>({1, 2, 3, 4}));
    
    // Set column formatting.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPointCollection> dataPoints = series->get_DataPoints();
    dataPoints->idx_get(0)->get_Format()->get_Fill()->PresetTextured(Aspose::Words::Drawing::PresetTexture::Denim);
    dataPoints->idx_get(1)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
    dataPoints->idx_get(2)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Yellow());
    dataPoints->idx_get(3)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Blue());
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataPointsFormatting.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, DataPointsFormatting)
{
    s_instance->DataPointsFormatting();
}

} // namespace gtest_test

void ExCharts::LegendEntries()
{
    //ExStart
    //ExFor:ChartLegendEntryCollection
    //ExFor:ChartLegend.LegendEntries
    //ExFor:ChartLegendEntry.IsHidden
    //ExSummary:Shows how to work with a legend entry for chart series.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> series = chart->get_Series();
    series->Clear();
    
    auto categories = System::MakeArray<System::String>({u"AW Category 1", u"AW Category 2"});
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series1 = series->Add(u"Series 1", categories, System::MakeArray<double>({1, 2}));
    series->Add(u"Series 2", categories, System::MakeArray<double>({3, 4}));
    series->Add(u"Series 3", categories, System::MakeArray<double>({5, 6}));
    series->Add(u"Series 4", categories, System::MakeArray<double>({0, 0}));
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartLegendEntryCollection> legendEntries = chart->get_Legend()->get_LegendEntries();
    legendEntries->idx_get(3)->set_IsHidden(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.LegendEntries.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, LegendEntries)
{
    s_instance->LegendEntries();
}

} // namespace gtest_test

void ExCharts::LegendFont()
{
    //ExStart:LegendFont
    //GistId:470c0da51e4317baae82ad9495747fed
    //ExFor:ChartLegendEntry
    //ExFor:ChartLegendEntry.Font
    //ExFor:ChartLegend.Font
    //ExFor:ChartSeries.LegendEntry
    //ExSummary:Shows how to work with a legend font.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Reporting engine template - Chart series.docx");
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartLegend> chartLegend = chart->get_Legend();
    // Set default font size all legend entries.
    chartLegend->get_Font()->set_Size(14);
    // Change font for specific legend entry.
    chartLegend->get_LegendEntries()->idx_get(1)->get_Font()->set_Italic(true);
    chartLegend->get_LegendEntries()->idx_get(1)->get_Font()->set_Size(12);
    // Get legend entry for chart series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartLegendEntry> legendEntry = chart->get_Series()->idx_get(0)->get_LegendEntry();
    
    doc->Save(get_ArtifactsDir() + u"Charts.LegendFont.docx");
    //ExEnd:LegendFont
}

namespace gtest_test
{

TEST_F(ExCharts, LegendFont)
{
    s_instance->LegendFont();
}

} // namespace gtest_test

void ExCharts::RemoveSpecificChartSeries()
{
    //ExStart
    //ExFor:ChartSeries.SeriesType
    //ExFor:ChartSeriesType
    //ExSummary:Shows how to remove specific chart serie.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Reporting engine template - Chart series.docx");
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = (System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true)))->get_Chart();
    
    // Remove all series of the Column type.
    for (int32_t i = chart->get_Series()->get_Count() - 1; i >= 0; i--)
    {
        if (chart->get_Series()->idx_get(i)->get_SeriesType() == Aspose::Words::Drawing::Charts::ChartSeriesType::Column)
        {
            chart->get_Series()->RemoveAt(i);
        }
    }
    
    chart->get_Series()->Add(u"Aspose Series", System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"}), System::MakeArray<double>({5.6, 7.1, 2.9, 8.9}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.RemoveSpecificChartSeries.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, RemoveSpecificChartSeries)
{
    s_instance->RemoveSpecificChartSeries();
}

} // namespace gtest_test

void ExCharts::PopulateChartWithData()
{
    //ExStart
    //ExFor:ChartXValue
    //ExFor:ChartXValue.FromDouble(Double)
    //ExFor:ChartYValue.FromDouble(Double)
    //ExFor:ChartSeries.Add(ChartXValue)
    //ExFor:ChartSeries.Add(ChartXValue, ChartYValue)
    //ExFor:ChartSeries.Add(ChartXValue, ChartYValue, double)
    //ExFor:ChartSeries.ClearValues
    //ExFor:ChartSeries.Clear
    //ExSummary:Shows how to populate chart series with data.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series1 = chart->get_Series()->idx_get(0);
    
    // Clear X and Y values of the first series.
    series1->ClearValues();
    
    // Populate the series with data.
    series1->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(3), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(10), 10);
    series1->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(5), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(5));
    series1->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(7), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(11));
    series1->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(9));
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series2 = chart->get_Series()->idx_get(1);
    // Clear X and Y values of the second series.
    series2->Clear();
    
    // Populate the series with data.
    series2->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(2), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(4));
    series2->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(4), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(7));
    series2->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(6), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(14));
    series2->Add(Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(8), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(7));
    
    doc->Save(get_ArtifactsDir() + u"Charts.PopulateChartWithData.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, PopulateChartWithData)
{
    s_instance->PopulateChartWithData();
}

} // namespace gtest_test

void ExCharts::GetChartSeriesData()
{
    //ExStart
    //ExFor:ChartXValueCollection
    //ExFor:ChartYValueCollection
    //ExSummary:Shows how to get chart series data.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->idx_get(0);
    
    double minValue = std::numeric_limits<double>::max();
    int32_t minValueIndex = 0;
    double maxValue = std::numeric_limits<double>::lowest();
    int32_t maxValueIndex = 0;
    
    for (int32_t i = 0; i < series->get_YValues()->get_Count(); i++)
    {
        // Clear individual format of all data points.
        // Data points and data values are one-to-one in column charts.
        series->get_DataPoints()->idx_get(i)->ClearFormat();
        
        // Get Y value.
        double yValue = series->get_YValues()->idx_get(i)->get_DoubleValue();
        
        if (yValue < minValue)
        {
            minValue = yValue;
            minValueIndex = i;
        }
        
        if (yValue > maxValue)
        {
            maxValue = yValue;
            maxValueIndex = i;
        }
    }
    
    // Change colors of the max and min values.
    series->get_DataPoints()->idx_get(minValueIndex)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
    series->get_DataPoints()->idx_get(maxValueIndex)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Green());
    
    doc->Save(get_ArtifactsDir() + u"Charts.GetChartSeriesData.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, GetChartSeriesData)
{
    s_instance->GetChartSeriesData();
}

} // namespace gtest_test

void ExCharts::ChartDataValues()
{
    //ExStart
    //ExFor:ChartXValue.FromString(String)
    //ExFor:ChartSeries.Remove(Int32)
    //ExFor:ChartSeries.Add(ChartXValue, ChartYValue)
    //ExSummary:Shows how to add/remove chart data values.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>();
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> department1Series = chart->get_Series()->idx_get(0);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> department2Series = chart->get_Series()->idx_get(1);
    
    // Remove the first value in the both series.
    department1Series->Remove(0);
    department2Series->Remove(0);
    
    // Add new values to the both series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartXValue> newXCategory = Aspose::Words::Drawing::Charts::ChartXValue::FromString(u"Q1, 2023");
    department1Series->Add(newXCategory, Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(10.3));
    department2Series->Add(newXCategory, Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(5.7));
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartDataValues.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, ChartDataValues)
{
    s_instance->ChartDataValues();
}

} // namespace gtest_test

void ExCharts::FormatDataLables()
{
    //ExStart
    //ExFor:ChartDataLabelCollection.Format
    //ExFor:ChartFormat.ShapeType
    //ExFor:ChartShapeType
    //ExSummary:Shows how to set fill, stroke and callout formatting for chart data labels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add new series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"AW Series 1", System::MakeArray<System::String>({u"AW Category 1", u"AW Category 2", u"AW Category 3", u"AW Category 4"}), System::MakeArray<double>({100, 200, 300, 400}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    series->get_DataLabels()->set_ShowValue(true);
    
    // Format data labels as callouts.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartFormat> format = series->get_DataLabels()->get_Format();
    format->set_ShapeType(Aspose::Words::Drawing::Charts::ChartShapeType::WedgeRectCallout);
    format->get_Stroke()->set_Color(System::Drawing::Color::get_DarkGreen());
    format->get_Fill()->Solid(System::Drawing::Color::get_Green());
    series->get_DataLabels()->get_Font()->set_Color(System::Drawing::Color::get_Yellow());
    
    // Change fill and stroke of an individual data label.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartFormat> labelFormat = series->get_DataLabels()->idx_get(0)->get_Format();
    labelFormat->get_Stroke()->set_Color(System::Drawing::Color::get_DarkBlue());
    labelFormat->get_Fill()->Solid(System::Drawing::Color::get_Blue());
    
    doc->Save(get_ArtifactsDir() + u"Charts.FormatDataLables.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, FormatDataLables)
{
    s_instance->FormatDataLables();
}

} // namespace gtest_test

void ExCharts::ChartAxisTitle()
{
    //ExStart:ChartAxisTitle
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ChartAxis.Title
    //ExFor:ChartAxisTitle
    //ExFor:ChartAxisTitle.Text
    //ExFor:ChartAxisTitle.Show
    //ExFor:ChartAxisTitle.Overlay
    //ExFor:ChartAxisTitle.Font
    //ExSummary:Shows how to set chart axis title.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    // Delete default generated series.
    seriesColl->Clear();
    
    seriesColl->Add(u"AW Series 1", System::MakeArray<System::String>({u"AW Category 1", u"AW Category 2"}), System::MakeArray<double>({1, 2}));
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxisTitle> chartAxisXTitle = chart->get_AxisX()->get_Title();
    chartAxisXTitle->set_Text(u"Categories");
    chartAxisXTitle->set_Show(true);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartAxisTitle> chartAxisYTitle = chart->get_AxisY()->get_Title();
    chartAxisYTitle->set_Text(u"Values");
    chartAxisYTitle->set_Show(true);
    chartAxisYTitle->set_Overlay(true);
    chartAxisYTitle->get_Font()->set_Size(12);
    chartAxisYTitle->get_Font()->set_Color(System::Drawing::Color::get_Blue());
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartAxisTitle.docx");
    //ExEnd:ChartAxisTitle
}

namespace gtest_test
{

TEST_F(ExCharts, ChartAxisTitle)
{
    s_instance->ChartAxisTitle();
}

} // namespace gtest_test

void ExCharts::CopyDataPointFormat()
{
    //ExStart:CopyDataPointFormat
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ChartSeries.CopyFormatFrom(int)
    //ExFor:ChartDataPointCollection.HasDefaultFormat(int)
    //ExFor:ChartDataPointCollection.CopyFormat(int, int)
    //ExSummary:Shows how to copy data point format.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DataPoint format.docx");
    
    // Get the chart and series to update format.
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = shape->get_Chart()->get_Series()->idx_get(0);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPointCollection> dataPoints = series->get_DataPoints();
    
    ASSERT_TRUE(dataPoints->HasDefaultFormat(0));
    ASSERT_FALSE(dataPoints->HasDefaultFormat(1));
    
    // Copy format of the data point with index 1 to the data point with index 2
    // so that the data point 2 looks the same as the data point 1.
    dataPoints->CopyFormat(0, 1);
    
    ASSERT_TRUE(dataPoints->HasDefaultFormat(0));
    ASSERT_TRUE(dataPoints->HasDefaultFormat(1));
    
    // Copy format of the data point with index 0 to the series defaults so that all data points
    // in the series that have the default format look the same as the data point 0.
    series->CopyFormatFrom(1);
    
    ASSERT_TRUE(dataPoints->HasDefaultFormat(0));
    ASSERT_TRUE(dataPoints->HasDefaultFormat(1));
    
    doc->Save(get_ArtifactsDir() + u"Charts.CopyDataPointFormat.docx");
    //ExEnd:CopyDataPointFormat
}

namespace gtest_test
{

TEST_F(ExCharts, CopyDataPointFormat)
{
    s_instance->CopyDataPointFormat();
}

} // namespace gtest_test

void ExCharts::ResetDataPointFill()
{
    //ExStart:ResetDataPointFill
    //GistId:3428e84add5beb0d46a8face6e5fc858
    //ExFor:ChartFormat.IsDefined
    //ExFor:ChartFormat.SetDefaultFill
    //ExSummary:Shows how to reset the fill to the default value defined in the series.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"DataPoint format.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = shape->get_Chart()->get_Series()->idx_get(0);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataPoint> dataPoint = series->get_DataPoints()->idx_get(1);
    
    ASSERT_TRUE(dataPoint->get_Format()->get_IsDefined());
    
    dataPoint->get_Format()->SetDefaultFill();
    
    doc->Save(get_ArtifactsDir() + u"Charts.ResetDataPointFill.docx");
    //ExEnd:ResetDataPointFill
}

namespace gtest_test
{

TEST_F(ExCharts, ResetDataPointFill)
{
    s_instance->ResetDataPointFill();
}

} // namespace gtest_test

void ExCharts::DataTable()
{
    //ExStart:DataTable
    //GistId:a775441ecb396eea917a2717cb9e8f8f
    //ExFor:Chart.DataTable
    //ExFor:ChartDataTable
    //ExFor:ChartDataTable.Show
    //ExFor:ChartDataTable.Format
    //ExFor:ChartDataTable.Font
    //ExFor:ChartDataTable.HasLegendKeys
    //ExFor:ChartDataTable.HasHorizontalBorder
    //ExFor:ChartDataTable.HasVerticalBorder
    //ExFor:ChartDataTable.HasOutlineBorder
    //ExSummary:Shows how to show data table with chart series data.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> series = chart->get_Series();
    series->Clear();
    auto xValues = System::MakeArray<double>({2020, 2021, 2022, 2023});
    series->Add(u"Series1", xValues, System::MakeArray<double>({5, 11, 2, 7}));
    series->Add(u"Series2", xValues, System::MakeArray<double>({6, 5.5, 7, 7.8}));
    series->Add(u"Series3", xValues, System::MakeArray<double>({10, 8, 7, 9}));
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataTable> dataTable = chart->get_DataTable();
    dataTable->set_Show(true);
    
    dataTable->set_HasLegendKeys(false);
    dataTable->set_HasHorizontalBorder(false);
    dataTable->set_HasVerticalBorder(false);
    dataTable->set_HasOutlineBorder(false);
    
    dataTable->get_Font()->set_Italic(true);
    dataTable->get_Format()->get_Stroke()->set_Weight(1);
    dataTable->get_Format()->get_Stroke()->set_DashStyle(Aspose::Words::Drawing::DashStyle::ShortDot);
    dataTable->get_Format()->get_Stroke()->set_Color(System::Drawing::Color::get_DarkBlue());
    
    doc->Save(get_ArtifactsDir() + u"Charts.DataTable.docx");
    //ExEnd:DataTable
}

namespace gtest_test
{

TEST_F(ExCharts, DataTable)
{
    s_instance->DataTable();
}

} // namespace gtest_test

void ExCharts::ChartFormat()
{
    //ExStart:ChartFormat
    //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
    //ExFor:ChartFormat
    //ExFor:Chart.Format
    //ExFor:ChartTitle.Format
    //ExFor:ChartAxisTitle.Format
    //ExFor:ChartLegend.Format
    //ExFor:Fill.Solid(Color)
    //ExSummary:Shows how to use chart formating.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Delete series generated by default.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> series = chart->get_Series();
    series->Clear();
    
    auto categories = System::MakeArray<System::String>({u"Category 1", u"Category 2"});
    series->Add(u"Series 1", categories, System::MakeArray<double>({1, 2}));
    series->Add(u"Series 2", categories, System::MakeArray<double>({3, 4}));
    
    // Format chart background.
    chart->get_Format()->get_Fill()->Solid(System::Drawing::Color::get_DarkSlateGray());
    
    // Hide axis tick labels.
    chart->get_AxisX()->get_TickLabels()->set_Position(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::None);
    chart->get_AxisY()->get_TickLabels()->set_Position(Aspose::Words::Drawing::Charts::AxisTickLabelPosition::None);
    
    // Format chart title.
    chart->get_Title()->get_Format()->get_Fill()->Solid(System::Drawing::Color::get_LightGoldenrodYellow());
    
    // Format axis title.
    chart->get_AxisX()->get_Title()->set_Show(true);
    chart->get_AxisX()->get_Title()->get_Format()->get_Fill()->Solid(System::Drawing::Color::get_LightGoldenrodYellow());
    
    // Format legend.
    chart->get_Legend()->get_Format()->get_Fill()->Solid(System::Drawing::Color::get_LightGoldenrodYellow());
    
    doc->Save(get_ArtifactsDir() + u"Charts.ChartFormat.docx");
    //ExEnd:ChartFormat
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.ChartFormat.docx");
    
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    chart = shape->get_Chart();
    
    ASSERT_EQ(System::Drawing::Color::get_DarkSlateGray().ToArgb(), chart->get_Format()->get_Fill()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightGoldenrodYellow().ToArgb(), chart->get_Title()->get_Format()->get_Fill()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightGoldenrodYellow().ToArgb(), chart->get_AxisX()->get_Title()->get_Format()->get_Fill()->get_Color().ToArgb());
    ASSERT_EQ(System::Drawing::Color::get_LightGoldenrodYellow().ToArgb(), chart->get_Legend()->get_Format()->get_Fill()->get_Color().ToArgb());
}

namespace gtest_test
{

TEST_F(ExCharts, ChartFormat)
{
    s_instance->ChartFormat();
}

} // namespace gtest_test

void ExCharts::SecondaryAxis()
{
    //ExStart:SecondaryAxis
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ChartSeriesGroup
    //ExFor:ChartSeriesGroup.SeriesType
    //ExFor:ChartSeriesGroup.AxisGroup
    //ExFor:ChartSeriesGroup.AxisX
    //ExFor:ChartSeriesGroup.AxisY
    //ExFor:ChartSeriesGroup.Series
    //ExFor:ChartSeriesGroupCollection
    //ExFor:ChartSeriesGroupCollection.Add(ChartSeriesType)
    //ExFor:AxisGroup
    //ExSummary:Shows how to work with the secondary axis of chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 450, 250);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> series = chart->get_Series();
    
    // Delete default generated series.
    series->Clear();
    
    auto categories = System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3"});
    series->Add(u"Series 1 of primary series group", categories, System::MakeArray<double>({2, 3, 4}));
    series->Add(u"Series 2 of primary series group", categories, System::MakeArray<double>({5, 2, 3}));
    
    // Create an additional series group, also of the line type.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroup> newSeriesGroup = chart->get_SeriesGroups()->Add(Aspose::Words::Drawing::Charts::ChartSeriesType::Line);
    // Specify the use of secondary axes for the new series group.
    newSeriesGroup->set_AxisGroup(Aspose::Words::Drawing::Charts::AxisGroup::Secondary);
    // Hide the secondary X axis.
    newSeriesGroup->get_AxisX()->set_Hidden(true);
    // Define title of the secondary Y axis.
    newSeriesGroup->get_AxisY()->get_Title()->set_Show(true);
    newSeriesGroup->get_AxisY()->get_Title()->set_Text(u"Secondary Y axis");
    
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartSeriesType::Line, newSeriesGroup->get_SeriesType());
    
    // Add a series to the new series group.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series3 = newSeriesGroup->get_Series()->Add(u"Series of secondary series group", categories, System::MakeArray<double>({13, 11, 16}));
    series3->get_Format()->get_Stroke()->set_Weight(3.5);
    
    doc->Save(get_ArtifactsDir() + u"Charts.SecondaryAxis.docx");
    //ExEnd:SecondaryAxis
}

namespace gtest_test
{

TEST_F(ExCharts, SecondaryAxis)
{
    s_instance->SecondaryAxis();
}

} // namespace gtest_test

void ExCharts::ConfigureGapOverlap()
{
    //ExStart:ConfigureGapOverlap
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:Chart.SeriesGroups
    //ExFor:ChartSeriesGroup.GapWidth
    //ExFor:ChartSeriesGroup.Overlap
    //ExSummary:Show how to configure gap width and overlap.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 450, 250);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroup> seriesGroup = shape->get_Chart()->get_SeriesGroups()->idx_get(0);
    
    // Set column gap width and overlap.
    seriesGroup->set_GapWidth(450);
    seriesGroup->set_Overlap(-75);
    
    doc->Save(get_ArtifactsDir() + u"Charts.ConfigureGapOverlap.docx");
    //ExEnd:ConfigureGapOverlap
}

namespace gtest_test
{

TEST_F(ExCharts, ConfigureGapOverlap)
{
    s_instance->ConfigureGapOverlap();
}

} // namespace gtest_test

void ExCharts::BubbleScale()
{
    //ExStart:BubbleScale
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ChartSeriesGroup.BubbleScale
    //ExSummary:Show how to set size of the bubbles.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a bubble 3D chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble3D, 450, 250);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroup> seriesGroup = shape->get_Chart()->get_SeriesGroups()->idx_get(0);
    
    // Set bubble scale to 200%.
    seriesGroup->set_BubbleScale(200);
    
    doc->Save(get_ArtifactsDir() + u"Charts.BubbleScale.docx");
    //ExEnd:BubbleScale
}

namespace gtest_test
{

TEST_F(ExCharts, BubbleScale)
{
    s_instance->BubbleScale();
}

} // namespace gtest_test

void ExCharts::RemoveSecondaryAxis()
{
    //ExStart:RemoveSecondaryAxis
    //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
    //ExFor:ChartSeriesGroupCollection.Count
    //ExFor:ChartSeriesGroupCollection.Item(Int32)
    //ExFor:ChartSeriesGroupCollection.RemoveAt(Int32)
    //ExSummary:Show how to remove secondary axis.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Combo chart.docx");
    
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroupCollection> seriesGroups = chart->get_SeriesGroups();
    
    // Find secondary axis and remove from the collection.
    for (int32_t i = 0; i < seriesGroups->get_Count(); i++)
    {
        if (seriesGroups->idx_get(i)->get_AxisGroup() == Aspose::Words::Drawing::Charts::AxisGroup::Secondary)
        {
            seriesGroups->RemoveAt(i);
        }
    }
    //ExEnd:RemoveSecondaryAxis
}

namespace gtest_test
{

TEST_F(ExCharts, RemoveSecondaryAxis)
{
    s_instance->RemoveSecondaryAxis();
}

} // namespace gtest_test

void ExCharts::TreemapChart()
{
    //ExStart:TreemapChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, ChartMultilevelValue[], double[])
    //ExFor:ChartMultilevelValue
    //ExFor:ChartMultilevelValue.#ctor(String, String, String)
    //ExFor:ChartMultilevelValue.#ctor(String, String)
    //ExFor:ChartMultilevelValue.#ctor(String)
    //ExSummary:Shows how to create treemap chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Treemap chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Treemap, 450, 280);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"World Population");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Population by Region", System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartMultilevelValue>>({System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"China"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"India"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Indonesia"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Pakistan"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Bangladesh"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Japan"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Philippines"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Asia", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Africa", u"Nigeria"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Africa", u"Ethiopia"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Africa", u"Egypt"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Africa", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Europe", u"Russia"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Europe", u"Germany"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Europe", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Latin America", u"Brazil"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Latin America", u"Mexico"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Latin America", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Northern America", u"United States", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Northern America", u"Other"), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Oceania")}), System::MakeArray<double>({1409670000, 1400744000, 279118866, 241499431, 169828911, 123930000, 112892781, 764000000, 223800000, 107334000, 105914499, 903000000, 146150789, 84607016, 516000000, 203080756, 129713690, 310000000, 335893238, 35000000, 42000000}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    series->get_DataLabels()->set_ShowValue(true);
    series->get_DataLabels()->set_ShowCategoryName(true);
    System::String thousandSeparator = System::Globalization::CultureInfo::get_CurrentCulture()->get_NumberFormat()->get_CurrencyGroupSeparator();
    series->get_DataLabels()->get_NumberFormat()->set_FormatCode(System::String::Format(u"#{0}0", thousandSeparator));
    
    doc->Save(get_ArtifactsDir() + u"Charts.Treemap.docx");
    //ExEnd:TreemapChart
}

namespace gtest_test
{

TEST_F(ExCharts, TreemapChart)
{
    s_instance->TreemapChart();
}

} // namespace gtest_test

void ExCharts::SunburstChart()
{
    //ExStart:SunburstChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, ChartMultilevelValue[], double[])
    //ExSummary:Shows how to create sunburst chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Sunburst chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Sunburst, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"Sales");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Sales", System::MakeArray<System::SharedPtr<Aspose::Words::Drawing::Charts::ChartMultilevelValue>>({System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Europe", u"UK", u"London Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Europe", u"UK", u"Liverpool Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Europe", u"UK", u"Manchester Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Europe", u"France", u"Paris Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Europe", u"France", u"Lyon Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"USA", u"Denver Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"USA", u"Seattle Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"USA", u"Detroit Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"USA", u"Houston Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"Canada", u"Toronto Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - NA", u"Canada", u"Montreal Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Oceania", u"Australia", u"Sydney Dep."), System::MakeObject<Aspose::Words::Drawing::Charts::ChartMultilevelValue>(u"Sales - Oceania", u"New Zealand", u"Auckland Dep.")}), System::MakeArray<double>({1236, 851, 536, 468, 179, 527, 799, 1148, 921, 457, 482, 761, 694}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    series->get_DataLabels()->set_ShowValue(false);
    series->get_DataLabels()->set_ShowCategoryName(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.Sunburst.docx");
    //ExEnd:SunburstChart
}

namespace gtest_test
{

TEST_F(ExCharts, SunburstChart)
{
    s_instance->SunburstChart();
}

} // namespace gtest_test

void ExCharts::HistogramChart()
{
    //ExStart:HistogramChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, double[])
    //ExSummary:Shows how to create histogram chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Histogram chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Histogram, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"Avg Temperature since 1991");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    chart->get_Series()->Add(u"Avg Temperature", System::MakeArray<double>({51.8, 53.6, 50.3, 54.7, 53.9, 54.3, 53.4, 52.9, 53.3, 53.7, 53.8, 52.0, 55.0, 52.1, 53.4, 53.8, 53.8, 51.9, 52.1, 52.7, 51.8, 56.6, 53.3, 55.6, 56.3, 56.2, 56.1, 56.2, 53.6, 55.7, 56.3, 55.9, 55.6}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.Histogram.docx");
    //ExEnd:HistogramChart
}

namespace gtest_test
{

TEST_F(ExCharts, HistogramChart)
{
    s_instance->HistogramChart();
}

} // namespace gtest_test

void ExCharts::ParetoChart()
{
    //ExStart:ParetoChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, String[], double[])
    //ExSummary:Shows how to create pareto chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Pareto chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Pareto, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"Best-Selling Car");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    chart->get_Series()->Add(u"Best-Selling Car", System::MakeArray<System::String>({u"Tesla Model Y", u"Toyota Corolla", u"Toyota RAV4", u"Ford F-Series", u"Honda CR-V"}), System::MakeArray<double>({1.43, 0.91, 1.17, 0.98, 0.85}));
    
    doc->Save(get_ArtifactsDir() + u"Charts.Pareto.docx");
    //ExEnd:ParetoChart
}

namespace gtest_test
{

TEST_F(ExCharts, ParetoChart)
{
    s_instance->ParetoChart();
}

} // namespace gtest_test

void ExCharts::BoxAndWhiskerChart()
{
    //ExStart:BoxAndWhiskerChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, String[], double[])
    //ExSummary:Shows how to create box and whisker chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Box & Whisker chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::BoxAndWhisker, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"Points by Years");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Points by Years", System::MakeArray<System::String>({u"WC", u"WC", u"WC", u"WC", u"WC", u"WC", u"WC", u"WC", u"WC", u"WC", u"NR", u"NR", u"NR", u"NR", u"NR", u"NR", u"NR", u"NR", u"NR", u"NR", u"NA", u"NA", u"NA", u"NA", u"NA", u"NA", u"NA", u"NA", u"NA", u"NA"}), System::MakeArray<double>({91, 80, 100, 77, 90, 104, 105, 118, 120, 101, 114, 107, 110, 60, 79, 78, 77, 102, 101, 113, 94, 93, 84, 71, 80, 103, 80, 94, 100, 101}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.BoxAndWhisker.docx");
    //ExEnd:BoxAndWhiskerChart
}

namespace gtest_test
{

TEST_F(ExCharts, BoxAndWhiskerChart)
{
    s_instance->BoxAndWhiskerChart();
}

} // namespace gtest_test

void ExCharts::WaterfallChart()
{
    //ExStart:WaterfallChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, String[], double[], bool[])
    //ExSummary:Shows how to create waterfall chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Waterfall chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Waterfall, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"New Zealand GDP");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"New Zealand GDP", System::MakeArray<System::String>({u"2018", u"2019 growth", u"2020 growth", u"2020", u"2021 growth", u"2022 growth", u"2022"}), System::MakeArray<double>({100, 0.57, -0.25, 100.32, 20.22, -2.92, 117.62}), System::MakeArray<bool>({true, false, false, true, false, false, true}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    
    doc->Save(get_ArtifactsDir() + u"Charts.Waterfall.docx");
    //ExEnd:WaterfallChart
}

namespace gtest_test
{

TEST_F(ExCharts, WaterfallChart)
{
    s_instance->WaterfallChart();
}

} // namespace gtest_test

void ExCharts::FunnelChart()
{
    //ExStart:FunnelChart
    //GistId:65919861586e42e24f61a3ccb65f8f4e
    //ExFor:ChartSeriesCollection.Add(String, String[], double[])
    //ExSummary:Shows how to create funnel chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Funnel chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Funnel, 450, 450);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    chart->get_Title()->set_Text(u"Population by Age Group");
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    // Add a series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Population by Age Group", System::MakeArray<System::String>({u"0-9", u"10-19", u"20-29", u"30-39", u"40-49", u"50-59", u"60-69", u"70-79", u"80-89", u"90-"}), System::MakeArray<double>({0.121, 0.128, 0.132, 0.146, 0.124, 0.124, 0.111, 0.075, 0.032, 0.007}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    System::String decimalSeparator = System::Globalization::CultureInfo::get_CurrentCulture()->get_NumberFormat()->get_CurrencyDecimalSeparator();
    series->get_DataLabels()->get_NumberFormat()->set_FormatCode(System::String::Format(u"0{0}0%", decimalSeparator));
    
    doc->Save(get_ArtifactsDir() + u"Charts.Funnel.docx");
    //ExEnd:FunnelChart
}

namespace gtest_test
{

TEST_F(ExCharts, FunnelChart)
{
    s_instance->FunnelChart();
}

} // namespace gtest_test

void ExCharts::LabelOrientationRotation()
{
    //ExStart:LabelOrientationRotation
    //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
    //ExFor:ChartDataLabelCollection.Orientation
    //ExFor:ChartDataLabelCollection.Rotation
    //ExFor:ChartDataLabel.Rotation
    //ExFor:ChartDataLabel.Orientation
    //ExFor:ShapeTextOrientation
    //ExSummary:Shows how to change orientation and rotation for data labels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = shape->get_Chart()->get_Series()->idx_get(0);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    
    // Show data labels.
    series->set_HasDataLabels(true);
    dataLabels->set_ShowValue(true);
    dataLabels->set_ShowCategoryName(true);
    
    // Define data label shape.
    dataLabels->get_Format()->set_ShapeType(Aspose::Words::Drawing::Charts::ChartShapeType::UpArrow);
    dataLabels->get_Format()->get_Stroke()->get_Fill()->Solid(System::Drawing::Color::get_DarkBlue());
    
    // Set data label orientation and rotation for the entire series.
    dataLabels->set_Orientation(Aspose::Words::Drawing::ShapeTextOrientation::VerticalFarEast);
    dataLabels->set_Rotation(-45);
    
    // Change orientation and rotation of the first data label.
    dataLabels->idx_get(0)->set_Orientation(Aspose::Words::Drawing::ShapeTextOrientation::Horizontal);
    dataLabels->idx_get(0)->set_Rotation(45);
    
    doc->Save(get_ArtifactsDir() + u"Charts.LabelOrientationRotation.docx");
    //ExEnd:LabelOrientationRotation
}

namespace gtest_test
{

TEST_F(ExCharts, LabelOrientationRotation)
{
    s_instance->LabelOrientationRotation();
}

} // namespace gtest_test

void ExCharts::TickLabelsOrientationRotation()
{
    //ExStart:TickLabelsOrientationRotation
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:AxisTickLabels.Rotation
    //ExFor:AxisTickLabels.Orientation
    //ExSummary:Shows how to change orientation and rotation for axis tick labels.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a column chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::AxisTickLabels> xTickLabels = shape->get_Chart()->get_AxisX()->get_TickLabels();
    System::SharedPtr<Aspose::Words::Drawing::Charts::AxisTickLabels> yTickLabels = shape->get_Chart()->get_AxisY()->get_TickLabels();
    
    // Set axis tick label orientation and rotation.
    xTickLabels->set_Orientation(Aspose::Words::Drawing::ShapeTextOrientation::VerticalFarEast);
    xTickLabels->set_Rotation(-30);
    yTickLabels->set_Orientation(Aspose::Words::Drawing::ShapeTextOrientation::Horizontal);
    yTickLabels->set_Rotation(45);
    
    doc->Save(get_ArtifactsDir() + u"Charts.TickLabelsOrientationRotation.docx");
    //ExEnd:TickLabelsOrientationRotation
}

namespace gtest_test
{

TEST_F(ExCharts, TickLabelsOrientationRotation)
{
    s_instance->TickLabelsOrientationRotation();
}

} // namespace gtest_test

void ExCharts::DoughnutChart()
{
    //ExStart:DoughnutChart
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:ChartSeriesGroup.DoughnutHoleSize
    //ExFor:ChartSeriesGroup.FirstSliceAngle
    //ExSummary:Shows how to create and format Doughnut chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Doughnut, 400, 400);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    // Delete the default generated series.
    chart->get_Series()->Clear();
    
    auto categories = System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3"});
    chart->get_Series()->Add(u"Series 1", categories, System::MakeArray<double>({4, 2, 5}));
    
    // Format the Doughnut chart.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroup> seriesGroup = chart->get_SeriesGroups()->idx_get(0);
    seriesGroup->set_DoughnutHoleSize(10);
    seriesGroup->set_FirstSliceAngle(270);
    
    doc->Save(get_ArtifactsDir() + u"Charts.DoughnutChart.docx");
    //ExEnd:DoughnutChart
}

namespace gtest_test
{

TEST_F(ExCharts, DoughnutChart)
{
    s_instance->DoughnutChart();
}

} // namespace gtest_test

void ExCharts::PieOfPieChart()
{
    //ExStart:PieOfPieChart
    //GistId:bb594993b5fe48692541e16f4d354ac2
    //ExFor:ChartSeriesGroup.SecondSectionSize
    //ExSummary:Shows how to create and format pie of Pie chart.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::PieOfPie, 440, 300);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    // Delete the default generated series.
    chart->get_Series()->Clear();
    
    auto categories = System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"});
    chart->get_Series()->Add(u"Series 1", categories, System::MakeArray<double>({11, 8, 4, 3}));
    
    // Format the Pie of Pie chart.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesGroup> seriesGroup = chart->get_SeriesGroups()->idx_get(0);
    seriesGroup->set_GapWidth(10);
    seriesGroup->set_SecondSectionSize(77);
    
    doc->Save(get_ArtifactsDir() + u"Charts.PieOfPieChart.docx");
    //ExEnd:PieOfPieChart
}

namespace gtest_test
{

TEST_F(ExCharts, PieOfPieChart)
{
    s_instance->PieOfPieChart();
}

} // namespace gtest_test

void ExCharts::FormatCode()
{
    //ExStart:FormatCode
    //GistId:366eb64fd56dec3c2eaa40410e594182
    //ExFor:ChartXValueCollection.FormatCode
    //ExFor:ChartYValueCollection.FormatCode
    //ExFor:BubbleSizeCollection.FormatCode
    //ExFor:ChartSeries.BubbleSizes
    //ExFor:ChartSeries.XValues
    //ExFor:ChartSeries.YValues
    //ExSummary:Shows how to work with the format code of the chart data.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a Bubble chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Bubble, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Delete default generated series.
    chart->get_Series()->Clear();
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = chart->get_Series()->Add(u"Series1", System::MakeArray<double>({1, 1.9, 2.45, 3}), System::MakeArray<double>({1, -0.9, 1.82, 0}), System::MakeArray<double>({2, 1.1, 2.95, 2}));
    
    // Show data labels.
    series->set_HasDataLabels(true);
    series->get_DataLabels()->set_ShowCategoryName(true);
    series->get_DataLabels()->set_ShowValue(true);
    series->get_DataLabels()->set_ShowBubbleSize(true);
    
    // Set data format codes.
    series->get_XValues()->set_FormatCode(u"#,##0.0#");
    series->get_YValues()->set_FormatCode(u"#,##0.0#;[Red]\\-#,##0.0#");
    series->get_BubbleSizes()->set_FormatCode(u"#,##0.0#");
    
    doc->Save(get_ArtifactsDir() + u"Charts.FormatCode.docx");
    //ExEnd:FormatCode
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.FormatCode.docx");
    shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    chart = shape->get_Chart();
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesCollection = chart->get_Series();
    for (auto&& seriesProperties : System::IterateOver(seriesCollection))
    {
        ASSERT_EQ(u"#,##0.0#", seriesProperties->get_XValues()->get_FormatCode());
        ASSERT_EQ(u"#,##0.0#;[Red]\\-#,##0.0#", seriesProperties->get_YValues()->get_FormatCode());
        ASSERT_EQ(u"#,##0.0#", seriesProperties->get_BubbleSizes()->get_FormatCode());
    }
}

namespace gtest_test
{

TEST_F(ExCharts, FormatCode)
{
    s_instance->FormatCode();
}

} // namespace gtest_test

void ExCharts::DataLablePosition()
{
    //ExStart:DataLablePosition
    //GistId:695136dbbe4f541a8a0a17b3d3468689
    //ExFor:ChartDataLabelCollection.Position
    //ExFor:ChartDataLabel.Position
    //ExFor:ChartDataLabelPosition
    //ExSummary:Shows how to set the position of the data label.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert column chart.
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    
    // Delete default generated series.
    seriesColl->Clear();
    
    // Add series.
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = seriesColl->Add(u"Series 1", System::MakeArray<System::String>({u"Category 1", u"Category 2", u"Category 3"}), System::MakeArray<double>({4, 5, 6}));
    
    // Show data labels and set font color.
    series->set_HasDataLabels(true);
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    dataLabels->set_ShowValue(true);
    dataLabels->get_Font()->set_Color(System::Drawing::Color::get_White());
    
    // Set data label position.
    dataLabels->set_Position(Aspose::Words::Drawing::Charts::ChartDataLabelPosition::InsideBase);
    dataLabels->idx_get(0)->set_Position(Aspose::Words::Drawing::Charts::ChartDataLabelPosition::OutsideEnd);
    dataLabels->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_DarkRed());
    
    doc->Save(get_ArtifactsDir() + u"Charts.LabelPosition.docx");
    //ExEnd:DataLablePosition
}

namespace gtest_test
{

TEST_F(ExCharts, DataLablePosition)
{
    s_instance->DataLablePosition();
}

} // namespace gtest_test

void ExCharts::DoughnutChartLabelPosition()
{
    //ExStart:DoughnutChartLabelPosition
    //GistId:695136dbbe4f541a8a0a17b3d3468689
    //ExFor:ChartDataLabel.Left
    //ExFor:ChartDataLabel.LeftMode
    //ExFor:ChartDataLabel.Top
    //ExFor:ChartDataLabel.TopMode
    //ExFor:ChartDataLabelLocationMode
    //ExSummary:Shows how to place data labels of doughnut chart outside doughnut.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    const int32_t chartWidth = 432;
    const int32_t chartHeight = 252;
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Doughnut, chartWidth, chartHeight);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeriesCollection> seriesColl = chart->get_Series();
    // Delete default generated series.
    seriesColl->Clear();
    
    // Hide the legend.
    chart->get_Legend()->set_Position(Aspose::Words::Drawing::Charts::LegendPosition::None);
    
    // Generate data.
    const int32_t dataLength = 20;
    double totalValue = 0;
    auto categories = System::MakeArray<System::String>(dataLength);
    auto values = System::MakeArray<double>(dataLength, 0);
    
    for (int32_t i = 0; i < dataLength; i++)
    {
        categories[i] = System::String::Format(u"Category {0}", i);
        values[i] = dataLength - i;
        totalValue = totalValue + values[i];
    }
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series = seriesColl->Add(u"Series 1", categories, values);
    series->set_HasDataLabels(true);
    
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabelCollection> dataLabels = series->get_DataLabels();
    dataLabels->set_ShowValue(true);
    dataLabels->set_ShowLeaderLines(true);
    
    // The Position property cannot be used for doughnut charts. Let's place data labels using the Left and Top
    // properties around a circle outside of the chart doughnut.
    // The origin is in the upper left corner of the chart.
    
    const double titleAreaHeight = 25.5;
    // This can be calculated using title text and font.
    const double doughnutCenterY = titleAreaHeight + (chartHeight - titleAreaHeight) / 2;
    const double doughnutCenterX = chartWidth / 2.0;
    const double labelHeight = 16.5;
    // This can be calculated using label font.
    const double oneCharLabelWidth = 12.75;
    // This can be calculated for each label using its text and font.
    const double twoCharLabelWidth = 17.25;
    // This can be calculated for each label using its text and font.
    const double yMargin = 0.75;
    const double labelMargin = 1.5;
    const double labelCircleRadius = chartHeight - doughnutCenterY - yMargin - labelHeight / 2;
    
    // Because the data points start at the top, the X coordinates used in the Left and Top properties of
    // the data labels point to the right and the Y coordinates point down, the starting angle is -PI/2.
    double totalAngle = -System::Math::PI / 2;
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabel> previousLabel;
    
    for (int32_t i = 0; i < series->get_YValues()->get_Count(); i++)
    {
        System::SharedPtr<Aspose::Words::Drawing::Charts::ChartDataLabel> dataLabel = dataLabels->idx_get(i);
        
        double value = series->get_YValues()->idx_get(i)->get_DoubleValue();
        double labelWidth;
        if (value < 10)
        {
            labelWidth = oneCharLabelWidth;
        }
        else
        {
            labelWidth = twoCharLabelWidth;
        }
        double labelSegmentAngle = value / totalValue * 2 * System::Math::PI;
        double labelAngle = labelSegmentAngle / 2 + totalAngle;
        double labelCenterX = labelCircleRadius * System::Math::Cos(labelAngle) + doughnutCenterX;
        double labelCenterY = labelCircleRadius * System::Math::Sin(labelAngle) + doughnutCenterY;
        double labelLeft = labelCenterX - labelWidth / 2;
        double labelTop = labelCenterY - labelHeight / 2;
        
        // If the current data label overlaps other labels, move it horizontally.
        if ((previousLabel != nullptr) && (System::Math::Abs(previousLabel->get_Top() - labelTop) < labelHeight) && (System::Math::Abs(previousLabel->get_Left() - labelLeft) < labelWidth))
        {
            // Move right on the top, left on the bottom.
            bool isOnTop = (totalAngle < 0) || (totalAngle >= System::Math::PI);
            int32_t factor;
            if (isOnTop)
            {
                factor = 1;
            }
            else
            {
                factor = -1;
            }
            
            labelLeft = previousLabel->get_Left() + labelWidth * factor + labelMargin;
        }
        
        dataLabel->set_Left(labelLeft);
        dataLabel->set_LeftMode(Aspose::Words::Drawing::Charts::ChartDataLabelLocationMode::Absolute);
        dataLabel->set_Top(labelTop);
        dataLabel->set_TopMode(Aspose::Words::Drawing::Charts::ChartDataLabelLocationMode::Absolute);
        
        totalAngle = totalAngle + labelSegmentAngle;
        previousLabel = dataLabel;
    }
    
    doc->Save(get_ArtifactsDir() + u"Charts.DoughnutChartLabelPosition.docx");
    //ExEnd:DoughnutChartLabelPosition
}

namespace gtest_test
{

TEST_F(ExCharts, DoughnutChartLabelPosition)
{
    s_instance->DoughnutChartLabelPosition();
}

} // namespace gtest_test

void ExCharts::InsertChartSeries()
{
    //ExStart
    //ExFor:ChartSeries.Insert(Int32, ChartXValue)
    //ExFor:ChartSeries.Insert(Int32, ChartXValue, ChartYValue)
    //ExFor:ChartSeries.Insert(Int32, ChartXValue, ChartYValue, double)
    //ExSummary:Shows how to insert data into a chart series.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    System::SharedPtr<Aspose::Words::Drawing::Shape> shape = builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Line, 432, 252);
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series1 = chart->get_Series()->idx_get(0);
    
    // Clear X and Y values of the first series.
    series1->ClearValues();
    // Populate the series with data.
    series1->Insert(0, Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(3));
    series1->Insert(1, Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(3), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(10));
    series1->Insert(2, Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(3), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(10));
    series1->Insert(3, Aspose::Words::Drawing::Charts::ChartXValue::FromDouble(3), Aspose::Words::Drawing::Charts::ChartYValue::FromDouble(10), 10);
    
    doc->Save(get_ArtifactsDir() + u"Charts.PopulateChartWithData.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, InsertChartSeries)
{
    s_instance->InsertChartSeries();
}

} // namespace gtest_test

void ExCharts::SetChartStyle()
{
    //ExStart
    //ExFor:ChartStyle
    //ExSummary:Shows how to set and get chart style.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Insert a chart in the Black style.
    builder->InsertChart(Aspose::Words::Drawing::Charts::ChartType::Column, 400, 250, Aspose::Words::Drawing::Charts::ChartStyle::Black);
    
    doc->Save(get_ArtifactsDir() + u"Charts.SetChartStyle.docx");
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Charts.SetChartStyle.docx");
    
    // Get a chart to update.
    auto shape = System::ExplicitCast<Aspose::Words::Drawing::Shape>(doc->GetChild(Aspose::Words::NodeType::Shape, 0, true));
    System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> chart = shape->get_Chart();
    
    // Get the chart style.
    ASSERT_EQ(Aspose::Words::Drawing::Charts::ChartStyle::Black, chart->get_Style());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExCharts, SetChartStyle)
{
    s_instance->SetChartStyle();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
