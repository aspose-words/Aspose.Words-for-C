#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisBound.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisBuiltInUnit.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisCategoryType.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisCrosses.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisDisplayUnit.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisScaleType.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisScaling.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisTickLabelPosition.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisTickMark.h>
#include <Aspose.Words.Cpp/Drawing/Charts/AxisTimeUnit.h>
#include <Aspose.Words.Cpp/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartAxis.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartAxisType.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartDataPoint.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartDataPointCollection.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartFormat.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartLegend.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartMarker.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartNumberFormat.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Drawing/Charts/LegendPosition.h>
#include <Aspose.Words.Cpp/Drawing/Charts/MarkerSymbol.h>
#include <Aspose.Words.Cpp/Drawing/Fill.h>
#include <Aspose.Words.Cpp/Drawing/PresetTexture.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Drawing/Stroke.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <drawing/color.h>
#include <system/array.h>
#include <system/collections/ienumerator.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/linq/enumerable.h>
#include <system/primitive_types.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace ApiExamples {

class ExCharts : public ApiExampleBase
{
public:
    void ChartTitle_()
    {
        //ExStart
        //ExFor:Chart
        //ExFor:Chart.Title
        //ExFor:ChartTitle
        //ExFor:ChartTitle.Overlay
        //ExFor:ChartTitle.Show
        //ExFor:ChartTitle.Text
        //ExSummary:Shows how to insert a chart and set a title.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a chart shape with a document builder and get its chart.
        SharedPtr<Shape> chartShape = builder->InsertChart(ChartType::Bar, 400, 300);
        SharedPtr<Chart> chart = chartShape->get_Chart();

        // Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
        SharedPtr<ChartTitle> title = chart->get_Title();
        title->set_Text(u"My Chart");

        // Set the "Show" property to "true" to make the title visible.
        title->set_Show(true);

        // Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
        title->set_Overlay(true);

        doc->Save(ArtifactsDir + u"Charts.ChartTitle.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartTitle.docx");
        chartShape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASSERT_EQ(ShapeType::NonPrimitive, chartShape->get_ShapeType());
        ASSERT_TRUE(chartShape->get_HasChart());

        title = chartShape->get_Chart()->get_Title();

        ASSERT_EQ(u"My Chart", title->get_Text());
        ASSERT_TRUE(title->get_Overlay());
        ASSERT_TRUE(title->get_Show());
    }

    void DataLabelNumberFormat()
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.NumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExSummary:Shows how to enable and configure data labels for a chart series.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add a line chart, then clear its demo data series to start with a clean chart,
        // and then set a title.
        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Clear();
        chart->get_Title()->set_Text(u"Monthly sales report");

        // Insert a custom chart series with months as categories for the X-axis,
        // and respective decimal amounts for the Y-axis.
        SharedPtr<ChartSeries> series =
            chart->get_Series()->Add(u"Revenue", MakeArray<String>({u"January", u"February", u"March"}), MakeArray<double>({25.611, 21.439, 33.750}));

        // Enable data labels, and then apply a custom number format for values displayed in the data labels.
        // This format will treat displayed decimal values as millions of US Dollars.
        series->set_HasDataLabels(true);
        SharedPtr<ChartDataLabelCollection> dataLabels = series->get_DataLabels();
        dataLabels->set_ShowValue(true);
        dataLabels->get_NumberFormat()->set_FormatCode(u"\"US$\" #,##0.000\"M\"");

        doc->Save(ArtifactsDir + u"Charts.DataLabelNumberFormat.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.DataLabelNumberFormat.docx");
        series = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

        ASSERT_TRUE(series->get_HasDataLabels());
        ASSERT_TRUE(series->get_DataLabels()->get_ShowValue());
        ASSERT_EQ(u"\"US$\" #,##0.000\"M\"", series->get_DataLabels()->get_NumberFormat()->get_FormatCode());
    }

    void DataArraysWrongSize()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        SharedPtr<ChartSeriesCollection> seriesColl = chart->get_Series();
        seriesColl->Clear();

        ArrayPtr<String> categories = MakeArray<String>({u"Cat1", nullptr, u"Cat3", u"Cat4", u"Cat5", nullptr});
        seriesColl->Add(u"AW Series 1", categories, MakeArray<double>({1, 2, std::numeric_limits<double>::quiet_NaN(), 4, 5, 6}));
        seriesColl->Add(u"AW Series 2", categories, MakeArray<double>({2, 3, std::numeric_limits<double>::quiet_NaN(), 5, 6, 7}));

        ASSERT_THROW(static_cast<std::function<SharedPtr<ChartSeries>()>>(
                         [&seriesColl, &categories]() -> SharedPtr<ChartSeries> {
                             return seriesColl->Add(u"AW Series 3", categories,
                                                    MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), 4, 5, std::numeric_limits<double>::quiet_NaN(),
                                                                       std::numeric_limits<double>::quiet_NaN()}));
                         })(),
                     System::ArgumentException);
        ASSERT_THROW(static_cast<std::function<SharedPtr<ChartSeries>()>>(
                         [&seriesColl, &categories]() -> SharedPtr<ChartSeries> {
                             return seriesColl->Add(u"AW Series 4", categories,
                                                    MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(),
                                                                       std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(),
                                                                       std::numeric_limits<double>::quiet_NaN()}));
                         })(),
                     System::ArgumentException);
    }

    void EmptyValuesInChartData()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        SharedPtr<ChartSeriesCollection> seriesColl = chart->get_Series();
        seriesColl->Clear();

        ArrayPtr<String> categories = MakeArray<String>({u"Cat1", nullptr, u"Cat3", u"Cat4", u"Cat5", nullptr});
        seriesColl->Add(u"AW Series 1", categories, MakeArray<double>({1, 2, std::numeric_limits<double>::quiet_NaN(), 4, 5, 6}));
        seriesColl->Add(u"AW Series 2", categories, MakeArray<double>({2, 3, std::numeric_limits<double>::quiet_NaN(), 5, 6, 7}));
        seriesColl->Add(u"AW Series 3", categories,
                        MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), 4, 5, std::numeric_limits<double>::quiet_NaN(), 7, 8}));
        seriesColl->Add(
            u"AW Series 4", categories,
            MakeArray<double>({std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(),
                               std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), 9}));

        doc->Save(ArtifactsDir + u"Charts.EmptyValuesInChartData.docx");
    }

    void AxisProperties()
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
        //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
        chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({640, 320, 280, 120, 150}));

        // Chart axes have various options that can change their appearance,
        // such as their direction, major/minor unit ticks, and tick marks.
        SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        xAxis->set_CategoryType(AxisCategoryType::Category);
        xAxis->set_Crosses(AxisCrosses::Minimum);
        xAxis->set_ReverseOrder(false);
        xAxis->set_MajorTickMark(AxisTickMark::Inside);
        xAxis->set_MinorTickMark(AxisTickMark::Cross);
        xAxis->set_MajorUnit(10.0);
        xAxis->set_MinorUnit(15.0);
        xAxis->set_TickLabelOffset(50);
        xAxis->set_TickLabelPosition(AxisTickLabelPosition::Low);
        xAxis->set_TickLabelSpacingIsAuto(false);
        xAxis->set_TickMarkSpacing(1);

        SharedPtr<ChartAxis> yAxis = chart->get_AxisY();
        yAxis->set_CategoryType(AxisCategoryType::Automatic);
        yAxis->set_Crosses(AxisCrosses::Maximum);
        yAxis->set_ReverseOrder(true);
        yAxis->set_MajorTickMark(AxisTickMark::Inside);
        yAxis->set_MinorTickMark(AxisTickMark::Cross);
        yAxis->set_MajorUnit(100.0);
        yAxis->set_MinorUnit(20.0);
        yAxis->set_TickLabelPosition(AxisTickLabelPosition::NextToAxis);

        // Column charts do not have a Z-axis.
        ASSERT_TRUE(chart->get_AxisZ() == nullptr);

        doc->Save(ArtifactsDir + u"Charts.AxisProperties.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisProperties.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASSERT_EQ(AxisCategoryType::Category, chart->get_AxisX()->get_CategoryType());
        ASSERT_EQ(AxisCrosses::Minimum, chart->get_AxisX()->get_Crosses());
        ASSERT_FALSE(chart->get_AxisX()->get_ReverseOrder());
        ASSERT_EQ(AxisTickMark::Inside, chart->get_AxisX()->get_MajorTickMark());
        ASSERT_EQ(AxisTickMark::Cross, chart->get_AxisX()->get_MinorTickMark());
        ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MajorUnit());
        ASPOSE_ASSERT_EQ(0.5, chart->get_AxisX()->get_MinorUnit());
        ASSERT_EQ(50, chart->get_AxisX()->get_TickLabelOffset());
        ASSERT_EQ(AxisTickLabelPosition::Low, chart->get_AxisX()->get_TickLabelPosition());
        ASSERT_FALSE(chart->get_AxisX()->get_TickLabelSpacingIsAuto());
        ASSERT_EQ(1, chart->get_AxisX()->get_TickMarkSpacing());

        ASSERT_EQ(AxisCategoryType::Category, chart->get_AxisY()->get_CategoryType());
        ASSERT_EQ(AxisCrosses::Maximum, chart->get_AxisY()->get_Crosses());
        ASSERT_TRUE(chart->get_AxisY()->get_ReverseOrder());
        ASSERT_EQ(AxisTickMark::Inside, chart->get_AxisY()->get_MajorTickMark());
        ASSERT_EQ(AxisTickMark::Cross, chart->get_AxisY()->get_MinorTickMark());
        ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
        ASPOSE_ASSERT_EQ(20.0, chart->get_AxisY()->get_MinorUnit());
        ASSERT_EQ(AxisTickLabelPosition::NextToAxis, chart->get_AxisY()->get_TickLabelPosition());
    }

    void DateTimeValues()
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

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
        chart->get_Series()->Add(u"Aspose Test Series",
                                 MakeArray<System::DateTime>({System::DateTime(2017, 11, 6), System::DateTime(2017, 11, 9), System::DateTime(2017, 11, 15),
                                                              System::DateTime(2017, 11, 21), System::DateTime(2017, 11, 25), System::DateTime(2017, 11, 29)}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2, 5.3}));

        // Set lower and upper bounds for the X-axis.
        SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        xAxis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(System::DateTime(2017, 11, 5).ToOADate()));
        xAxis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(System::DateTime(2017, 12, 3)));

        // Set the major units of the X-axis to a week, and the minor units to a day.
        xAxis->set_BaseTimeUnit(AxisTimeUnit::Days);
        xAxis->set_MajorUnit(7.0);
        xAxis->set_MajorTickMark(AxisTickMark::Cross);
        xAxis->set_MinorUnit(1.0);
        xAxis->set_MinorTickMark(AxisTickMark::Outside);

        // Define Y-axis properties for decimal values.
        SharedPtr<ChartAxis> yAxis = chart->get_AxisY();
        yAxis->set_TickLabelPosition(AxisTickLabelPosition::High);
        yAxis->set_MajorUnit(100.0);
        yAxis->set_MinorUnit(50.0);
        yAxis->get_DisplayUnit()->set_Unit(AxisBuiltInUnit::Hundreds);
        yAxis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(100.0));
        yAxis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(700.0));

        doc->Save(ArtifactsDir + u"Charts.DateTimeValues.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.DateTimeValues.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(System::DateTime(2017, 11, 5).ToOADate()), chart->get_AxisX()->get_Scaling()->get_Minimum());
        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(System::DateTime(2017, 12, 3)), chart->get_AxisX()->get_Scaling()->get_Maximum());
        ASSERT_EQ(AxisTimeUnit::Days, chart->get_AxisX()->get_BaseTimeUnit());
        ASPOSE_ASSERT_EQ(7.0, chart->get_AxisX()->get_MajorUnit());
        ASPOSE_ASSERT_EQ(1.0, chart->get_AxisX()->get_MinorUnit());
        ASSERT_EQ(AxisTickMark::Cross, chart->get_AxisX()->get_MajorTickMark());
        ASSERT_EQ(AxisTickMark::Outside, chart->get_AxisX()->get_MinorTickMark());

        ASSERT_EQ(AxisTickLabelPosition::High, chart->get_AxisY()->get_TickLabelPosition());
        ASPOSE_ASSERT_EQ(100.0, chart->get_AxisY()->get_MajorUnit());
        ASPOSE_ASSERT_EQ(50.0, chart->get_AxisY()->get_MinorUnit());
        ASSERT_EQ(AxisBuiltInUnit::Hundreds, chart->get_AxisY()->get_DisplayUnit()->get_Unit());
        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(100.0), chart->get_AxisY()->get_Scaling()->get_Minimum());
        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(700.0), chart->get_AxisY()->get_Scaling()->get_Maximum());
    }

    void HideChartAxis()
    {
        //ExStart
        //ExFor:ChartAxis.Hidden
        //ExSummary:Shows how to hide chart axes.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
        chart->get_Series()->Add(u"AW Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        // Hide the chart axes to simplify the appearance of the chart.
        chart->get_AxisX()->set_Hidden(true);
        chart->get_AxisY()->set_Hidden(true);

        doc->Save(ArtifactsDir + u"Charts.HideChartAxis.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.HideChartAxis.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASSERT_TRUE(chart->get_AxisX()->get_Hidden());
        ASSERT_TRUE(chart->get_AxisY()->get_Hidden());
    }

    void SetNumberFormatToChartAxis()
    {
        //ExStart
        //ExFor:ChartAxis.NumberFormat
        //ExFor:Charts.ChartNumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
        //ExSummary:Shows how to set formatting for chart values.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Add a custom series to the chart with categories for the X-axis,
        // and large respective numeric values for the Y-axis.
        chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

        // Set the number format of the Y-axis tick labels to not group digits with commas.
        chart->get_AxisY()->get_NumberFormat()->set_FormatCode(u"#,##0");

        // This flag can override the above value and draw the number format from the source cell.
        ASSERT_FALSE(chart->get_AxisY()->get_NumberFormat()->get_IsLinkedToSource());

        doc->Save(ArtifactsDir + u"Charts.SetNumberFormatToChartAxis.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.SetNumberFormatToChartAxis.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASSERT_EQ(u"#,##0", chart->get_AxisY()->get_NumberFormat()->get_FormatCode());
    }

    void TestDisplayChartsWithConversion(ChartType chartType)
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(chartType, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

        doc->Save(ArtifactsDir + u"Charts.TestDisplayChartsWithConversion.docx");
        doc->Save(ArtifactsDir + u"Charts.TestDisplayChartsWithConversion.pdf");
    }

    void Surface3DChart()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Surface3D, 500, 300);
        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Test Series 1", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

        chart->get_Series()->Add(u"Aspose Test Series 2", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({900000, 50000, 1100000, 400000, 2500000}));

        chart->get_Series()->Add(u"Aspose Test Series 3", MakeArray<String>({u"Word", u"PDF", u"Excel", u"GoogleDocs", u"Note"}),
                                 MakeArray<double>({500000, 820000, 1500000, 400000, 100000}));

        doc->Save(ArtifactsDir + u"Charts.SurfaceChart.docx");
        doc->Save(ArtifactsDir + u"Charts.SurfaceChart.pdf");
    }

    void DataLabelsBubbleChart()
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.Separator
        //ExFor:ChartDataLabelCollection.ShowBubbleSize
        //ExFor:ChartDataLabelCollection.ShowCategoryName
        //ExFor:ChartDataLabelCollection.ShowSeriesName
        //ExSummary:Shows how to work with data labels of a bubble chart.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Chart> chart = builder->InsertChart(ChartType::Bubble, 500, 300)->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Add a custom series with X/Y coordinates and diameter of each of the bubbles.
        SharedPtr<ChartSeries> series = chart->get_Series()->Add(u"Aspose Test Series", MakeArray<double>({2.9, 3.5, 1.1, 4.0, 4.0}),
                                                                 MakeArray<double>({1.9, 8.5, 2.1, 6.0, 1.5}), MakeArray<double>({9.0, 4.5, 2.5, 8.0, 5.0}));

        // Enable data labels, and then modify their appearance.
        series->set_HasDataLabels(true);
        SharedPtr<ChartDataLabelCollection> dataLabels = series->get_DataLabels();
        dataLabels->set_ShowBubbleSize(true);
        dataLabels->set_ShowCategoryName(true);
        dataLabels->set_ShowSeriesName(true);
        dataLabels->set_Separator(u" & ");

        doc->Save(ArtifactsDir + u"Charts.DataLabelsBubbleChart.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.DataLabelsBubbleChart.docx");
        dataLabels = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();

        ASSERT_TRUE(dataLabels->get_ShowBubbleSize());
        ASSERT_TRUE(dataLabels->get_ShowCategoryName());
        ASSERT_TRUE(dataLabels->get_ShowSeriesName());
        ASSERT_EQ(u" & ", dataLabels->get_Separator());
    }

    void DataLabelsPieChart()
    {
        //ExStart
        //ExFor:ChartDataLabelCollection.Separator
        //ExFor:ChartDataLabelCollection.ShowLeaderLines
        //ExFor:ChartDataLabelCollection.ShowLegendKey
        //ExFor:ChartDataLabelCollection.ShowPercentage
        //ExFor:ChartDataLabelCollection.ShowValue
        //ExSummary:Shows how to work with data labels of a pie chart.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Chart> chart = builder->InsertChart(ChartType::Pie, 500, 300)->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Insert a custom chart series with a category name for each of the sectors, and their frequency table.
        SharedPtr<ChartSeries> series =
            chart->get_Series()->Add(u"Aspose Test Series", MakeArray<String>({u"Word", u"PDF", u"Excel"}), MakeArray<double>({2.7, 3.2, 0.8}));

        // Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
        series->set_HasDataLabels(true);
        SharedPtr<ChartDataLabelCollection> dataLabels = series->get_DataLabels();
        dataLabels->set_ShowLeaderLines(true);
        dataLabels->set_ShowLegendKey(true);
        dataLabels->set_ShowPercentage(true);
        dataLabels->set_ShowValue(true);
        dataLabels->set_Separator(u"; ");

        doc->Save(ArtifactsDir + u"Charts.DataLabelsPieChart.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.DataLabelsPieChart.docx");
        dataLabels = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0)->get_DataLabels();

        ASSERT_TRUE(dataLabels->get_ShowLeaderLines());
        ASSERT_TRUE(dataLabels->get_ShowLegendKey());
        ASSERT_TRUE(dataLabels->get_ShowPercentage());
        ASSERT_TRUE(dataLabels->get_ShowValue());
        ASSERT_EQ(u"; ", dataLabels->get_Separator());
    }

    //ExStart
    //ExFor:ChartSeries
    //ExFor:ChartSeries.DataLabels
    //ExFor:ChartSeries.DataPoints
    //ExFor:ChartSeries.Name
    //ExFor:ChartDataLabel
    //ExFor:ChartDataLabel.Index
    //ExFor:ChartDataLabel.IsVisible
    //ExFor:ChartDataLabel.NumberFormat
    //ExFor:ChartDataLabel.Separator
    //ExFor:ChartDataLabel.ShowCategoryName
    //ExFor:ChartDataLabel.ShowDataLabelsRange
    //ExFor:ChartDataLabel.ShowLeaderLines
    //ExFor:ChartDataLabel.ShowLegendKey
    //ExFor:ChartDataLabel.ShowPercentage
    //ExFor:ChartDataLabel.ShowSeriesName
    //ExFor:ChartDataLabel.ShowValue
    //ExFor:ChartDataLabel.IsHidden
    //ExFor:ChartDataLabelCollection
    //ExFor:ChartDataLabelCollection.Add(System.Int32)
    //ExFor:ChartDataLabelCollection.Clear
    //ExFor:ChartDataLabelCollection.Count
    //ExFor:ChartDataLabelCollection.GetEnumerator
    //ExFor:ChartDataLabelCollection.Item(System.Int32)
    //ExFor:ChartDataLabelCollection.RemoveAt(System.Int32)
    //ExSummary:Shows how to apply labels to data points in a line chart.
    void DataLabels()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> chartShape = builder->InsertChart(ChartType::Line, 400, 300);
        SharedPtr<Chart> chart = chartShape->get_Chart();

        ASSERT_EQ(3, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
        ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());

        // Apply data labels to every series in the chart.
        // These labels will appear next to each data point in the graph and display its value.
        for (const auto& series : System::IterateOver(chart->get_Series()))
        {
            ApplyDataLabels(series, 4, u"000.0", u", ");
            ASSERT_EQ(4, series->get_DataLabels()->get_Count());
        }

        // Change the separator string for every data label in a series.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ChartDataLabel>>> enumerator =
                chart->get_Series()->idx_get(0)->get_DataLabels()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                ASSERT_EQ(u", ", enumerator->get_Current()->get_Separator());
                enumerator->get_Current()->set_Separator(u" & ");
            }
        }

        // For a cleaner looking graph, we can remove data labels individually.
        chart->get_Series()->idx_get(1)->get_DataLabels()->idx_get(2)->ClearFormat();

        // We can also strip an entire series of its data labels at once.
        chart->get_Series()->idx_get(2)->get_DataLabels()->ClearFormat();

        doc->Save(ArtifactsDir + u"Charts.DataLabels.docx");
    }

    /// <summary>
    /// Apply data labels with custom number format and separator to several data points in a series.
    /// </summary>
    static void ApplyDataLabels(SharedPtr<ChartSeries> series, int labelsCount, String numberFormat, String separator)
    {
        for (int i = 0; i < labelsCount; i++)
        {
            series->set_HasDataLabels(true);

            ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsVisible());

            series->get_DataLabels()->idx_get(i)->set_ShowCategoryName(true);
            series->get_DataLabels()->idx_get(i)->set_ShowSeriesName(true);
            series->get_DataLabels()->idx_get(i)->set_ShowValue(true);
            series->get_DataLabels()->idx_get(i)->set_ShowLeaderLines(true);
            series->get_DataLabels()->idx_get(i)->set_ShowLegendKey(true);
            series->get_DataLabels()->idx_get(i)->set_ShowPercentage(false);
            series->get_DataLabels()->idx_get(i)->set_IsHidden(false);
            ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_ShowDataLabelsRange());

            series->get_DataLabels()->idx_get(i)->get_NumberFormat()->set_FormatCode(numberFormat);
            series->get_DataLabels()->idx_get(i)->set_Separator(separator);

            ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_ShowDataLabelsRange());
            ASSERT_TRUE(series->get_DataLabels()->idx_get(i)->get_IsVisible());
            ASSERT_FALSE(series->get_DataLabels()->idx_get(i)->get_IsHidden());
        }
    }
    //ExEnd

    //ExStart
    //ExFor:ChartSeries.Smooth
    //ExFor:ChartDataPoint
    //ExFor:ChartDataPoint.Index
    //ExFor:ChartDataPointCollection
    //ExFor:ChartDataPointCollection.ClearFormat
    //ExFor:ChartDataPointCollection.Count
    //ExFor:ChartDataPointCollection.GetEnumerator
    //ExFor:ChartDataPointCollection.Item(System.Int32)
    //ExFor:ChartMarker
    //ExFor:ChartMarker.Size
    //ExFor:ChartMarker.Symbol
    //ExFor:IChartDataPoint
    //ExFor:IChartDataPoint.InvertIfNegative
    //ExFor:IChartDataPoint.Marker
    //ExFor:MarkerSymbol
    //ExSummary:Shows how to work with data points on a line chart.
    void ChartDataPoint_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 500, 350);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(3, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
        ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());

        // Emphasize the chart's data points by making them appear as diamond shapes.
        for (const auto& series : System::IterateOver(chart->get_Series()))
        {
            ApplyDataPoints(series, 4, MarkerSymbol::Diamond, 15);
        }

        // Smooth out the line that represents the first data series.
        chart->get_Series()->idx_get(0)->set_Smooth(true);

        // Verify that data points for the first series will not invert their colors if the value is negative.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ChartDataPoint>>> enumerator =
                chart->get_Series()->idx_get(0)->get_DataPoints()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                ASSERT_FALSE(enumerator->get_Current()->get_InvertIfNegative());
            }
        }

        // For a cleaner looking graph, we can clear format individually.
        chart->get_Series()->idx_get(1)->get_DataPoints()->idx_get(2)->ClearFormat();

        // We can also strip an entire series of data points at once.
        chart->get_Series()->idx_get(2)->get_DataPoints()->ClearFormat();

        doc->Save(ArtifactsDir + u"Charts.ChartDataPoint.docx");
    }

    /// <summary>
    /// Applies a number of data points to a series.
    /// </summary>
    static void ApplyDataPoints(SharedPtr<ChartSeries> series, int dataPointsCount, MarkerSymbol markerSymbol, int dataPointSize)
    {
        for (int i = 0; i < dataPointsCount; i++)
        {
            SharedPtr<ChartDataPoint> point = series->get_DataPoints()->idx_get(i);
            point->get_Marker()->set_Symbol(markerSymbol);
            point->get_Marker()->set_Size(dataPointSize);

            ASSERT_EQ(i, point->get_Index());
        }
    }
    //ExEnd

    void PieChartExplosion()
    {
        //ExStart
        //ExFor:Charts.IChartDataPoint.Explosion
        //ExSummary:Shows how to move the slices of a pie chart away from the center.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Pie, 500, 350);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(1, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Sales", chart->get_Series()->idx_get(0)->get_Name());

        // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
        // Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
        // Aspose.Words create data points automatically if them does not exist.
        SharedPtr<ChartDataPoint> dataPoint = chart->get_Series()->idx_get(0)->get_DataPoints()->idx_get(0);
        dataPoint->set_Explosion(10);

        // Displace the second portion by a greater distance.
        dataPoint = chart->get_Series()->idx_get(0)->get_DataPoints()->idx_get(1);
        dataPoint->set_Explosion(40);

        doc->Save(ArtifactsDir + u"Charts.PieChartExplosion.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.PieChartExplosion.docx");
        SharedPtr<ChartSeries> series = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

        ASSERT_EQ(10, series->get_DataPoints()->idx_get(0)->get_Explosion());
        ASSERT_EQ(40, series->get_DataPoints()->idx_get(1)->get_Explosion());
    }

    void Bubble3D()
    {
        //ExStart
        //ExFor:Charts.ChartDataLabel.ShowBubbleSize
        //ExFor:Charts.IChartDataPoint.Bubble3D
        //ExSummary:Shows how to use 3D effects with bubble charts.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Bubble3D, 500, 350);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(1, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Y-Values", chart->get_Series()->idx_get(0)->get_Name());
        ASSERT_TRUE(chart->get_Series()->idx_get(0)->get_Bubble3D());

        // Apply a data label to each bubble that displays its diameter.
        for (int i = 0; i < 3; i++)
        {
            chart->get_Series()->idx_get(0)->set_HasDataLabels(true);
            chart->get_Series()->idx_get(0)->get_DataLabels()->idx_get(i)->set_ShowBubbleSize(true);
        }

        doc->Save(ArtifactsDir + u"Charts.Bubble3D.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.Bubble3D.docx");
        SharedPtr<ChartSeries> series = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Series()->idx_get(0);

        for (int i = 0; i < 3; i++)
        {
            ASSERT_TRUE(series->get_DataLabels()->idx_get(i)->get_ShowBubbleSize());
        }
    }

    //ExStart
    //ExFor:ChartAxis.Type
    //ExFor:ChartAxisType
    //ExFor:ChartType
    //ExFor:Chart.Series
    //ExFor:ChartSeriesCollection.Add(String,DateTime[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,String[],Double[])
    //ExSummary:Shows how to create an appropriate type of chart series for a graph type.
    void ChartSeriesCollection_()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // There are several ways of populating a chart's series collection.
        // Different series schemas are intended for different chart types.
        // 1 -  Column chart with columns grouped and banded along the X-axis by category:
        SharedPtr<Chart> chart = AppendChart(builder, ChartType::Column, 500, 300);

        ArrayPtr<String> categories = MakeArray<String>({u"Category 1", u"Category 2", u"Category 3"});

        // Insert two series of decimal values containing a value for each respective category.
        // This column chart will have three groups, each with two columns.
        chart->get_Series()->Add(u"Series 1", categories, MakeArray<double>({76.6, 82.1, 91.6}));
        chart->get_Series()->Add(u"Series 2", categories, MakeArray<double>({64.2, 79.5, 94.0}));

        // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
        ASSERT_EQ(ChartAxisType::Category, chart->get_AxisX()->get_Type());
        ASSERT_EQ(ChartAxisType::Value, chart->get_AxisY()->get_Type());

        // 2 -  Area chart with dates distributed along the X-axis:
        chart = AppendChart(builder, ChartType::Area, 500, 300);

        ArrayPtr<System::DateTime> dates =
            MakeArray<System::DateTime>({System::DateTime(2014, 3, 31), System::DateTime(2017, 1, 23), System::DateTime(2017, 6, 18),
                                         System::DateTime(2019, 11, 22), System::DateTime(2020, 9, 7)});

        // Insert a series with a decimal value for each respective date.
        // The dates will be distributed along a linear X-axis,
        // and the values added to this series will create data points.
        chart->get_Series()->Add(u"Series 1", dates, MakeArray<double>({15.8, 21.5, 22.9, 28.7, 33.1}));

        ASSERT_EQ(ChartAxisType::Category, chart->get_AxisX()->get_Type());
        ASSERT_EQ(ChartAxisType::Value, chart->get_AxisY()->get_Type());

        // 3 -  2D scatter plot:
        chart = AppendChart(builder, ChartType::Scatter, 500, 300);

        // Each series will need two decimal arrays of equal length.
        // The first array contains X-values, and the second contains corresponding Y-values
        // of data points on the chart's graph.
        chart->get_Series()->Add(u"Series 1", MakeArray<double>({3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6}),
                                 MakeArray<double>({3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9}));
        chart->get_Series()->Add(u"Series 2", MakeArray<double>({2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3}),
                                 MakeArray<double>({7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6}));

        ASSERT_EQ(ChartAxisType::Value, chart->get_AxisX()->get_Type());
        ASSERT_EQ(ChartAxisType::Value, chart->get_AxisY()->get_Type());

        // 4 -  Bubble chart:
        chart = AppendChart(builder, ChartType::Bubble, 500, 300);

        // Each series will need three decimal arrays of equal length.
        // The first array contains X-values, the second contains corresponding Y-values,
        // and the third contains diameters for each of the graph's data points.
        chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.1, 5.0, 9.8}), MakeArray<double>({1.2, 4.9, 9.9}), MakeArray<double>({2.0, 4.0, 8.0}));

        doc->Save(ArtifactsDir + u"Charts.ChartSeriesCollection.docx");
    }

    /// <summary>
    /// Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
    /// </summary>
    static SharedPtr<Chart> AppendChart(SharedPtr<DocumentBuilder> builder, ChartType chartType, double width, double height)
    {
        SharedPtr<Shape> chartShape = builder->InsertChart(chartType, width, height);
        SharedPtr<Chart> chart = chartShape->get_Chart();
        chart->get_Series()->Clear();
        EXPECT_EQ(0, chart->get_Series()->get_Count());
        //ExSkip

        return chart;
    }
    //ExEnd

    void ChartSeriesCollectionModify()
    {
        //ExStart
        //ExFor:ChartSeriesCollection
        //ExFor:ChartSeriesCollection.Clear
        //ExFor:ChartSeriesCollection.Count
        //ExFor:ChartSeriesCollection.GetEnumerator
        //ExFor:ChartSeriesCollection.Item(Int32)
        //ExFor:ChartSeriesCollection.RemoveAt(Int32)
        //ExSummary:Shows how to add and remove series data in a chart.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert a column chart that will contain three series of demo data by default.
        SharedPtr<Shape> chartShape = builder->InsertChart(ChartType::Column, 400, 300);
        SharedPtr<Chart> chart = chartShape->get_Chart();

        // Each series has four decimal values: one for each of the four categories.
        // Four clusters of three columns will represent this data.
        SharedPtr<ChartSeriesCollection> chartData = chart->get_Series();

        ASSERT_EQ(3, chartData->get_Count());

        // Print the name of every series in the chart.
        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<ChartSeries>>> enumerator = chart->get_Series()->GetEnumerator();
            while (enumerator->MoveNext())
            {
                std::cout << enumerator->get_Current()->get_Name() << std::endl;
            }
        }

        // These are the names of the categories in the chart.
        ArrayPtr<String> categories = MakeArray<String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"});

        // We can add a series with new values for existing categories.
        // This chart will now contain four clusters of four columns.
        chart->get_Series()->Add(u"Series 4", categories, MakeArray<double>({4.4, 7.0, 3.5, 2.1}));
        ASSERT_EQ(4, chartData->get_Count());
        //ExSkip
        ASSERT_EQ(u"Series 4", chartData->idx_get(3)->get_Name());
        //ExSkip

        // A chart series can also be removed by index, like this.
        // This will remove one of the three demo series that came with the chart.
        chartData->RemoveAt(2);

        ASSERT_FALSE(chartData->LINQ_Any([](SharedPtr<ChartSeries> s) { return s->get_Name() == u"Series 3"; }));
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

    void AxisScaling_()
    {
        //ExStart
        //ExFor:AxisScaleType
        //ExFor:AxisScaling
        //ExFor:AxisScaling.LogBase
        //ExFor:AxisScaling.Type
        //ExSummary:Shows how to apply logarithmic scaling to a chart axis.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> chartShape = builder->InsertChart(ChartType::Scatter, 450, 300);
        SharedPtr<Chart> chart = chartShape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Insert a series with X/Y coordinates for five points.
        chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.0, 2.0, 3.0, 4.0, 5.0}), MakeArray<double>({1.0, 20.0, 400.0, 8000.0, 160000.0}));

        // The scaling of the X-axis is linear by default,
        // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
        // A linear axis is not ideal for our Y-values
        // since the points with the smaller Y-values will be harder to read.
        // A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
        // will spread the plotted points, allowing us to read their values on the chart more easily.
        chart->get_AxisY()->get_Scaling()->set_Type(AxisScaleType::Logarithmic);
        chart->get_AxisY()->get_Scaling()->set_LogBase(20);

        doc->Save(ArtifactsDir + u"Charts.AxisScaling.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisScaling.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASSERT_EQ(AxisScaleType::Linear, chart->get_AxisX()->get_Scaling()->get_Type());
        ASSERT_EQ(AxisScaleType::Logarithmic, chart->get_AxisY()->get_Scaling()->get_Type());
        ASPOSE_ASSERT_EQ(20.0, chart->get_AxisY()->get_Scaling()->get_LogBase());
    }

    void AxisBound_()
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

        SharedPtr<Shape> chartShape = builder->InsertChart(ChartType::Scatter, 450, 300);
        SharedPtr<Chart> chart = chartShape->get_Chart();

        // Clear the chart's demo data series to start with a clean chart.
        chart->get_Series()->Clear();

        // Add a series with two decimal arrays. The first array contains the X-values,
        // and the second contains corresponding Y-values for points in the scatter chart.
        chart->get_Series()->Add(u"Series 1", MakeArray<double>({1.1, 5.4, 7.9, 3.5, 2.1, 9.7}), MakeArray<double>({2.1, 0.3, 0.6, 3.3, 1.4, 1.9}));

        // By default, default scaling is applied to the graph's X and Y-axes,
        // so that both their ranges are big enough to encompass every X and Y-value of every series.
        ASSERT_TRUE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());

        // We can define our own axis bounds.
        // In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
        chart->get_AxisX()->get_Scaling()->set_Minimum(MakeObject<AxisBound>(0.0));
        chart->get_AxisX()->get_Scaling()->set_Maximum(MakeObject<AxisBound>(10.0));
        chart->get_AxisY()->get_Scaling()->set_Minimum(MakeObject<AxisBound>(0.0));
        chart->get_AxisY()->get_Scaling()->set_Maximum(MakeObject<AxisBound>(10.0));

        ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
        ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());

        // Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
        chartShape = builder->InsertChart(ChartType::Line, 450, 300);
        chart = chartShape->get_Chart();
        chart->get_Series()->Clear();

        ArrayPtr<System::DateTime> dates =
            MakeArray<System::DateTime>({System::DateTime(1973, 5, 11), System::DateTime(1981, 2, 4), System::DateTime(1985, 9, 23),
                                         System::DateTime(1989, 6, 28), System::DateTime(1994, 12, 15)});

        chart->get_Series()->Add(u"Series 1", dates, MakeArray<double>({3.0, 4.7, 5.9, 7.1, 8.9}));

        // We can set axis bounds in the form of dates as well, limiting the chart to a period.
        // Setting the range to 1980-1990 will omit the two of the series values
        // that are outside of the range from the graph.
        chart->get_AxisX()->get_Scaling()->set_Minimum(MakeObject<AxisBound>(System::DateTime(1980, 1, 1)));
        chart->get_AxisX()->get_Scaling()->set_Maximum(MakeObject<AxisBound>(System::DateTime(1990, 1, 1)));

        doc->Save(ArtifactsDir + u"Charts.AxisBound.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisBound.docx");
        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart();

        ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
        ASPOSE_ASSERT_EQ(0.0, chart->get_AxisX()->get_Scaling()->get_Minimum()->get_Value());
        ASPOSE_ASSERT_EQ(10.0, chart->get_AxisX()->get_Scaling()->get_Maximum()->get_Value());

        ASSERT_FALSE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
        ASPOSE_ASSERT_EQ(0.0, chart->get_AxisY()->get_Scaling()->get_Minimum()->get_Value());
        ASPOSE_ASSERT_EQ(10.0, chart->get_AxisY()->get_Scaling()->get_Maximum()->get_Value());

        chart = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 1, true)))->get_Chart();

        ASSERT_FALSE(chart->get_AxisX()->get_Scaling()->get_Minimum()->get_IsAuto());
        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(System::DateTime(1980, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Minimum());
        ASPOSE_ASSERT_EQ(MakeObject<AxisBound>(System::DateTime(1990, 1, 1)), chart->get_AxisX()->get_Scaling()->get_Maximum());

        ASSERT_TRUE(chart->get_AxisY()->get_Scaling()->get_Minimum()->get_IsAuto());
    }

    void ChartLegend_()
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

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 450, 300);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(3, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
        ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());

        // Move the chart's legend to the top right corner.
        SharedPtr<ChartLegend> legend = chart->get_Legend();
        legend->set_Position(LegendPosition::TopRight);

        // Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
        legend->set_Overlay(true);

        doc->Save(ArtifactsDir + u"Charts.ChartLegend.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.ChartLegend.docx");

        legend = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_Legend();

        ASSERT_TRUE(legend->get_Overlay());
        ASSERT_EQ(LegendPosition::TopRight, legend->get_Position());
    }

    void AxisCross()
    {
        //ExStart
        //ExFor:ChartAxis.AxisBetweenCategories
        //ExFor:ChartAxis.CrossesAt
        //ExSummary:Shows how to get a graph axis to cross at a custom location.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 450, 250);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(3, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Series 1", chart->get_Series()->idx_get(0)->get_Name());
        ASSERT_EQ(u"Series 2", chart->get_Series()->idx_get(1)->get_Name());
        ASSERT_EQ(u"Series 3", chart->get_Series()->idx_get(2)->get_Name());

        // For column charts, the Y-axis crosses at zero by default,
        // which means that columns for all values below zero point down to represent negative values.
        // We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
        SharedPtr<ChartAxis> axis = chart->get_AxisX();
        axis->set_Crosses(AxisCrosses::Custom);
        axis->set_CrossesAt(3);
        axis->set_AxisBetweenCategories(true);

        doc->Save(ArtifactsDir + u"Charts.AxisCross.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisCross.docx");
        axis = (System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_Chart()->get_AxisX();

        ASSERT_TRUE(axis->get_AxisBetweenCategories());
        ASSERT_EQ(AxisCrosses::Custom, axis->get_Crosses());
        ASPOSE_ASSERT_EQ(3.0, axis->get_CrossesAt());
    }

    void AxisDisplayUnit_()
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

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Scatter, 450, 250);
        SharedPtr<Chart> chart = shape->get_Chart();

        ASSERT_EQ(1, chart->get_Series()->get_Count());
        ASSERT_EQ(u"Y-Values", chart->get_Series()->idx_get(0)->get_Name());

        // Set the minor tick marks of the Y-axis to point away from the plot area,
        // and the major tick marks to cross the axis.
        SharedPtr<ChartAxis> axis = chart->get_AxisY();
        axis->set_MajorTickMark(AxisTickMark::Cross);
        axis->set_MinorTickMark(AxisTickMark::Outside);

        // Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
        axis->set_MajorUnit(10);
        axis->set_MinorUnit(1);

        // Set the Y-axis bounds to -10 and 20.
        // This Y-axis will now display 4 major tick marks and 27 minor tick marks.
        axis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(-10.0));
        axis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(20.0));

        // For the X-axis, set the major tick marks at every 10 units,
        // every minor tick mark at 2.5 units.
        axis = chart->get_AxisX();
        axis->set_MajorUnit(10);
        axis->set_MinorUnit(2.5);

        // Configure both types of tick marks to appear inside the graph plot area.
        axis->set_MajorTickMark(AxisTickMark::Inside);
        axis->set_MinorTickMark(AxisTickMark::Inside);

        // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
        axis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(-10.0));
        axis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(30.0));
        axis->set_TickLabelAlignment(ParagraphAlignment::Right);

        ASSERT_EQ(1, axis->get_TickLabelSpacing());

        // Set the tick labels to display their value in millions.
        axis->get_DisplayUnit()->set_Unit(AxisBuiltInUnit::Millions);

        // We can set a more specific value by which tick labels will display their values.
        // This statement is equivalent to the one above.
        axis->get_DisplayUnit()->set_CustomUnit(1000000);
        ASSERT_EQ(AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
        //ExSkip

        doc->Save(ArtifactsDir + u"Charts.AxisDisplayUnit.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Charts.AxisDisplayUnit.docx");
        shape = System::DynamicCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        ASPOSE_ASSERT_EQ(450.0, shape->get_Width());
        ASPOSE_ASSERT_EQ(250.0, shape->get_Height());

        axis = shape->get_Chart()->get_AxisX();

        ASSERT_EQ(AxisTickMark::Inside, axis->get_MajorTickMark());
        ASSERT_EQ(AxisTickMark::Inside, axis->get_MinorTickMark());
        ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
        ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
        ASPOSE_ASSERT_EQ(30.0, axis->get_Scaling()->get_Maximum()->get_Value());
        ASSERT_EQ(1, axis->get_TickLabelSpacing());
        ASSERT_EQ(ParagraphAlignment::Right, axis->get_TickLabelAlignment());
        ASSERT_EQ(AxisBuiltInUnit::Custom, axis->get_DisplayUnit()->get_Unit());
        ASPOSE_ASSERT_EQ(1000000.0, axis->get_DisplayUnit()->get_CustomUnit());

        axis = shape->get_Chart()->get_AxisY();

        ASSERT_EQ(AxisTickMark::Cross, axis->get_MajorTickMark());
        ASSERT_EQ(AxisTickMark::Outside, axis->get_MinorTickMark());
        ASPOSE_ASSERT_EQ(10.0, axis->get_MajorUnit());
        ASPOSE_ASSERT_EQ(1.0, axis->get_MinorUnit());
        ASPOSE_ASSERT_EQ(-10.0, axis->get_Scaling()->get_Minimum()->get_Value());
        ASPOSE_ASSERT_EQ(20.0, axis->get_Scaling()->get_Maximum()->get_Value());
    }

    void MarkerFormatting()
    {
        //ExStart
        //ExFor:ChartMarker.Format
        //ExFor:ChartFormat.Fill
        //ExFor:ChartFormat.Stroke
        //ExFor:Stroke.ForeColor
        //ExFor:Stroke.BackColor
        //ExFor:Stroke.Visible
        //ExFor:Stroke.Transparency
        //ExFor:Fill.PresetTextured(PresetTexture)
        //ExSummary:Show how to set marker formatting.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Scatter, 432, 252);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Delete default generated series.
        chart->get_Series()->Clear();
        SharedPtr<ChartSeries> series =
            chart->get_Series()->Add(u"AW Series 1", MakeArray<double>({0.7, 1.8, 2.6, 3.9}), MakeArray<double>({2.7, 3.2, 0.8, 1.7}));

        // Set marker formatting.
        series->get_Marker()->set_Size(40);
        series->get_Marker()->set_Symbol(MarkerSymbol::Square);
        SharedPtr<ChartDataPointCollection> dataPoints = series->get_DataPoints();
        dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Fill()->PresetTextured(PresetTexture::Denim);
        dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
        dataPoints->idx_get(0)->get_Marker()->get_Format()->get_Stroke()->set_BackColor(System::Drawing::Color::get_Red());
        dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Fill()->PresetTextured(PresetTexture::WaterDroplets);
        dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
        dataPoints->idx_get(1)->get_Marker()->get_Format()->get_Stroke()->set_Visible(false);
        dataPoints->idx_get(2)->get_Marker()->get_Format()->get_Fill()->PresetTextured(PresetTexture::GreenMarble);
        dataPoints->idx_get(2)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
        dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Fill()->PresetTextured(PresetTexture::Oak);
        dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Stroke()->set_ForeColor(System::Drawing::Color::get_Yellow());
        dataPoints->idx_get(3)->get_Marker()->get_Format()->get_Stroke()->set_Transparency(0.5);

        doc->Save(ArtifactsDir + u"Charts.MarkerFormatting.docx");
        //ExEnd
    }

    void SeriesColor()
    {
        //ExStart
        //ExFor:ChartSeries.Format
        //ExSummary:Sows how to set series color.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        SharedPtr<ChartSeriesCollection> seriesColl = chart->get_Series();

        // Delete default generated series.
        seriesColl->Clear();

        // Create category names array.
        auto categories = MakeArray<String>({u"Category 1", u"Category 2"});

        // Adding new series. Value and category arrays must be the same size.
        SharedPtr<ChartSeries> series1 = seriesColl->Add(u"Series 1", categories, MakeArray<double>({1, 2}));
        SharedPtr<ChartSeries> series2 = seriesColl->Add(u"Series 2", categories, MakeArray<double>({3, 4}));
        SharedPtr<ChartSeries> series3 = seriesColl->Add(u"Series 3", categories, MakeArray<double>({5, 6}));

        // Set series color.
        series1->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
        series2->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Yellow());
        series3->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Blue());

        doc->Save(ArtifactsDir + u"Charts.SeriesColor.docx");
        //ExEnd
    }

    void DataPointsFormatting()
    {
        //ExStart
        //ExFor:ChartDataPoint.Format
        //ExSummary:Shows how to set individual formatting for categories of a column chart.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        SharedPtr<Chart> chart = shape->get_Chart();

        // Delete default generated series.
        chart->get_Series()->Clear();

        // Adding new series.
        SharedPtr<ChartSeries> series = chart->get_Series()->Add(u"Series 1", MakeArray<String>({u"Category 1", u"Category 2", u"Category 3", u"Category 4"}),
                                                                 MakeArray<double>({1, 2, 3, 4}));

        // Set column formatting.
        SharedPtr<ChartDataPointCollection> dataPoints = series->get_DataPoints();
        dataPoints->idx_get(0)->get_Format()->get_Fill()->PresetTextured(PresetTexture::Denim);
        dataPoints->idx_get(1)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Red());
        dataPoints->idx_get(2)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Yellow());
        dataPoints->idx_get(3)->get_Format()->get_Fill()->set_ForeColor(System::Drawing::Color::get_Blue());

        doc->Save(ArtifactsDir + u"Charts.DataPointsFormatting.docx");
        //ExEnd
    }
};

} // namespace ApiExamples
