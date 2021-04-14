#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBound.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisBuiltInUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCategoryType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisCrosses.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisDisplayUnit.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisScaling.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickLabelPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/AxisTickMark.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartAxis.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabel.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataLabelCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPoint.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartDataPointCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartLegend.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartMarker.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartNumberFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeriesCollection.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartTitle.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/LegendPosition.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/MarkerSymbol.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <system/array.h>
#include <system/date_time.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements {

class WorkingWithCharts : public DocsExamplesBase
{
public:
    void FormatNumberOfDataLabel()
    {
        //ExStart:FormatNumberOfDataLabel
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Title()->set_Text(u"Data Labels With Different Number Format");

        // Delete default generated series.
        chart->get_Series()->Clear();

        SharedPtr<ChartSeries> series1 =
            chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Category 1", u"Category 2", u"Category 3"}), MakeArray<double>({2.5, 1.5, 3.5}));

        series1->set_HasDataLabels(true);
        series1->get_DataLabels()->set_ShowValue(true);
        series1->get_DataLabels()->idx_get(0)->get_NumberFormat()->set_FormatCode(u"\"$\"#,##0.00");
        series1->get_DataLabels()->idx_get(1)->get_NumberFormat()->set_FormatCode(u"dd/mm/yyyy");
        series1->get_DataLabels()->idx_get(2)->get_NumberFormat()->set_FormatCode(u"0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        series1->get_DataLabels()->idx_get(2)->get_NumberFormat()->set_IsLinkedToSource(true);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.FormatNumberOfDataLabel.docx");
        //ExEnd:FormatNumberOfDataLabel
    }

    void CreateChartUsingShape()
    {
        //ExStart:CreateChartUsingShape
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Title()->set_Show(true);
        chart->get_Title()->set_Text(u"Line Chart Title");
        chart->get_Title()->set_Overlay(false);

        // Please note if null or empty value is specified as title text, auto generated title will be shown.

        chart->get_Legend()->set_Position(LegendPosition::Left);
        chart->get_Legend()->set_Overlay(true);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.CreateChartUsingShape.docx");
        //ExEnd:CreateChartUsingShape
    }

    void InsertSimpleColumnChart()
    {
        //ExStart:InsertSimpleColumnChart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // You can specify different chart types and sizes.
        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        //ExStart:ChartSeriesCollection
        SharedPtr<ChartSeriesCollection> seriesColl = chart->get_Series();

        std::cout << seriesColl->get_Count() << std::endl;
        //ExEnd:ChartSeriesCollection

        // Delete default generated series.
        seriesColl->Clear();

        // Create category names array, in this example we have two categories.
        auto categories = MakeArray<String>({u"Category 1", u"Category 2"});

        // Please note, data arrays must not be empty and arrays must be the same size.
        seriesColl->Add(u"Aspose Series 1", categories, MakeArray<double>({1, 2}));
        seriesColl->Add(u"Aspose Series 2", categories, MakeArray<double>({3, 4}));
        seriesColl->Add(u"Aspose Series 3", categories, MakeArray<double>({5, 6}));
        seriesColl->Add(u"Aspose Series 4", categories, MakeArray<double>({7, 8}));
        seriesColl->Add(u"Aspose Series 5", categories, MakeArray<double>({9, 10}));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.InsertSimpleColumnChart.docx");
        //ExEnd:InsertSimpleColumnChart
    }

    void InsertColumnChart()
    {
        //ExStart:InsertColumnChart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Category 1", u"Category 2"}), MakeArray<double>({1, 2}));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.InsertColumnChart.docx");
        //ExEnd:InsertColumnChart
    }

    void InsertAreaChart()
    {
        //ExStart:InsertAreaChart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Area, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Add(u"Aspose Series 1",
                                 MakeArray<System::DateTime>({System::DateTime(2002, 5, 1), System::DateTime(2002, 6, 1), System::DateTime(2002, 7, 1),
                                                              System::DateTime(2002, 8, 1), System::DateTime(2002, 9, 1)}),
                                 MakeArray<double>({32, 32, 28, 12, 15}));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.InsertAreaChart.docx");
        //ExEnd:InsertAreaChart
    }

    void InsertBubbleChart()
    {
        //ExStart:InsertBubbleChart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Bubble, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<double>({0.7, 1.8, 2.6}), MakeArray<double>({2.7, 3.2, 0.8}), MakeArray<double>({10, 4, 8}));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.InsertBubbleChart.docx");
        //ExEnd:InsertBubbleChart
    }

    void InsertScatterChart()
    {
        //ExStart:InsertScatterChart
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Scatter, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<double>({0.7, 1.8, 2.6}), MakeArray<double>({2.7, 3.2, 0.8}));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.InsertScatterChart.docx");
        //ExEnd:InsertScatterChart
    }

    void DefineXYAxisProperties()
    {
        //ExStart:DefineXYAxisProperties
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert chart
        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Area, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1",
                                 MakeArray<System::DateTime>({System::DateTime(2002, 1, 1), System::DateTime(2002, 6, 1), System::DateTime(2002, 7, 1),
                                                              System::DateTime(2002, 8, 1), System::DateTime(2002, 9, 1)}),
                                 MakeArray<double>({640, 320, 280, 120, 150}));

        SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        SharedPtr<ChartAxis> yAxis = chart->get_AxisY();

        // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
        xAxis->set_CategoryType(AxisCategoryType::Category);
        xAxis->set_Crosses(AxisCrosses::Custom);
        xAxis->set_CrossesAt(3);
        // Measured in display units of the Y axis (hundreds).
        xAxis->set_ReverseOrder(true);
        xAxis->set_MajorTickMark(AxisTickMark::Cross);
        xAxis->set_MinorTickMark(AxisTickMark::Outside);
        xAxis->set_TickLabelOffset(200);

        yAxis->set_TickLabelPosition(AxisTickLabelPosition::High);
        yAxis->set_MajorUnit(100);
        yAxis->set_MinorUnit(50);
        yAxis->get_DisplayUnit()->set_Unit(AxisBuiltInUnit::Hundreds);
        yAxis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(100.0));
        yAxis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(700.0));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.DefineXYAxisProperties.docx");
        //ExEnd:DefineXYAxisProperties
    }

    void DateTimeValuesToAxis()
    {
        //ExStart:SetDateTimeValuesToAxis
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);
        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1",
                                 MakeArray<System::DateTime>({System::DateTime(2017, 11, 6), System::DateTime(2017, 11, 9), System::DateTime(2017, 11, 15),
                                                              System::DateTime(2017, 11, 21), System::DateTime(2017, 11, 25), System::DateTime(2017, 11, 29)}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2, 5.3}));

        SharedPtr<ChartAxis> xAxis = chart->get_AxisX();
        xAxis->get_Scaling()->set_Minimum(MakeObject<AxisBound>(System::DateTime(2017, 11, 5).ToOADate()));
        xAxis->get_Scaling()->set_Maximum(MakeObject<AxisBound>(System::DateTime(2017, 12, 3).ToOADate()));

        // Set major units to a week and minor units to a day.
        xAxis->set_MajorUnit(7);
        xAxis->set_MinorUnit(1);
        xAxis->set_MajorTickMark(AxisTickMark::Cross);
        xAxis->set_MinorTickMark(AxisTickMark::Outside);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.DateTimeValuesToAxis.docx");
        //ExEnd:SetDateTimeValuesToAxis
    }

    void NumberFormatForAxis()
    {
        //ExStart:SetNumberFormatForAxis
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 MakeArray<double>({1900000, 850000, 2100000, 600000, 1500000}));

        chart->get_AxisY()->get_NumberFormat()->set_FormatCode(u"#,##0");

        doc->Save(ArtifactsDir + u"WorkingWithCharts.NumberFormatForAxis.docx");
        //ExEnd:SetNumberFormatForAxis
    }

    void BoundsOfAxis()
    {
        //ExStart:SetboundsOfAxis
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        chart->get_AxisY()->get_Scaling()->set_Minimum(MakeObject<AxisBound>(0.0));
        chart->get_AxisY()->get_Scaling()->set_Maximum(MakeObject<AxisBound>(6.0));

        doc->Save(ArtifactsDir + u"WorkingWithCharts.BoundsOfAxis.docx");
        //ExEnd:SetboundsOfAxis
    }

    void IntervalUnitBetweenLabelsOnAxis()
    {
        //ExStart:SetIntervalUnitBetweenLabelsOnAxis
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        chart->get_AxisX()->set_TickLabelSpacing(2);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.IntervalUnitBetweenLabelsOnAxis.docx");
        //ExEnd:SetIntervalUnitBetweenLabelsOnAxis
    }

    void HideChartAxis()
    {
        //ExStart:HideChartAxis
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Column, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Item 1", u"Item 2", u"Item 3", u"Item 4", u"Item 5"}),
                                 MakeArray<double>({1.2, 0.3, 2.1, 2.9, 4.2}));

        chart->get_AxisY()->set_Hidden(true);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.HideChartAxis.docx");
        //ExEnd:HideChartAxis
    }

    void TickMultiLineLabelAlignment()
    {
        //ExStart:TickMultiLineLabelAlignment
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Scatter, 450, 250);

        SharedPtr<ChartAxis> axis = shape->get_Chart()->get_AxisX();
        // This property has effect only for multi-line labels.
        axis->set_TickLabelAlignment(ParagraphAlignment::Right);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.TickMultiLineLabelAlignment.docx");
        //ExEnd:TickMultiLineLabelAlignment
    }

    void ChartDataLabel()
    {
        //ExStart:WorkWithChartDataLabel
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Bar, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        SharedPtr<ChartSeries> series0 = shape->get_Chart()->get_Series()->idx_get(0);

        SharedPtr<ChartDataLabelCollection> labels = series0->get_DataLabels();
        labels->set_ShowLegendKey(true);
        // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
        // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its
        // corresponding data point.
        labels->set_ShowLeaderLines(true);
        labels->set_ShowCategoryName(false);
        labels->set_ShowPercentage(false);
        labels->set_ShowSeriesName(true);
        labels->set_ShowValue(true);
        labels->set_Separator(u"/");
        labels->set_ShowValue(true);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.ChartDataLabel.docx");
        //ExEnd:WorkWithChartDataLabel
    }

    void DefaultOptionsForDataLabels()
    {
        //ExStart:DefaultOptionsForDataLabels
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Pie, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        chart->get_Series()->Clear();

        SharedPtr<ChartSeries> series =
            chart->get_Series()->Add(u"Aspose Series 1", MakeArray<String>({u"Category 1", u"Category 2", u"Category 3"}), MakeArray<double>({2.7, 3.2, 0.8}));

        SharedPtr<ChartDataLabelCollection> labels = series->get_DataLabels();
        labels->set_ShowPercentage(true);
        labels->set_ShowValue(true);
        labels->set_ShowLeaderLines(false);
        labels->set_Separator(u" - ");

        doc->Save(ArtifactsDir + u"WorkingWithCharts.DefaultOptionsForDataLabels.docx");
        //ExEnd:DefaultOptionsForDataLabels
    }

    void SingleChartDataPoint()
    {
        //ExStart:WorkWithSingleChartDataPoint
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();
        SharedPtr<ChartSeries> series0 = chart->get_Series()->idx_get(0);
        SharedPtr<ChartSeries> series1 = chart->get_Series()->idx_get(1);

        SharedPtr<ChartDataPointCollection> dataPointCollection = series0->get_DataPoints();

        dataPointCollection->idx_get(0)->set_Explosion(50);
        dataPointCollection->idx_get(0)->get_Marker()->set_Symbol(MarkerSymbol::Circle);
        dataPointCollection->idx_get(0)->get_Marker()->set_Size(15);

        dataPointCollection->idx_get(1)->get_Marker()->set_Symbol(MarkerSymbol::Diamond);
        dataPointCollection->idx_get(1)->get_Marker()->set_Size(20);

        series1->get_DataPoints()->idx_get(2)->set_InvertIfNegative(true);
        series1->get_DataPoints()->idx_get(2)->get_Marker()->set_Symbol(MarkerSymbol::Star);
        series1->get_DataPoints()->idx_get(2)->get_Marker()->set_Size(20);

        doc->Save(ArtifactsDir + u"WorkingWithCharts.SingleChartDataPoint.docx");
        //ExEnd:WorkWithSingleChartDataPoint
    }

    void SingleChartSeries()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        SharedPtr<Shape> shape = builder->InsertChart(ChartType::Line, 432, 252);

        SharedPtr<Chart> chart = shape->get_Chart();

        //ExStart:WorkWithSingleChartSeries
        SharedPtr<ChartSeries> series0 = chart->get_Series()->idx_get(0);
        SharedPtr<ChartSeries> series1 = chart->get_Series()->idx_get(1);

        series0->set_Name(u"Chart Series Name 1");
        series1->set_Name(u"Chart Series Name 2");

        // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
        series0->set_Smooth(true);
        series1->set_Smooth(true);
        //ExEnd:WorkWithSingleChartSeries

        //ExStart:ChartDataPoint
        // Specifies whether by default the parent element shall inverts its colors if the value is negative.
        series0->set_InvertIfNegative(true);

        series0->get_Marker()->set_Symbol(MarkerSymbol::Circle);
        series0->get_Marker()->set_Size(15);

        series1->get_Marker()->set_Symbol(MarkerSymbol::Star);
        series1->get_Marker()->set_Size(10);
        //ExEnd:ChartDataPoint

        doc->Save(ArtifactsDir + u"WorkingWithCharts.SingleChartSeries.docx");
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements
