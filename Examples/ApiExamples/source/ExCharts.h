#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/MarkerSymbol.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExCharts : public ApiExampleBase
{
    typedef ExCharts ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void ChartTitle();
    void DataLabelNumberFormat();
    void AxisProperties();
    void AxisCollection();
    void DateTimeValues();
    void HideChartAxis();
    void SetNumberFormatToChartAxis();
    void TestDisplayChartsWithConversion(Aspose::Words::Drawing::Charts::ChartType chartType);
    void Surface3DChart();
    void DataLabelsBubbleChart();
    void DataLabelsPieChart();
    void DataLabels();
    void ChartDataPoint();
    void PieChartExplosion();
    void Bubble3D();
    void ChartSeriesCollection();
    void ChartSeriesCollectionModify();
    void AxisScaling();
    void AxisBound();
    void ChartLegend();
    void AxisCross();
    void AxisDisplayUnit();
    void MarkerFormatting();
    void SeriesColor();
    void DataPointsFormatting();
    void LegendEntries();
    void LegendFont();
    void RemoveSpecificChartSeries();
    void PopulateChartWithData();
    void GetChartSeriesData();
    void ChartDataValues();
    void FormatDataLables();
    void ChartAxisTitle();
    void CopyDataPointFormat();
    void ResetDataPointFill();
    void DataTable();
    void ChartFormat();
    void SecondaryAxis();
    void ConfigureGapOverlap();
    void BubbleScale();
    void RemoveSecondaryAxis();
    void TreemapChart();
    void SunburstChart();
    void HistogramChart();
    void ParetoChart();
    void BoxAndWhiskerChart();
    void WaterfallChart();
    void FunnelChart();
    void LabelOrientationRotation();
    void TickLabelsOrientationRotation();
    void DoughnutChart();
    void PieOfPieChart();
    void FormatCode();
    void DataLablePosition();
    void DoughnutChartLabelPosition();
    void InsertChartSeries();
    void SetChartStyle();
    
protected:

    /// <summary>
    /// Apply data labels with custom number format and separator to several data points in a series.
    /// </summary>
    static void ApplyDataLabels(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t labelsCount, System::String numberFormat, System::String separator);
    /// <summary>
    /// Applies a number of data points to a series.
    /// </summary>
    static void ApplyDataPoints(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t dataPointsCount, Aspose::Words::Drawing::Charts::MarkerSymbol markerSymbol, int32_t dataPointSize);
    /// <summary>
    /// Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
    /// </summary>
    static System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> AppendChart(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::Drawing::Charts::ChartType chartType, double width, double height);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


