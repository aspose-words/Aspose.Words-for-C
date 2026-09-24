#pragma once

#include <system/string.h>
#include <gtest/gtest.h>
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
    //ExStart
    //ExFor:ChartSeries
    //ExFor:ChartSeries.DataLabels
    //ExFor:ChartSeries.DataPoints
    //ExFor:ChartSeries.Name
    //ExFor:ChartSeries.Explosion
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
    //ExFor:ChartDataLabel.Format
    //ExFor:ChartDataLabel.ClearFormat
    //ExFor:ChartDataLabelCollection
    //ExFor:ChartDataLabelCollection.ShowDataLabelsRange
    //ExFor:ChartDataLabelCollection.ClearFormat
    //ExFor:ChartDataLabelCollection.Count
    //ExFor:ChartDataLabelCollection.GetEnumerator
    //ExFor:ChartDataLabelCollection.Item(Int32)
    //ExSummary:Shows how to apply labels to data points in a line chart.
    void DataLabels();
    //ExEnd
    //ExStart
    //ExFor:ChartSeries.Smooth
    //ExFor:ChartSeries.InvertIfNegative
    //ExFor:ChartDataPoint
    //ExFor:ChartDataPoint.Format
    //ExFor:ChartDataPoint.ClearFormat
    //ExFor:ChartDataPoint.Index
    //ExFor:ChartDataPointCollection
    //ExFor:ChartDataPointCollection.ClearFormat
    //ExFor:ChartDataPointCollection.Count
    //ExFor:ChartDataPointCollection.GetEnumerator
    //ExFor:ChartDataPointCollection.Item(Int32)
    //ExFor:ChartMarker
    //ExFor:ChartMarker.Size
    //ExFor:ChartMarker.Symbol
    //ExFor:IChartDataPoint
    //ExFor:IChartDataPoint.InvertIfNegative
    //ExFor:ChartDataPoint.InvertIfNegative
    //ExFor:IChartDataPoint.Marker
    //ExFor:MarkerSymbol
    //ExSummary:Shows how to work with data points on a line chart.
    void ChartDataPoint();
    //ExEnd
    void PieChartExplosion();
    void Bubble3D();
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
    void ChartSeriesCollection();
    //ExEnd
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
    void TitleOrientation();
    
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


