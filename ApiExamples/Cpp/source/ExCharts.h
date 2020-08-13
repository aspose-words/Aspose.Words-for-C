#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/test_tools/method_argument_tuple.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Drawing { namespace Charts { enum class ChartType; } } } }
namespace Aspose { namespace Words { namespace Drawing { namespace Charts { class ChartSeries; } } } }
namespace Aspose { namespace Words { namespace Drawing { namespace Charts { enum class MarkerSymbol; } } } }
namespace Aspose { namespace Words { namespace Drawing { namespace Charts { class Chart; } } } }
namespace Aspose { namespace Words { class DocumentBuilder; } }

namespace ApiExamples {

class ExCharts : public ApiExampleBase
{
public:

    void ChartTitle();
    void DefineNumberFormatForDataLabels();
    void DataArraysWrongSize();
    void EmptyValuesInChartData();
    void AxisProperties();
    void DateTimeValues();
    void HideChartAxis();
    void SetNumberFormatToChartAxis();
    void TestDisplayChartsWithConversion(Aspose::Words::Drawing::Charts::ChartType chartType);
    void Surface3DChart();
    void ChartDataLabelCollection();
    void ChartDataLabels();
    void ChartDataPoint();
    void PieChartExplosion();
    void Bubble3D();
    void ChartSeriesCollection();
    void ChartSeriesCollectionModify();
    void AxisScaling();
    void AxisBound();
    void ChartLegend();
    void AxisCross();
    void ChartAxisDisplayUnit();
    
protected:

    /// <summary>
    /// Apply uniform data labels with custom number format and separator to a number (determined by labelsCount) of data points in a series
    /// </summary>
    static void ApplyDataLabels(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t labelsCount, System::String numberFormat, System::String separator);
    /// <summary>
    /// Applies a number of data points to a series
    /// </summary>
    static void ApplyDataPoints(System::SharedPtr<Aspose::Words::Drawing::Charts::ChartSeries> series, int32_t dataPointsCount, Aspose::Words::Drawing::Charts::MarkerSymbol markerSymbol, int32_t dataPointSize);
    /// <summary>
    /// Get the DocumentBuilder to insert a chart of a specified ChartType, width and height and clean out its default data
    /// </summary>
    static System::SharedPtr<Aspose::Words::Drawing::Charts::Chart> AppendChart(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, Aspose::Words::Drawing::Charts::ChartType chartType, double width, double height);
    
};

} // namespace ApiExamples


