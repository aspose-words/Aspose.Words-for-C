// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExCharts.h"

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/Chart.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartSeries.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/ChartType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Charts/MarkerSymbol.h>
#include <system/string.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
namespace ApiExamples { namespace gtest_test {

class ExCharts : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExCharts> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExCharts>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExCharts> ExCharts::s_instance;

TEST_F(ExCharts, ChartTitle)
{
    s_instance->ChartTitle_();
}

TEST_F(ExCharts, DataLabelNumberFormat)
{
    s_instance->DataLabelNumberFormat();
}

TEST_F(ExCharts, DataArraysWrongSize)
{
    s_instance->DataArraysWrongSize();
}

TEST_F(ExCharts, EmptyValuesInChartData)
{
    s_instance->EmptyValuesInChartData();
}

TEST_F(ExCharts, AxisProperties)
{
    s_instance->AxisProperties();
}

TEST_F(ExCharts, DateTimeValues)
{
    s_instance->DateTimeValues();
}

TEST_F(ExCharts, HideChartAxis)
{
    s_instance->HideChartAxis();
}

TEST_F(ExCharts, SetNumberFormatToChartAxis)
{
    s_instance->SetNumberFormatToChartAxis();
}

using TestDisplayChartsWithConversion_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExCharts::TestDisplayChartsWithConversion)>::type;

struct TestDisplayChartsWithConversionVP : public ExCharts,
                                           public ApiExamples::ExCharts,
                                           public ::testing::WithParamInterface<TestDisplayChartsWithConversion_Args>
{
    static std::vector<TestDisplayChartsWithConversion_Args> TestCases()
    {
        return {
            std::make_tuple(ChartType::Column), std::make_tuple(ChartType::Line), std::make_tuple(ChartType::Pie),
            std::make_tuple(ChartType::Bar),    std::make_tuple(ChartType::Area),
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

TEST_F(ExCharts, Surface3DChart)
{
    s_instance->Surface3DChart();
}

TEST_F(ExCharts, DataLabelsBubbleChart)
{
    s_instance->DataLabelsBubbleChart();
}

TEST_F(ExCharts, DataLabelsPieChart)
{
    s_instance->DataLabelsPieChart();
}

TEST_F(ExCharts, DataLabels)
{
    s_instance->DataLabels();
}

TEST_F(ExCharts, ChartDataPoint)
{
    s_instance->ChartDataPoint_();
}

TEST_F(ExCharts, PieChartExplosion)
{
    s_instance->PieChartExplosion();
}

TEST_F(ExCharts, Bubble3D)
{
    s_instance->Bubble3D();
}

TEST_F(ExCharts, ChartSeriesCollection)
{
    s_instance->ChartSeriesCollection_();
}

TEST_F(ExCharts, ChartSeriesCollectionModify)
{
    s_instance->ChartSeriesCollectionModify();
}

TEST_F(ExCharts, AxisScaling)
{
    s_instance->AxisScaling_();
}

TEST_F(ExCharts, AxisBound)
{
    s_instance->AxisBound_();
}

TEST_F(ExCharts, ChartLegend)
{
    s_instance->ChartLegend_();
}

TEST_F(ExCharts, AxisCross)
{
    s_instance->AxisCross();
}

TEST_F(ExCharts, AxisDisplayUnit)
{
    s_instance->AxisDisplayUnit_();
}

}} // namespace ApiExamples::gtest_test
