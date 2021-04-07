#include "Working with Charts.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Charts;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements { namespace gtest_test {

class WorkingWithCharts : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithCharts> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithCharts>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithCharts> WorkingWithCharts::s_instance;

TEST_F(WorkingWithCharts, FormatNumberOfDataLabel)
{
    s_instance->FormatNumberOfDataLabel();
}

TEST_F(WorkingWithCharts, CreateChartUsingShape)
{
    s_instance->CreateChartUsingShape();
}

TEST_F(WorkingWithCharts, InsertSimpleColumnChart)
{
    s_instance->InsertSimpleColumnChart();
}

TEST_F(WorkingWithCharts, InsertColumnChart)
{
    s_instance->InsertColumnChart();
}

TEST_F(WorkingWithCharts, InsertAreaChart)
{
    s_instance->InsertAreaChart();
}

TEST_F(WorkingWithCharts, InsertBubbleChart)
{
    s_instance->InsertBubbleChart();
}

TEST_F(WorkingWithCharts, InsertScatterChart)
{
    s_instance->InsertScatterChart();
}

TEST_F(WorkingWithCharts, DefineXYAxisProperties)
{
    s_instance->DefineXYAxisProperties();
}

TEST_F(WorkingWithCharts, DateTimeValuesToAxis)
{
    s_instance->DateTimeValuesToAxis();
}

TEST_F(WorkingWithCharts, NumberFormatForAxis)
{
    s_instance->NumberFormatForAxis();
}

TEST_F(WorkingWithCharts, BoundsOfAxis)
{
    s_instance->BoundsOfAxis();
}

TEST_F(WorkingWithCharts, IntervalUnitBetweenLabelsOnAxis)
{
    s_instance->IntervalUnitBetweenLabelsOnAxis();
}

TEST_F(WorkingWithCharts, HideChartAxis)
{
    s_instance->HideChartAxis();
}

TEST_F(WorkingWithCharts, TickMultiLineLabelAlignment)
{
    s_instance->TickMultiLineLabelAlignment();
}

TEST_F(WorkingWithCharts, ChartDataLabel)
{
    s_instance->ChartDataLabel();
}

TEST_F(WorkingWithCharts, DefaultOptionsForDataLabels)
{
    s_instance->DefaultOptionsForDataLabels();
}

TEST_F(WorkingWithCharts, SingleChartDataPoint)
{
    s_instance->SingleChartDataPoint();
}

TEST_F(WorkingWithCharts, SingleChartSeries)
{
    s_instance->SingleChartSeries();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::gtest_test
