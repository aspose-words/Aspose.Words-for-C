#include "Working with Shapes.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;
namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements { namespace gtest_test {

class WorkingWithShapes : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithShapes> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithShapes>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithShapes> WorkingWithShapes::s_instance;

TEST_F(WorkingWithShapes, AddGroupShape)
{
    s_instance->AddGroupShape();
}

TEST_F(WorkingWithShapes, InsertShape)
{
    s_instance->InsertShape();
}

TEST_F(WorkingWithShapes, AspectRatioLocked)
{
    s_instance->AspectRatioLocked();
}

TEST_F(WorkingWithShapes, LayoutInCell)
{
    s_instance->LayoutInCell();
}

TEST_F(WorkingWithShapes, AddCornersSnipped)
{
    s_instance->AddCornersSnipped();
}

TEST_F(WorkingWithShapes, GetActualShapeBoundsPoints)
{
    s_instance->GetActualShapeBoundsPoints();
}

TEST_F(WorkingWithShapes, VerticalAnchor)
{
    s_instance->VerticalAnchor();
}

TEST_F(WorkingWithShapes, DetectSmartArtShape)
{
    s_instance->DetectSmartArtShape();
}

TEST_F(WorkingWithShapes, UpdateSmartArtDrawing)
{
    s_instance->UpdateSmartArtDrawing();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::gtest_test
