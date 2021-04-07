#include "Rendering shapes.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Rendering;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Rendering_and_Printing { namespace gtest_test {

class RenderingShapes : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Rendering_and_Printing::RenderingShapes> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Rendering_and_Printing::RenderingShapes>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Rendering_and_Printing::RenderingShapes> RenderingShapes::s_instance;

TEST_F(RenderingShapes, RenderShapeAsEmf)
{
    s_instance->RenderShapeAsEmf();
}

TEST_F(RenderingShapes, RenderShapeAsJpeg)
{
    s_instance->RenderShapeAsJpeg();
}

TEST_F(RenderingShapes, RenderShapeToGraphics)
{
    s_instance->RenderShapeToGraphics();
}

TEST_F(RenderingShapes, RenderCellToImage)
{
    s_instance->RenderCellToImage();
}

TEST_F(RenderingShapes, RenderRowToImage)
{
    s_instance->RenderRowToImage();
}

TEST_F(RenderingShapes, RenderParagraphToImage)
{
    s_instance->RenderParagraphToImage();
}

TEST_F(RenderingShapes, FindShapeSizes)
{
    s_instance->FindShapeSizes();
}

TEST_F(RenderingShapes, RenderShapeImage)
{
    s_instance->RenderShapeImage();
}

}}} // namespace DocsExamples::Rendering_and_Printing::gtest_test
