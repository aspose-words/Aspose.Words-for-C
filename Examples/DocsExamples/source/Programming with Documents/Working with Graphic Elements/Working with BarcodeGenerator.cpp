#include "Working with BarcodeGenerator.h"

#ifdef ASPOSE_BARCODE_AVAILABLE

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Graphic_Elements { namespace gtest_test {

class WorkingWithBarcodeGenerator : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithBarcodeGenerator> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithBarcodeGenerator>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::WorkingWithBarcodeGenerator>
    WorkingWithBarcodeGenerator::s_instance;

TEST_F(WorkingWithBarcodeGenerator, BarcodeGenerator)
{
    s_instance->BarcodeGenerator();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Graphic_Elements::gtest_test

#endif // ASPOSE_BARCODE_AVAILABLE
