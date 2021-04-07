#include "Working with SDT.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment { namespace gtest_test {

class WorkingWithSdt : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithSdt> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithSdt>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithSdt> WorkingWithSdt::s_instance;

TEST_F(WorkingWithSdt, CheckBoxTypeContentControl)
{
    s_instance->CheckBoxTypeContentControl();
}

TEST_F(WorkingWithSdt, CurrentStateOfCheckBox)
{
    s_instance->CurrentStateOfCheckBox();
}

TEST_F(WorkingWithSdt, ModifyContentControls)
{
    s_instance->ModifyContentControls();
}

TEST_F(WorkingWithSdt, ComboBoxContentControl)
{
    s_instance->ComboBoxContentControl();
}

TEST_F(WorkingWithSdt, RichTextBoxContentControl)
{
    s_instance->RichTextBoxContentControl();
}

TEST_F(WorkingWithSdt, SetContentControlColor)
{
    s_instance->SetContentControlColor();
}

TEST_F(WorkingWithSdt, ClearContentsControl)
{
    s_instance->ClearContentsControl();
}

TEST_F(WorkingWithSdt, BindSdTtoCustomXmlPart)
{
    s_instance->BindSdTtoCustomXmlPart();
}

TEST_F(WorkingWithSdt, SetContentControlStyle)
{
    s_instance->SetContentControlStyle();
}

TEST_F(WorkingWithSdt, CreatingTableRepeatingSectionMappedToCustomXmlPart)
{
    s_instance->CreatingTableRepeatingSectionMappedToCustomXmlPart();
}

TEST_F(WorkingWithSdt, MultiSection)
{
    s_instance->MultiSection();
}

TEST_F(WorkingWithSdt, StructuredDocumentTagRangeStartXmlMapping)
{
    s_instance->StructuredDocumentTagRangeStartXmlMapping();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment::gtest_test
