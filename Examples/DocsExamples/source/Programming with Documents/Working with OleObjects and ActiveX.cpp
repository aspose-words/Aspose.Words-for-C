#include "Working with OleObjects and ActiveX.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Drawing::Ole;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithOleObjectsAndActiveX : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithOleObjectsAndActiveX> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithOleObjectsAndActiveX>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithOleObjectsAndActiveX> WorkingWithOleObjectsAndActiveX::s_instance;

TEST_F(WorkingWithOleObjectsAndActiveX, InsertOleObject)
{
    s_instance->InsertOleObject();
}

TEST_F(WorkingWithOleObjectsAndActiveX, InsertOleObjectWithOlePackage)
{
    s_instance->InsertOleObjectWithOlePackage();
}

TEST_F(WorkingWithOleObjectsAndActiveX, InsertOleObjectAsIcon)
{
    s_instance->InsertOleObjectAsIcon();
}

TEST_F(WorkingWithOleObjectsAndActiveX, InsertOleObjectAsIconUsingStream)
{
    s_instance->InsertOleObjectAsIconUsingStream();
}

TEST_F(WorkingWithOleObjectsAndActiveX, ReadActiveXControlProperties)
{
    s_instance->ReadActiveXControlProperties();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
