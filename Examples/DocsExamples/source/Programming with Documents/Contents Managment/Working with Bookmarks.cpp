#include "Working with Bookmarks.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace DocsExamples { namespace Programming_with_Documents { namespace Contents_Managment { namespace gtest_test {

class WorkingWithBookmarks : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithBookmarks> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithBookmarks>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::Contents_Managment::WorkingWithBookmarks> WorkingWithBookmarks::s_instance;

TEST_F(WorkingWithBookmarks, AccessBookmarks)
{
    s_instance->AccessBookmarks();
}

TEST_F(WorkingWithBookmarks, UpdateBookmarkData)
{
    s_instance->UpdateBookmarkData();
}

TEST_F(WorkingWithBookmarks, BookmarkTableColumns)
{
    s_instance->BookmarkTableColumns();
}

TEST_F(WorkingWithBookmarks, CopyBookmarkedText)
{
    s_instance->CopyBookmarkedText();
}

TEST_F(WorkingWithBookmarks, CreateBookmark)
{
    s_instance->CreateBookmark();
}

TEST_F(WorkingWithBookmarks, ShowHideBookmarks)
{
    s_instance->ShowHideBookmarks();
}

TEST_F(WorkingWithBookmarks, UntangleRowBookmarks)
{
    s_instance->UntangleRowBookmarks();
}

}}}} // namespace DocsExamples::Programming_with_Documents::Contents_Managment::gtest_test
