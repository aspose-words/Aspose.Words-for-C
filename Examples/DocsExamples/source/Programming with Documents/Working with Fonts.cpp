#include "Working with Fonts.h"

using namespace Aspose::Words;
using namespace Aspose::Words::Fonts;
using namespace Aspose::Words::Loading;
namespace DocsExamples { namespace Programming_with_Documents { namespace gtest_test {

class WorkingWithFonts : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFonts> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::Programming_with_Documents::WorkingWithFonts>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::Programming_with_Documents::WorkingWithFonts> WorkingWithFonts::s_instance;

TEST_F(WorkingWithFonts, FontFormatting)
{
    s_instance->FontFormatting();
}

TEST_F(WorkingWithFonts, GetFontLineSpacing)
{
    s_instance->GetFontLineSpacing();
}

TEST_F(WorkingWithFonts, CheckDMLTextEffect)
{
    s_instance->CheckDMLTextEffect();
}

TEST_F(WorkingWithFonts, SetFontFormatting)
{
    s_instance->SetFontFormatting();
}

TEST_F(WorkingWithFonts, SetFontEmphasisMark)
{
    s_instance->SetFontEmphasisMark();
}

TEST_F(WorkingWithFonts, SetFontsFolders)
{
    s_instance->SetFontsFolders();
}

TEST_F(WorkingWithFonts, EnableDisableFontSubstitution)
{
    s_instance->EnableDisableFontSubstitution();
}

TEST_F(WorkingWithFonts, SetFontFallbackSettings)
{
    s_instance->SetFontFallbackSettings();
}

TEST_F(WorkingWithFonts, NotoFallbackSettings)
{
    s_instance->NotoFallbackSettings();
}

TEST_F(WorkingWithFonts, SetFontsFoldersDefaultInstance)
{
    s_instance->SetFontsFoldersDefaultInstance();
}

TEST_F(WorkingWithFonts, SetFontsFoldersMultipleFolders)
{
    s_instance->SetFontsFoldersMultipleFolders();
}

TEST_F(WorkingWithFonts, SetFontsFoldersSystemAndCustomFolder)
{
    s_instance->SetFontsFoldersSystemAndCustomFolder();
}

TEST_F(WorkingWithFonts, SetFontsFoldersWithPriority)
{
    s_instance->SetFontsFoldersWithPriority();
}

TEST_F(WorkingWithFonts, SetTrueTypeFontsFolder)
{
    s_instance->SetTrueTypeFontsFolder();
}

TEST_F(WorkingWithFonts, SpecifyDefaultFontWhenRendering)
{
    s_instance->SpecifyDefaultFontWhenRendering();
}

TEST_F(WorkingWithFonts, FontSettingsWithLoadOptions)
{
    s_instance->FontSettingsWithLoadOptions();
}

TEST_F(WorkingWithFonts, SetFontsFolder)
{
    s_instance->SetFontsFolder();
}

TEST_F(WorkingWithFonts, FontSettingsWithLoadOption)
{
    s_instance->FontSettingsWithLoadOption();
}

TEST_F(WorkingWithFonts, FontSettingsDefaultInstance)
{
    s_instance->FontSettingsDefaultInstance();
}

TEST_F(WorkingWithFonts, GetListOfAvailableFonts)
{
    s_instance->GetListOfAvailableFonts();
}

TEST_F(WorkingWithFonts, ReceiveNotificationsOfFonts)
{
    s_instance->ReceiveNotificationsOfFonts();
}

TEST_F(WorkingWithFonts, ReceiveWarningNotification)
{
    s_instance->ReceiveWarningNotification();
}

TEST_F(WorkingWithFonts, ResourceSteamFontSourceExample)
{
    s_instance->ResourceSteamFontSourceExample();
}

TEST_F(WorkingWithFonts, GetSubstitutionWithoutSuffixes)
{
    s_instance->GetSubstitutionWithoutSuffixes();
}

}}} // namespace DocsExamples::Programming_with_Documents::gtest_test
