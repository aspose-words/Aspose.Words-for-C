#include "DocsExamplesBase.h"

#include <cstdint>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/path.h>
#include <system/string_comparison.h>
#include <system/threading/thread.h>
#include <system/uri.h>

namespace DocsExamples {

System::String DocsExamplesBase::MainDataDir;
System::String DocsExamplesBase::MyDir;
System::String DocsExamplesBase::ImagesDir;
System::String DocsExamplesBase::DatabaseDir;
System::String DocsExamplesBase::ArtifactsDir;
System::String DocsExamplesBase::FontsDir;

DocsExamplesBase::__StaticConstructor__ DocsExamplesBase::s_constructor__;

DocsExamplesBase::__StaticConstructor__::__StaticConstructor__()
{
    DocsExamplesBase::MainDataDir = DocsExamplesBase::GetCodeBaseDir(System::Reflection::Assembly::GetExecutingAssembly());
    DocsExamplesBase::ArtifactsDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/Artifacts/")->get_LocalPath();
    DocsExamplesBase::MyDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/")->get_LocalPath();
    DocsExamplesBase::ImagesDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/Images/")->get_LocalPath();
    DocsExamplesBase::DatabaseDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/Database/")->get_LocalPath();
    DocsExamplesBase::FontsDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/MyFonts/")->get_LocalPath();
}

void DocsExamplesBase::OneTimeSetUp()
{
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::Globalization::CultureInfo::get_InvariantCulture());

    if (!System::IO::Directory::Exists(ArtifactsDir))
    {
        System::IO::Directory::CreateDirectory_(ArtifactsDir);
    }
}

void DocsExamplesBase::SetUp()
{
    // Console.WriteLine($"Clr: {RuntimeInformation.FrameworkDescription}\n");
}

void DocsExamplesBase::OneTimeTearDown()
{
    if (System::IO::Directory::Exists(ArtifactsDir))
    {
        System::IO::Directory::Delete(ArtifactsDir, true);
    }
}

System::String DocsExamplesBase::GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly)
{
    auto uri = System::MakeObject<System::Uri>(assembly->get_CodeBase());
    System::String mainFolder =
        System::IO::Path::GetDirectoryName(uri->get_LocalPath()).Substring(0, uri->get_LocalPath().IndexOf(u"DocsExamples", System::StringComparison::Ordinal));

    return mainFolder;
}

namespace gtest_test {

class DocsExamplesBase : public ::testing::Test
{
protected:
    static System::SharedPtr<::DocsExamples::DocsExamplesBase> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::DocsExamples::DocsExamplesBase>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::DocsExamples::DocsExamplesBase> DocsExamplesBase::s_instance;

} // namespace gtest_test

} // namespace DocsExamples
