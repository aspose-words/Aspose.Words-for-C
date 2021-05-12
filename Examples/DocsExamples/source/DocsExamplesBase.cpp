#include "DocsExamplesBase.h"

#include <cstdint>
#include <Aspose.Words.Cpp/License.h>
#include <system/globalization/culture_info.h>
#include <system/io/directory.h>
#include <system/io/file.h>
#include <system/io/path.h>
#include <system/string_comparison.h>
#include <system/threading/thread.h>
#include <system/uri.h>

using namespace Aspose::Words;
namespace DocsExamples {

System::String DocsExamplesBase::MainDataDir;
System::String DocsExamplesBase::MyDir;
System::String DocsExamplesBase::ImagesDir;
System::String DocsExamplesBase::DatabaseDir;
System::String DocsExamplesBase::LicenseDir;
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
    DocsExamplesBase::LicenseDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/License/")->get_LocalPath();
    DocsExamplesBase::DatabaseDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/Database/")->get_LocalPath();
    DocsExamplesBase::FontsDir =
        System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(DocsExamplesBase::MainDataDir), u"Data/MyFonts/")->get_LocalPath();
}

void DocsExamplesBase::OneTimeSetUp()
{
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::Globalization::CultureInfo::get_InvariantCulture());

    SetUnlimitedLicense();

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

void DocsExamplesBase::SetUnlimitedLicense()
{
    // This is where the test license is on my development machine.
    System::String testLicenseFileName = System::IO::Path::Combine(LicenseDir, u"Aspose.Total.NET.lic");

    if (System::IO::File::Exists(testLicenseFileName))
    {
        // This shows how to use an Aspose.Words license when you have purchased one.
        // You don't have to specify full path as shown here. You can specify just the
        // file name if you copy the license file into the same folder as your application
        // binaries or you add the license to your project as an embedded resource.
        auto wordsLicense = System::MakeObject<License>();
        wordsLicense->SetLicense(testLicenseFileName);
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
