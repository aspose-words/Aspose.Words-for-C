// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ApiExampleBase.h"

#include <testing/test_predicates.h>
#include <system/uri.h>
#include <system/threading/thread.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string_comparison.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/reflection/assembly.h>
#include <system/object.h>
#include <system/io/path.h>
#include <system/io/file.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <gtest/gtest-skip-exception.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Licensing/License.h>

using namespace Aspose::Words;
namespace ApiExamples {

System::String ApiExampleBase::AssemblyDir;
System::String ApiExampleBase::CodeBaseDir;
System::String ApiExampleBase::LicenseDir;
System::String ApiExampleBase::ArtifactsDir;
System::String ApiExampleBase::GoldsDir;
System::String ApiExampleBase::MyDir;
System::String ApiExampleBase::ImageDir;
System::String ApiExampleBase::DatabaseDir;
System::String ApiExampleBase::FontsDir;
System::String ApiExampleBase::AsposeLogoUrl;

ApiExampleBase::__StaticConstructor__ ApiExampleBase::s_constructor__;

void ApiExampleBase::OneTimeSetUp()
{
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::Globalization::CultureInfo::get_InvariantCulture());

    if (CheckForSkipMono() && IsRunningOnMono())
    {
        throw testing::SkipException();
    }

    if (!CheckForSkipSetUp())
    {
        SetUnlimitedLicense();
    }

    if (!System::IO::Directory::Exists(ArtifactsDir))
    {
        System::IO::Directory::CreateDirectory_(ArtifactsDir);
    }
}

void ApiExampleBase::SetUp()
{
}

void ApiExampleBase::OneTimeTearDown()
{
    if (!CheckForSkipTearDown())
    {
        if (System::IO::Directory::Exists(ArtifactsDir))
        {
            System::IO::Directory::Delete(ArtifactsDir, true);
        }
    }
}

bool ApiExampleBase::CheckForSkipSetUp()
{
    // CPPDEFECT: Unit-test properties is not supported
    return false;
}

bool ApiExampleBase::CheckForSkipTearDown()
{
    // CPPDEFECT: Unit-test properties is not supported
    return false;
}

bool ApiExampleBase::CheckForSkipMono()
{
    // CPPDEFECT: Unit-test properties is not supported
    return false;
}

bool ApiExampleBase::IsRunningOnMono()
{
    // CPPDEFECT: Type.GetType not implemented
    return false;
}

void ApiExampleBase::SetUnlimitedLicense()
{
    // This is where the test license is on my development machine.
    System::String testLicenseFileName = System::IO::Path::Combine(LicenseDir, u"Aspose.Words.Cpp.lic");

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

System::String ApiExampleBase::GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly)
{
    // CodeBase is a full URI, such as file:///x:\blahblah.
    auto uri = System::MakeObject<System::Uri>(assembly->get_CodeBase());
    System::String mainFolder = System::IO::Path::GetDirectoryName(uri->get_LocalPath());
    if (mainFolder != nullptr)
    {
        mainFolder = mainFolder.Substring(0, uri->get_LocalPath().IndexOf(u"Cpp", System::StringComparison::Ordinal));
    }
    return mainFolder;
}

System::String ApiExampleBase::GetAssemblyDir(System::SharedPtr<System::Reflection::Assembly> assembly)
{
    // CodeBase is a full URI, such as file:///x:\blahblah.
    auto uri = System::MakeObject<System::Uri>(assembly->get_CodeBase());
    return System::IO::Path::GetDirectoryName(uri->get_LocalPath()) + System::IO::Path::DirectorySeparatorChar;
}

ApiExampleBase::__StaticConstructor__::__StaticConstructor__()
{
    ApiExampleBase::AssemblyDir = ApiExampleBase::GetAssemblyDir(System::Reflection::Assembly::GetExecutingAssembly());
    ApiExampleBase::CodeBaseDir = ApiExampleBase::GetCodeBaseDir(System::Reflection::Assembly::GetExecutingAssembly());
    ApiExampleBase::ArtifactsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/Artifacts/")->get_LocalPath();
    ApiExampleBase::LicenseDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"License/")->get_LocalPath();
    ApiExampleBase::GoldsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/Golds/")->get_LocalPath();
    ApiExampleBase::MyDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/")->get_LocalPath();
    ApiExampleBase::ImageDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/Images/")->get_LocalPath();
    ApiExampleBase::DatabaseDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/Database/")->get_LocalPath();
    ApiExampleBase::FontsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(ApiExampleBase::CodeBaseDir), u"Data/MyFonts/")->get_LocalPath();
    ApiExampleBase::AsposeLogoUrl = System::MakeObject<System::Uri>(u"https://www.aspose.cloud/templates/aspose/App_Themes/V3/images/words/header/aspose_words-for-net.png")->get_AbsoluteUri();
}

namespace gtest_test
{

class ApiExampleBase : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ApiExampleBase> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ApiExampleBase>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

System::SharedPtr<::ApiExamples::ApiExampleBase> ApiExampleBase::s_instance;

} // namespace gtest_test

} // namespace ApiExamples
