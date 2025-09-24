// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ApiExampleBase.h"

#include <system/uri.h>
#include <system/threading/thread.h>
#include <system/io/search_option.h>
#include <system/io/path.h>
#include <system/io/file.h>
#include <system/io/directory_info.h>
#include <system/io/directory.h>
#include <system/globalization/culture_info.h>
#include <system/array.h>
#include <Aspose.Words.Cpp/Licensing/License.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4165834238u, ::Aspose::Words::ApiExamples::ApiExampleBase, ThisTypeBaseTypesInfo);

System::String ApiExampleBase::pr_AssemblyDir;
System::String ApiExampleBase::pr_CodeBaseDir;
System::String ApiExampleBase::pr_LicenseDir;
System::String ApiExampleBase::pr_ArtifactsDir;
System::String ApiExampleBase::pr_GoldsDir;
System::String ApiExampleBase::pr_MyDir;
System::String ApiExampleBase::pr_ImageDir;
System::String ApiExampleBase::pr_DatabaseDir;
System::String ApiExampleBase::pr_FontsDir;
System::String ApiExampleBase::pr_ImageUrl;

System::String ApiExampleBase::get_AssemblyDir()
{
    __StaticConstructor__();
    return pr_AssemblyDir;
}

System::String ApiExampleBase::get_CodeBaseDir()
{
    __StaticConstructor__();
    return pr_CodeBaseDir;
}

System::String ApiExampleBase::get_LicenseDir()
{
    __StaticConstructor__();
    return pr_LicenseDir;
}

System::String ApiExampleBase::get_ArtifactsDir()
{
    __StaticConstructor__();
    return pr_ArtifactsDir;
}

System::String ApiExampleBase::get_GoldsDir()
{
    __StaticConstructor__();
    return pr_GoldsDir;
}

System::String ApiExampleBase::get_MyDir()
{
    __StaticConstructor__();
    return pr_MyDir;
}

System::String ApiExampleBase::get_ImageDir()
{
    __StaticConstructor__();
    return pr_ImageDir;
}

System::String ApiExampleBase::get_DatabaseDir()
{
    __StaticConstructor__();
    return pr_DatabaseDir;
}

System::String ApiExampleBase::get_FontsDir()
{
    __StaticConstructor__();
    return pr_FontsDir;
}

System::String ApiExampleBase::get_ImageUrl()
{
    __StaticConstructor__();
    return pr_ImageUrl;
}

void ApiExampleBase::OneTimeSetUp()
{
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::Globalization::CultureInfo::get_InvariantCulture());
    SetUnlimitedLicense();
    
    if (!System::IO::Directory::Exists(get_ArtifactsDir()))
    {
        System::IO::Directory::CreateDirectory_(get_ArtifactsDir());
    }
}

void ApiExampleBase::SetUp()
{
}

void ApiExampleBase::OneTimeTearDown()
{
    // Do not delete the artifacts folder so that you can use a symbolic link to another drive.
    if (System::IO::Directory::Exists(get_ArtifactsDir()))
    {
        for (System::String file : System::IO::Directory::GetFiles(get_ArtifactsDir(), u"*.*", System::IO::SearchOption::AllDirectories))
        {
            System::IO::File::Delete(file);
        }
        
        
        for (System::String subDir : System::IO::Directory::GetDirectories(get_ArtifactsDir(), u"*", System::IO::SearchOption::AllDirectories))
        {
            System::IO::Directory::Delete(subDir, true);
        }
        
    }
}

void ApiExampleBase::SetUnlimitedLicense()
{
    __StaticConstructor__();
    // This is where the test license is on my development machine.
    System::String testLicenseFileName = u"Aspose.Words.Cpp.lic";
    System::String testLicenseFilePath = System::IO::Path::Combine(get_LicenseDir(), testLicenseFileName);
    
    if (System::IO::File::Exists(testLicenseFilePath))
    {
        // This shows how to use an Aspose.Words license when you have purchased one.
        // You don't have to specify full path as shown here. You can specify just the
        // file name if you copy the license file into the same folder as your application
        // binaries or you add the license to your project as an embedded resource.
        auto wordsLicense = System::MakeObject<Aspose::Words::License>();
        wordsLicense->SetLicense(testLicenseFilePath);
    }
}

System::String ApiExampleBase::GetCodeBaseDir(System::SharedPtr<System::Reflection::Assembly> assembly)
{
    __StaticConstructor__();
    System::String mainFolder;
    System::SharedPtr<System::IO::DirectoryInfo> currentDir = System::MakeObject<System::IO::DirectoryInfo>(GetAssemblyDir(assembly))->get_Parent();
    while (currentDir != nullptr)
    {
        if (System::IO::Directory::Exists(System::IO::Path::Combine(currentDir->get_FullName(), u"Data")))
        {
            mainFolder = currentDir->get_FullName() + System::IO::Path::DirectorySeparatorChar;
            break;
        }
        currentDir = currentDir->get_Parent();
    }
    return mainFolder;
}

System::String ApiExampleBase::GetAssemblyDir(System::SharedPtr<System::Reflection::Assembly> assembly)
{
    __StaticConstructor__();
    // CodeBase is a full URI, such as file:///x:\blahblah.
    auto uri = System::MakeObject<System::Uri>(assembly->get_Location());
    return System::IO::Path::GetDirectoryName(uri->get_LocalPath()) + System::IO::Path::DirectorySeparatorChar;
}

void ApiExampleBase::__StaticConstructor__()
{
    thread_local static bool inProgress = false;
    if (inProgress) return;
    static const auto markFinished = [](bool *inProgress) { *inProgress = false; };
    inProgress = true;
    const std::unique_ptr<bool, decltype(markFinished)> finish(&inProgress, markFinished);
    
    static std::once_flag once;
    std::call_once(once, []
    {
        Aspose::Words::ApiExamples::ApiExampleBase::pr_AssemblyDir = Aspose::Words::ApiExamples::ApiExampleBase::GetAssemblyDir(System::Reflection::Assembly::GetExecutingAssembly());
        Aspose::Words::ApiExamples::ApiExampleBase::pr_CodeBaseDir = Aspose::Words::ApiExamples::ApiExampleBase::GetCodeBaseDir(System::Reflection::Assembly::GetExecutingAssembly());
        Aspose::Words::ApiExamples::ApiExampleBase::pr_ArtifactsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/Artifacts/")->get_LocalPath();
        System::String licenseDirShift = u"License/";
        Aspose::Words::ApiExamples::ApiExampleBase::pr_LicenseDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), licenseDirShift)->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_GoldsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/Golds/")->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_MyDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/")->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_ImageDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/Images/")->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_DatabaseDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/Database/")->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_FontsDir = System::MakeObject<System::Uri>(System::MakeObject<System::Uri>(Aspose::Words::ApiExamples::ApiExampleBase::get_CodeBaseDir()), u"Data/MyFonts/")->get_LocalPath();
        Aspose::Words::ApiExamples::ApiExampleBase::pr_ImageUrl = System::MakeObject<System::Uri>(u"https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png")->get_AbsoluteUri();
    });
}

ApiExampleBase::ApiExampleBase()
{
    __StaticConstructor__();
    
}


namespace gtest_test
{

class ApiExampleBase : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ApiExampleBase> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ApiExampleBase>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ApiExampleBase> ApiExampleBase::s_instance;

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
