// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExLicense.h"

#include <system/string.h>
#include <system/io/stream.h>
#include <system/io/path.h>
#include <system/io/file_stream.h>
#include <system/io/file.h>
#include <system/details/dispose_guard.h>
#include <Aspose.Words.Cpp/Licensing/License.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1914105745u, ::Aspose::Words::ApiExamples::ExLicense, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExLicense : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExLicense> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExLicense>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExLicense> ExLicense::s_instance;

} // namespace gtest_test

void ExLicense::LicenseFromFileNoPath()
{
    //ExStart
    //ExFor:License
    //ExFor:License.#ctor
    //ExFor:License.SetLicense(String)
    //ExSummary:Shows how to initialize a license for Aspose.Words using a license file in the local file system.
    
    System::String testLicenseFileName = u"Aspose.Words.Cpp.lic";
    
    // Set the license for our Aspose.Words product by passing the local file system filename of a valid license file.
    System::String licenseFileName = System::IO::Path::Combine(get_LicenseDir(), testLicenseFileName);
    
    auto license = System::MakeObject<Aspose::Words::License>();
    license->SetLicense(licenseFileName);
    
    // Create a copy of our license file in the binaries folder of our application.
    System::String licenseCopyFileName = System::IO::Path::Combine(get_AssemblyDir(), testLicenseFileName);
    System::IO::File::Copy(licenseFileName, licenseCopyFileName);
    
    // If we pass a file's name without a path,
    // the SetLicense will search several local file system locations for this file.
    // One of those locations will be the "bin" folder, which contains a copy of our license file.
    license->SetLicense(testLicenseFileName);
    //ExEnd
    
    license->SetLicense(u"");
    System::IO::File::Delete(licenseCopyFileName);
}

namespace gtest_test
{

TEST_F(ExLicense, LicenseFromFileNoPath)
{
    s_instance->LicenseFromFileNoPath();
}

} // namespace gtest_test

void ExLicense::LicenseFromStream()
{
    //ExStart
    //ExFor:License.SetLicense(Stream)
    //ExSummary:Shows how to initialize a license for Aspose.Words from a stream.
    System::String testLicenseFileName = u"Aspose.Words.Cpp.lic";
    // Set the license for our Aspose.Words product by passing a stream for a valid license file in our local file system.
    {
        System::SharedPtr<System::IO::Stream> myStream = System::IO::File::OpenRead(System::IO::Path::Combine(get_LicenseDir(), testLicenseFileName));
        auto license = System::MakeObject<Aspose::Words::License>();
        license->SetLicense(myStream);
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExLicense, LicenseFromStream)
{
    s_instance->LicenseFromStream();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
