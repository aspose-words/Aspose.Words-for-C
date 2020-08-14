// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildVersion.h"

#include <system/string.h>
#include <system/console.h>
#include <Aspose.Words.Cpp/Model/Document/BuildVersionInfo.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
namespace ApiExamples {

namespace gtest_test
{

class ExBuildVersion : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExBuildVersion> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExBuildVersion>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExBuildVersion> ExBuildVersion::s_instance;

} // namespace gtest_test

void ExBuildVersion::PrintBuildVersionInfo()
{
    //ExStart
    //ExFor:BuildVersionInfo
    //ExFor:BuildVersionInfo.Product
    //ExFor:BuildVersionInfo.Version
    //ExSummary:Shows how to use BuildVersionInfo to display version information about this product.
    System::Console::WriteLine(String::Format(u"I am currently using {0}, version number {1}!",BuildVersionInfo::get_Product(),BuildVersionInfo::get_Version()));
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExBuildVersion, PrintBuildVersionInfo)
{
    s_instance->PrintBuildVersionInfo();
}

} // namespace gtest_test

} // namespace ApiExamples
