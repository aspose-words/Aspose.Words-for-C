// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExBuildVersion.h"

#include <testing/test_predicates.h>
#include <system/text/regularexpressions/regex.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <iostream>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Document/BuildVersionInfo.h>

namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(1502683546u, ::Aspose::Words::ApiExamples::ExBuildVersion, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExBuildVersion : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExBuildVersion> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExBuildVersion>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExBuildVersion> ExBuildVersion::s_instance;

} // namespace gtest_test

void ExBuildVersion::PrintBuildVersionInfo()
{
    //ExStart
    //ExFor:BuildVersionInfo
    //ExFor:BuildVersionInfo.Product
    //ExFor:BuildVersionInfo.Version
    //ExSummary:Shows how to display information about your installed version of Aspose.Words.
    std::cout << System::String::Format(u"I am currently using {0}, version number {1}!", Aspose::Words::BuildVersionInfo::get_Product(), Aspose::Words::BuildVersionInfo::get_Version()) << std::endl;
    //ExEnd
    
    ASSERT_EQ(u"Aspose.Words for C++", Aspose::Words::BuildVersionInfo::get_Product());
    ASSERT_TRUE(System::Text::RegularExpressions::Regex::IsMatch(Aspose::Words::BuildVersionInfo::get_Version(), u"[0-9]{2}.[0-9]{1,2}"));
}

namespace gtest_test
{

TEST_F(ExBuildVersion, PrintBuildVersionInfo)
{
    s_instance->PrintBuildVersionInfo();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
