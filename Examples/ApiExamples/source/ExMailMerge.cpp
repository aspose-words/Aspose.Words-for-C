// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMailMerge.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words::Fields;
using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples { namespace gtest_test {

class ExMailMerge : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExMailMerge> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExMailMerge>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExMailMerge> ExMailMerge::s_instance;

TEST_F(ExMailMerge, MailMergeRegionInfo)
{
    s_instance->MailMergeRegionInfo_();
}

using ExMailMerge_TrimWhiteSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::TrimWhiteSpaces)>::type;

struct ExMailMerge_TrimWhiteSpaces : public ExMailMerge, public ApiExamples::ExMailMerge, public ::testing::WithParamInterface<ExMailMerge_TrimWhiteSpaces_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExMailMerge_TrimWhiteSpaces, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrimWhiteSpaces(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExMailMerge_TrimWhiteSpaces, ::testing::ValuesIn(ExMailMerge_TrimWhiteSpaces::TestCases()));

TEST_F(ExMailMerge, DeleteFields)
{
    s_instance->DeleteFields();
}

using ExMailMerge_RemoveColonBetweenEmptyMergeFields_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::RemoveColonBetweenEmptyMergeFields)>::type;

struct ExMailMerge_RemoveColonBetweenEmptyMergeFields : public ExMailMerge,
                                                        public ApiExamples::ExMailMerge,
                                                        public ::testing::WithParamInterface<ExMailMerge_RemoveColonBetweenEmptyMergeFields_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(u"!", false, u""),           std::make_tuple(u", ", false, u""),        std::make_tuple(u" . ", false, u""),
            std::make_tuple(u" :", false, u""),          std::make_tuple(u"  ; ", false, u""),      std::make_tuple(u" ?  ", false, u""),
            std::make_tuple(u"  ¡  ", false, u""),       std::make_tuple(u"  ¿  ", false, u""),     std::make_tuple(u"!", true, u"!\f"),
            std::make_tuple(u", ", true, u", \f"),       std::make_tuple(u" . ", true, u" . \f"),   std::make_tuple(u" :", true, u" :\f"),
            std::make_tuple(u"  ; ", true, u"  ; \f"),   std::make_tuple(u" ?  ", true, u" ?  \f"), std::make_tuple(u"  ¡  ", true, u"  ¡  \f"),
            std::make_tuple(u"  ¿  ", true, u"  ¿  \f"),
        };
    }
};

TEST_P(ExMailMerge_RemoveColonBetweenEmptyMergeFields, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemoveColonBetweenEmptyMergeFields(std::get<0>(params), std::get<1>(params), std::get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_, ExMailMerge_RemoveColonBetweenEmptyMergeFields,
                         ::testing::ValuesIn(ExMailMerge_RemoveColonBetweenEmptyMergeFields::TestCases()));

TEST_F(ExMailMerge, GetFieldNames)
{
    s_instance->GetFieldNames();
}

TEST_F(ExMailMerge, TestMailMergeGetRegionsHierarchy)
{
    s_instance->TestMailMergeGetRegionsHierarchy();
}

TEST_F(ExMailMerge, GetRegionsByName)
{
    s_instance->GetRegionsByName();
}

TEST_F(ExMailMerge, RestartListsAtEachSection)
{
    s_instance->RestartListsAtEachSection();
}

}} // namespace ApiExamples::gtest_test
