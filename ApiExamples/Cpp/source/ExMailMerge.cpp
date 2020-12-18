// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
#include "ExMailMerge.h"

#include <system/string.h>

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

using TrimWhiteSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::TrimWhiteSpaces)>::type;

struct TrimWhiteSpacesVP : public ExMailMerge, public ApiExamples::ExMailMerge, public ::testing::WithParamInterface<TrimWhiteSpaces_Args>
{
    static std::vector<TrimWhiteSpaces_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(TrimWhiteSpacesVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->TrimWhiteSpaces(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExMailMerge, TrimWhiteSpacesVP, ::testing::ValuesIn(TrimWhiteSpacesVP::TestCases()));

TEST_F(ExMailMerge, MailMergeGetFieldNames)
{
    s_instance->MailMergeGetFieldNames();
}

TEST_F(ExMailMerge, DeleteFields)
{
    s_instance->DeleteFields();
}

TEST_F(ExMailMerge, RemoveContainingFields)
{
    s_instance->RemoveContainingFields();
}

TEST_F(ExMailMerge, RemoveUnusedFields)
{
    s_instance->RemoveUnusedFields();
}

TEST_F(ExMailMerge, RemoveEmptyParagraphs)
{
    s_instance->RemoveEmptyParagraphs();
}

using RemoveColonBetweenEmptyMergeFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::RemoveColonBetweenEmptyMergeFields)>::type;

struct RemoveColonBetweenEmptyMergeFieldsVP : public ExMailMerge,
                                              public ApiExamples::ExMailMerge,
                                              public ::testing::WithParamInterface<RemoveColonBetweenEmptyMergeFields_Args>
{
    static std::vector<RemoveColonBetweenEmptyMergeFields_Args> TestCases()
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

TEST_P(RemoveColonBetweenEmptyMergeFieldsVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->RemoveColonBetweenEmptyMergeFields(get<0>(params), get<1>(params), get<2>(params)));
}

INSTANTIATE_TEST_SUITE_P(DISABLED_ExMailMerge, RemoveColonBetweenEmptyMergeFieldsVP, ::testing::ValuesIn(RemoveColonBetweenEmptyMergeFieldsVP::TestCases()));

TEST_F(ExMailMerge, GetFieldNames)
{
    s_instance->GetFieldNames();
}

TEST_F(ExMailMerge, UseNonMergeFields)
{
    s_instance->UseNonMergeFields();
}

TEST_F(ExMailMerge, TestMailMergeGetRegionsHierarchy)
{
    s_instance->TestMailMergeGetRegionsHierarchy();
}

TEST_F(ExMailMerge, TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption)
{
    s_instance->TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption();
}

TEST_F(ExMailMerge, GetRegionsByName)
{
    s_instance->GetRegionsByName();
}

}} // namespace ApiExamples::gtest_test
