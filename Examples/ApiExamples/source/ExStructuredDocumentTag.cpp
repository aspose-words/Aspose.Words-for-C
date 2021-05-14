// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExStructuredDocumentTag.h"

#include <system/test_tools/method_argument_tuple.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Tables;
namespace ApiExamples { namespace gtest_test {

class ExStructuredDocumentTag : public ::testing::Test
{
protected:
    static System::SharedPtr<::ApiExamples::ExStructuredDocumentTag> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::ApiExamples::ExStructuredDocumentTag>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
};

System::SharedPtr<::ApiExamples::ExStructuredDocumentTag> ExStructuredDocumentTag::s_instance;

TEST_F(ExStructuredDocumentTag, RepeatingSection)
{
    s_instance->RepeatingSection();
}

TEST_F(ExStructuredDocumentTag, ApplyStyle)
{
    s_instance->ApplyStyle();
}

TEST_F(ExStructuredDocumentTag, CheckBox)
{
    s_instance->CheckBox();
}

TEST_F(ExStructuredDocumentTag, SkipMono_Date)
{
    RecordProperty("category", "SkipMono");

    s_instance->Date();
}

TEST_F(ExStructuredDocumentTag, PlainText)
{
    s_instance->PlainText();
}

using ExStructuredDocumentTag_IsTemporary_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExStructuredDocumentTag::IsTemporary)>::type;

struct ExStructuredDocumentTag_IsTemporary : public ExStructuredDocumentTag,
                                             public ApiExamples::ExStructuredDocumentTag,
                                             public ::testing::WithParamInterface<ExStructuredDocumentTag_IsTemporary_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExStructuredDocumentTag_IsTemporary, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->IsTemporary(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExStructuredDocumentTag_IsTemporary, ::testing::ValuesIn(ExStructuredDocumentTag_IsTemporary::TestCases()));

using ExStructuredDocumentTag_PlaceholderBuildingBlock_Args =
    System::MethodArgumentTuple<decltype(&ApiExamples::ExStructuredDocumentTag::PlaceholderBuildingBlock)>::type;

struct ExStructuredDocumentTag_PlaceholderBuildingBlock : public ExStructuredDocumentTag,
                                                          public ApiExamples::ExStructuredDocumentTag,
                                                          public ::testing::WithParamInterface<ExStructuredDocumentTag_PlaceholderBuildingBlock_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExStructuredDocumentTag_PlaceholderBuildingBlock, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->PlaceholderBuildingBlock(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExStructuredDocumentTag_PlaceholderBuildingBlock,
                         ::testing::ValuesIn(ExStructuredDocumentTag_PlaceholderBuildingBlock::TestCases()));

TEST_F(ExStructuredDocumentTag, Lock)
{
    s_instance->Lock();
}

TEST_F(ExStructuredDocumentTag, ListItemCollection)
{
    s_instance->ListItemCollection();
}

TEST_F(ExStructuredDocumentTag, CreatingCustomXml)
{
    s_instance->CreatingCustomXml();
}

TEST_F(ExStructuredDocumentTag, XmlMapping)
{
    s_instance->XmlMapping_();
}

TEST_F(ExStructuredDocumentTag, StructuredDocumentTagRangeStartXmlMapping)
{
    s_instance->StructuredDocumentTagRangeStartXmlMapping();
}

TEST_F(ExStructuredDocumentTag, CustomXmlSchemaCollection)
{
    s_instance->CustomXmlSchemaCollection_();
}

TEST_F(ExStructuredDocumentTag, CustomXmlPartStoreItemIdReadOnly)
{
    s_instance->CustomXmlPartStoreItemIdReadOnly();
}

TEST_F(ExStructuredDocumentTag, CustomXmlPartStoreItemIdReadOnlyNull)
{
    s_instance->CustomXmlPartStoreItemIdReadOnlyNull();
}

TEST_F(ExStructuredDocumentTag, ClearTextFromStructuredDocumentTags)
{
    s_instance->ClearTextFromStructuredDocumentTags();
}

TEST_F(ExStructuredDocumentTag, AccessToBuildingBlockPropertiesFromDocPartObjSdt)
{
    s_instance->AccessToBuildingBlockPropertiesFromDocPartObjSdt();
}

TEST_F(ExStructuredDocumentTag, AccessToBuildingBlockPropertiesFromPlainTextSdt)
{
    s_instance->AccessToBuildingBlockPropertiesFromPlainTextSdt();
}

TEST_F(ExStructuredDocumentTag, BuildingBlockCategories)
{
    s_instance->BuildingBlockCategories();
}

using ExStructuredDocumentTag_UpdateSdtContent_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExStructuredDocumentTag::UpdateSdtContent)>::type;

struct ExStructuredDocumentTag_UpdateSdtContent : public ExStructuredDocumentTag,
                                                  public ApiExamples::ExStructuredDocumentTag,
                                                  public ::testing::WithParamInterface<ExStructuredDocumentTag_UpdateSdtContent_Args>
{
    static std::vector<ParamType> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(ExStructuredDocumentTag_UpdateSdtContent, Test)
{
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateSdtContent(std::get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(, ExStructuredDocumentTag_UpdateSdtContent, ::testing::ValuesIn(ExStructuredDocumentTag_UpdateSdtContent::TestCases()));

TEST_F(ExStructuredDocumentTag, FillTableUsingRepeatingSectionItem)
{
    s_instance->FillTableUsingRepeatingSectionItem();
}

TEST_F(ExStructuredDocumentTag, CustomXmlPart)
{
    s_instance->CustomXmlPart_();
}

TEST_F(ExStructuredDocumentTag, MultiSectionTags)
{
    s_instance->MultiSectionTags();
}

TEST_F(ExStructuredDocumentTag, SdtChildNodes)
{
    s_instance->SdtChildNodes();
}

}} // namespace ApiExamples::gtest_test
