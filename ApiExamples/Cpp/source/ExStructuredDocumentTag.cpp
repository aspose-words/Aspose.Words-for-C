// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: Aspose.Pdf is not supported
#include "ExStructuredDocumentTag.h"

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

TEST_F(ExStructuredDocumentTag, SetSpecificStyleToSdt)
{
    s_instance->SetSpecificStyleToSdt();
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

TEST_F(ExStructuredDocumentTag, IsTemporary)
{
    s_instance->IsTemporary();
}

TEST_F(ExStructuredDocumentTag, PlaceholderBuildingBlock)
{
    s_instance->PlaceholderBuildingBlock();
}

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

using UpdateSdtContent_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExStructuredDocumentTag::UpdateSdtContent)>::type;

struct UpdateSdtContentVP : public ExStructuredDocumentTag,
                            public ApiExamples::ExStructuredDocumentTag,
                            public ::testing::WithParamInterface<UpdateSdtContent_Args>
{
    static std::vector<UpdateSdtContent_Args> TestCases()
    {
        return {
            std::make_tuple(false),
            std::make_tuple(true),
        };
    }
};

TEST_P(UpdateSdtContentVP, Test)
{
    using std::get;
    const auto& params = GetParam();
    ASSERT_NO_FATAL_FAILURE(s_instance->UpdateSdtContent(get<0>(params)));
}

INSTANTIATE_TEST_SUITE_P(ExStructuredDocumentTag, UpdateSdtContentVP, ::testing::ValuesIn(UpdateSdtContentVP::TestCases()));

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

}} // namespace ApiExamples::gtest_test
