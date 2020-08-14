// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
#include "ExMailMerge.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/enum_helpers.h>
#include <system/collections/ilist.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMergeRegionInfo.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMergeCleanupOptions.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeCallback.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldMergeField.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldGreetingLine.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/MailMerge/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words::Fields;
using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2889428649u, ::ApiExamples::ExMailMerge::MailMergeCallbackStub, ThisTypeBaseTypesInfo);

int ExMailMerge::MailMergeCallbackStub::get_TagsReplacedCounter()
{
    return pr_TagsReplacedCounter;
}

void ExMailMerge::MailMergeCallbackStub::set_TagsReplacedCounter(int value)
{
    pr_TagsReplacedCounter = value;
}

void ExMailMerge::MailMergeCallbackStub::TagsReplaced()
{
    set_TagsReplacedCounter(get_TagsReplacedCounter() + 1);
}

ExMailMerge::MailMergeCallbackStub::MailMergeCallbackStub() : pr_TagsReplacedCounter(0)
{
}

namespace gtest_test
{

class ExMailMerge : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExMailMerge> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

public:
    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExMailMerge>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExMailMerge> ExMailMerge::s_instance;

} // namespace gtest_test

void ExMailMerge::MailMergeRegionInfo()
{
    //ExStart
    //ExFor:MailMerge.GetFieldNamesForRegion(System.String)
    //ExFor:MailMerge.GetFieldNamesForRegion(System.String,System.Int32)
    //ExFor:MailMerge.GetRegionsByName(System.String)
    //ExFor:MailMerge.RegionEndTag
    //ExFor:MailMerge.RegionStartTag
    //ExSummary:Shows how to create, list and read mail merge regions.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    // These tags, which go inside MERGEFIELDs, denote the strings that signify the starts and ends of mail merge regions
    ASSERT_EQ(u"TableStart", doc->get_MailMerge()->get_RegionStartTag());
    ASSERT_EQ(u"TableEnd", doc->get_MailMerge()->get_RegionEndTag());

    // By using these tags, we will start and end a "MailMergeRegion1", which will contain MERGEFIELDs for two columns
    builder->InsertField(u" MERGEFIELD TableStart:MailMergeRegion1");
    builder->InsertField(u" MERGEFIELD Column1");
    builder->Write(u", ");
    builder->InsertField(u" MERGEFIELD Column2");
    builder->InsertField(u" MERGEFIELD TableEnd:MailMergeRegion1");

    // We can keep track of merge regions and their columns by looking at these collections
    SharedPtr<System::Collections::Generic::IList<SharedPtr<Aspose::Words::MailMerging::MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"MailMergeRegion1");
    ASSERT_EQ(1, regions->get_Count());
    ASSERT_EQ(u"MailMergeRegion1", regions->idx_get(0)->get_Name());

    ArrayPtr<String> mergeFieldNames = doc->get_MailMerge()->GetFieldNamesForRegion(u"MailMergeRegion1");
    ASSERT_EQ(u"Column1", mergeFieldNames[0]);
    ASSERT_EQ(u"Column2", mergeFieldNames[1]);

    // Insert a region with the same name as an existing region, which will make it a duplicate
    builder->InsertParagraph();
    // A single row/paragraph cannot be shared by multiple regions
    builder->InsertField(u" MERGEFIELD TableStart:MailMergeRegion1");
    builder->InsertField(u" MERGEFIELD Column3");
    builder->InsertField(u" MERGEFIELD TableEnd:MailMergeRegion1");

    // Regions that share the same name are still accounted for and can be accessed by index
    regions = doc->get_MailMerge()->GetRegionsByName(u"MailMergeRegion1");
    ASSERT_EQ(2, regions->get_Count());

    mergeFieldNames = doc->get_MailMerge()->GetFieldNamesForRegion(u"MailMergeRegion1", 1);
    ASSERT_EQ(u"Column3", mergeFieldNames[0]);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, MailMergeRegionInfo)
{
    s_instance->MailMergeRegionInfo();
}

} // namespace gtest_test

void ExMailMerge::TrimWhiteSpaces(bool doTrimWhitespaces)
{
    //ExStart
    //ExFor:MailMerge.TrimWhitespaces
    //ExSummary:Shows how to trimmed whitespaces from mail merge values.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->InsertField(u"MERGEFIELD myMergeField", nullptr);

    doc->get_MailMerge()->set_TrimWhitespaces(doTrimWhitespaces);
    doc->get_MailMerge()->Execute(MakeArray<String>({u"myMergeField"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"\t hello world! ")}));

    if (doTrimWhitespaces)
    {
        ASSERT_EQ(u"hello world!\f", doc->GetText());
    }
    else
    {
        ASSERT_EQ(u"\t hello world! \f", doc->GetText());
    }
    //ExEnd
}

namespace gtest_test
{

using TrimWhiteSpaces_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::TrimWhiteSpaces)>::type;

struct TrimWhiteSpacesVP : public ExMailMerge, public ApiExamples::ExMailMerge, public ::testing::WithParamInterface<TrimWhiteSpaces_Args>
{
    static std::vector<TrimWhiteSpaces_Args> TestCases()
    {
        return
        {
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

} // namespace gtest_test

void ExMailMerge::MailMergeGetFieldNames()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.GetFieldNames
    //ExSummary:Shows how to get names of all merge fields in a document.
    ArrayPtr<String> fieldNames = doc->get_MailMerge()->GetFieldNames();
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, MailMergeGetFieldNames)
{
    s_instance->MailMergeGetFieldNames();
}

} // namespace gtest_test

void ExMailMerge::DeleteFields()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.DeleteFields
    //ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
    doc->get_MailMerge()->DeleteFields();
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, DeleteFields)
{
    s_instance->DeleteFields();
}

} // namespace gtest_test

void ExMailMerge::RemoveContainingFields()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.CleanupOptions
    //ExFor:MailMergeCleanupOptions
    //ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
    doc->get_MailMerge()->set_CleanupOptions(Aspose::Words::MailMerging::MailMergeCleanupOptions::RemoveContainingFields);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, RemoveContainingFields)
{
    s_instance->RemoveContainingFields();
}

} // namespace gtest_test

void ExMailMerge::RemoveUnusedFields()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.CleanupOptions
    //ExFor:MailMergeCleanupOptions
    //ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
    doc->get_MailMerge()->set_CleanupOptions(Aspose::Words::MailMerging::MailMergeCleanupOptions::RemoveUnusedFields);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, RemoveUnusedFields)
{
    s_instance->RemoveUnusedFields();
}

} // namespace gtest_test

void ExMailMerge::RemoveEmptyParagraphs()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.CleanupOptions
    //ExFor:MailMergeCleanupOptions
    //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
    doc->get_MailMerge()->set_CleanupOptions(Aspose::Words::MailMerging::MailMergeCleanupOptions::RemoveEmptyParagraphs);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, RemoveEmptyParagraphs)
{
    s_instance->RemoveEmptyParagraphs();
}

} // namespace gtest_test

void ExMailMerge::RemoveColonBetweenEmptyMergeFields(String punctuationMark, bool isCleanupParagraphsWithPunctuationMarks, String resultText)
{
    //ExStart
    //ExFor:MailMerge.CleanupParagraphsWithPunctuationMarks
    //ExSummary:Shows how to remove paragraphs with punctuation marks after mail merge operation.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    auto mergeFieldOption1 = System::DynamicCast<Aspose::Words::Fields::FieldMergeField>(builder->InsertField(u"MERGEFIELD", u"Option_1"));
    mergeFieldOption1->set_FieldName(u"Option_1");

    // Here is the complete list of cleanable punctuation marks: ! , . : ; ? ¡ ¿
    builder->Write(punctuationMark);

    auto mergeFieldOption2 = System::DynamicCast<Aspose::Words::Fields::FieldMergeField>(builder->InsertField(u"MERGEFIELD", u"Option_2"));
    mergeFieldOption2->set_FieldName(u"Option_2");

    doc->get_MailMerge()->set_CleanupOptions(Aspose::Words::MailMerging::MailMergeCleanupOptions::RemoveEmptyParagraphs);
    // The default value of the option is true which means that the behaviour was changed to mimic MS Word
    // If you rely on the old behavior are able to revert it by setting the option to false
    doc->get_MailMerge()->set_CleanupParagraphsWithPunctuationMarks(isCleanupParagraphsWithPunctuationMarks);

    doc->get_MailMerge()->Execute(MakeArray<String>({u"Option_1", u"Option_2"}), MakeArray<SharedPtr<System::Object>>({nullptr, nullptr}));

    doc->Save(ArtifactsDir + u"MailMerge.RemoveColonBetweenEmptyMergeFields.docx");
    //ExEnd

    ASSERT_EQ(resultText, doc->GetText());
}

namespace gtest_test
{

using RemoveColonBetweenEmptyMergeFields_Args = System::MethodArgumentTuple<decltype(&ApiExamples::ExMailMerge::RemoveColonBetweenEmptyMergeFields)>::type;

struct RemoveColonBetweenEmptyMergeFieldsVP : public ExMailMerge, public ApiExamples::ExMailMerge, public ::testing::WithParamInterface<RemoveColonBetweenEmptyMergeFields_Args>
{
    static std::vector<RemoveColonBetweenEmptyMergeFields_Args> TestCases()
    {
        return
        {
            std::make_tuple(u"!", false, u""),
            std::make_tuple(u", ", false, u""),
            std::make_tuple(u" . ", false, u""),
            std::make_tuple(u" :", false, u""),
            std::make_tuple(u"  ; ", false, u""),
            std::make_tuple(u" ?  ", false, u""),
            std::make_tuple(u"  ¡  ", false, u""),
            std::make_tuple(u"  ¿  ", false, u""),
            std::make_tuple(u"!", true, u"!\f"),
            std::make_tuple(u", ", true, u", \f"),
            std::make_tuple(u" . ", true, u" . \f"),
            std::make_tuple(u" :", true, u" :\f"),
            std::make_tuple(u"  ; ", true, u"  ; \f"),
            std::make_tuple(u" ?  ", true, u" ?  \f"),
            std::make_tuple(u"  ¡  ", true, u"  ¡  \f"),
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

} // namespace gtest_test

void ExMailMerge::GetFieldNames()
{
    //ExStart
    //ExFor:FieldAddressBlock
    //ExFor:FieldAddressBlock.GetFieldNames
    //ExSummary:Shows how to get mail merge field names used by the field.
    auto doc = MakeObject<Document>(MyDir + u"Field sample - ADDRESSBLOCK.docx");

    ArrayPtr<String> addressFieldsExpect = MakeArray<String>({u"Company", u"First Name", u"Middle Name", u"Last Name", u"Suffix", u"Address 1", u"City", u"State", u"Country or Region", u"Postal Code"});

    auto addressBlockField = System::DynamicCast<Aspose::Words::Fields::FieldAddressBlock>(doc->get_Range()->get_Fields()->idx_get(0));
    ArrayPtr<String> addressBlockFieldNames = addressBlockField->GetFieldNames();
    //ExEnd

    ASPOSE_ASSERT_EQ(addressFieldsExpect, addressBlockFieldNames);

    ArrayPtr<String> greetingFieldsExpect = MakeArray<String>({u"Courtesy Title", u"Last Name"});

    auto greetingLineField = System::DynamicCast<Aspose::Words::Fields::FieldGreetingLine>(doc->get_Range()->get_Fields()->idx_get(1));
    ArrayPtr<String> greetingLineFieldNames = greetingLineField->GetFieldNames();

    ASPOSE_ASSERT_EQ(greetingFieldsExpect, greetingLineFieldNames);
}

namespace gtest_test
{

TEST_F(ExMailMerge, GetFieldNames)
{
    s_instance->GetFieldNames();
}

} // namespace gtest_test

void ExMailMerge::UseNonMergeFields()
{
    auto doc = MakeObject<Document>();
    //ExStart
    //ExFor:MailMerge.UseNonMergeFields
    //ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
    doc->get_MailMerge()->set_UseNonMergeFields(true);
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, UseNonMergeFields)
{
    s_instance->UseNonMergeFields();
}

} // namespace gtest_test

void ExMailMerge::TestMailMergeGetRegionsHierarchy()
{
    //ExStart
    //ExFor:MailMerge.GetRegionsHierarchy
    //ExFor:MailMergeRegionInfo
    //ExFor:MailMergeRegionInfo.Regions
    //ExFor:MailMergeRegionInfo.Name
    //ExFor:MailMergeRegionInfo.Fields
    //ExFor:MailMergeRegionInfo.StartField
    //ExFor:MailMergeRegionInfo.EndField
    //ExFor:MailMergeRegionInfo.Level
    //ExSummary:Shows how to get MailMergeRegionInfo and work with it.
    auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

    // Returns a full hierarchy of regions (with fields) available in the document
    SharedPtr<Aspose::Words::MailMerging::MailMergeRegionInfo> regionInfo = doc->get_MailMerge()->GetRegionsHierarchy();

    // Get top regions in the document
    SharedPtr<System::Collections::Generic::IList<SharedPtr<Aspose::Words::MailMerging::MailMergeRegionInfo>>> topRegions = regionInfo->get_Regions();
    ASSERT_EQ(2, topRegions->get_Count());
    ASSERT_EQ(u"Region1", topRegions->idx_get(0)->get_Name());
    ASSERT_EQ(u"Region2", topRegions->idx_get(1)->get_Name());
    ASSERT_EQ(1, topRegions->idx_get(0)->get_Level());
    ASSERT_EQ(1, topRegions->idx_get(1)->get_Level());

    // Get nested region in first top region
    SharedPtr<System::Collections::Generic::IList<SharedPtr<Aspose::Words::MailMerging::MailMergeRegionInfo>>> nestedRegions = topRegions->idx_get(0)->get_Regions();
    ASSERT_EQ(2, nestedRegions->get_Count());
    ASSERT_EQ(u"NestedRegion1", nestedRegions->idx_get(0)->get_Name());
    ASSERT_EQ(u"NestedRegion2", nestedRegions->idx_get(1)->get_Name());
    ASSERT_EQ(2, nestedRegions->idx_get(0)->get_Level());
    ASSERT_EQ(2, nestedRegions->idx_get(1)->get_Level());

    // Get field list in first top region
    SharedPtr<System::Collections::Generic::IList<SharedPtr<Field>>> fieldList = topRegions->idx_get(0)->get_Fields();
    ASSERT_EQ(4, fieldList->get_Count());

    SharedPtr<FieldMergeField> startFieldMergeField = nestedRegions->idx_get(0)->get_StartField();
    ASSERT_EQ(u"TableStart:NestedRegion1", startFieldMergeField->get_FieldName());

    SharedPtr<FieldMergeField> endFieldMergeField = nestedRegions->idx_get(0)->get_EndField();
    ASSERT_EQ(u"TableEnd:NestedRegion1", endFieldMergeField->get_FieldName());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExMailMerge, TestMailMergeGetRegionsHierarchy)
{
    s_instance->TestMailMergeGetRegionsHierarchy();
}

} // namespace gtest_test

void ExMailMerge::TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption()
{
    //ExStart
    //ExFor:MailMerge.MailMergeCallback
    //ExFor:IMailMergeCallback
    //ExFor:IMailMergeCallback.TagsReplaced
    //ExSummary:Shows how to define custom logic for handling events during mail merge.
    auto document = MakeObject<Document>();
    document->get_MailMerge()->set_UseNonMergeFields(true);

    auto mailMergeCallbackStub = MakeObject<ExMailMerge::MailMergeCallbackStub>();
    document->get_MailMerge()->set_MailMergeCallback(mailMergeCallbackStub);

    document->get_MailMerge()->Execute(MakeArray<String>(0), MakeArray<SharedPtr<System::Object>>(0));

    ASSERT_EQ(1, mailMergeCallbackStub->get_TagsReplacedCounter());
}

namespace gtest_test
{

TEST_F(ExMailMerge, TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption)
{
    s_instance->TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption();
}

} // namespace gtest_test

void ExMailMerge::GetRegionsByName()
{
    auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

    SharedPtr<System::Collections::Generic::IList<SharedPtr<Aspose::Words::MailMerging::MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"Region1");
    ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region1")->get_Count());
    for (auto region : System::IterateOver(regions))
    {
        ASSERT_EQ(u"Region1", region->get_Name());
    }

    regions = doc->get_MailMerge()->GetRegionsByName(u"Region2");
    ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region2")->get_Count());
    for (auto region : System::IterateOver(regions))
    {
        ASSERT_EQ(u"Region2", region->get_Name());
    }

    regions = doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1");
    ASSERT_EQ(2, doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1")->get_Count());
    for (auto region : System::IterateOver(regions))
    {
        ASSERT_EQ(u"NestedRegion1", region->get_Name());
    }
}

namespace gtest_test
{

TEST_F(ExMailMerge, GetRegionsByName)
{
    s_instance->GetRegionsByName();
}

} // namespace gtest_test

} // namespace ApiExamples
