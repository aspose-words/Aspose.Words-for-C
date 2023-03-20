#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldGreetingLine.h>
#include <Aspose.Words.Cpp/Fields/FieldMergeField.h>
#include <Aspose.Words.Cpp/HtmlInsertOptions.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/MailMerging/MailMergeCleanupOptions.h>
#include <Aspose.Words.Cpp/MailMerging/MailMergeRegionInfo.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/enum_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words::Fields;
using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExMailMerge : public ApiExampleBase
{

public:
    void MailMergeRegionInfo_()
    {
        //ExStart
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String)
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String,System.Int32)
        //ExFor:MailMerge.GetRegionsByName(System.String)
        //ExFor:MailMerge.RegionEndTag
        //ExFor:MailMerge.RegionStartTag
        //ExFor:MailMergeRegionInfo.ParentRegion
        //ExSummary:Shows how to create, list, and read mail merge regions.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // "TableStart" and "TableEnd" tags, which go inside MERGEFIELDs,
        // denote the strings that signify the starts and ends of mail merge regions.
        ASSERT_EQ(u"TableStart", doc->get_MailMerge()->get_RegionStartTag());
        ASSERT_EQ(u"TableEnd", doc->get_MailMerge()->get_RegionEndTag());

        // Use these tags to start and end a mail merge region named "MailMergeRegion1",
        // which will contain MERGEFIELDs for two columns.
        builder->InsertField(u" MERGEFIELD TableStart:MailMergeRegion1");
        builder->InsertField(u" MERGEFIELD Column1");
        builder->Write(u", ");
        builder->InsertField(u" MERGEFIELD Column2");
        builder->InsertField(u" MERGEFIELD TableEnd:MailMergeRegion1");

        // We can keep track of merge regions and their columns by looking at these collections.
        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"MailMergeRegion1");

        ASSERT_EQ(1, regions->get_Count());
        ASSERT_EQ(u"MailMergeRegion1", regions->idx_get(0)->get_Name());

        ArrayPtr<String> mergeFieldNames = doc->get_MailMerge()->GetFieldNamesForRegion(u"MailMergeRegion1");

        ASSERT_EQ(u"Column1", mergeFieldNames[0]);
        ASSERT_EQ(u"Column2", mergeFieldNames[1]);

        // Insert a region with the same name inside the existing region, which will make it a parent.
        // Now a "Column2" field will be inside a new region.
        builder->MoveToField(regions->idx_get(0)->get_Fields()->idx_get(1), false);
        builder->InsertField(u" MERGEFIELD TableStart:MailMergeRegion1");
        builder->MoveToField(regions->idx_get(0)->get_Fields()->idx_get(1), true);
        builder->InsertField(u" MERGEFIELD TableEnd:MailMergeRegion1");

        // If we look up the name of duplicate regions using the "GetRegionsByName" method,
        // it will return all such regions in a collection.
        regions = doc->get_MailMerge()->GetRegionsByName(u"MailMergeRegion1");

        ASSERT_EQ(2, regions->get_Count());
        // Check that the second region now has a parent region.
        ASSERT_EQ(u"MailMergeRegion1", regions->idx_get(1)->get_ParentRegion()->get_Name());

        mergeFieldNames = doc->get_MailMerge()->GetFieldNamesForRegion(u"MailMergeRegion1", 1);

        ASSERT_EQ(u"Column2", mergeFieldNames[0]);
        //ExEnd
    }

    void TrimWhiteSpaces(bool trimWhitespaces)
    {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trim whitespaces from values of a data source while executing a mail merge.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD myMergeField", nullptr);

        doc->get_MailMerge()->set_TrimWhitespaces(trimWhitespaces);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"myMergeField"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"\t hello world! ")}));

        ASSERT_EQ(trimWhitespaces ? String(u"hello world!\f") : String(u"\t hello world! \f"), doc->GetText());
        //ExEnd
    }

    void DeleteFields()
    {
        //ExStart
        //ExFor:MailMerge.DeleteFields
        //ExSummary:Shows how to delete all MERGEFIELDs from a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Dear ");
        builder->InsertField(u" MERGEFIELD FirstName ");
        builder->Write(u" ");
        builder->InsertField(u" MERGEFIELD LastName ");
        builder->Writeln(u",");
        builder->Writeln(u"Greetings!");

        ASSERT_EQ(u"Dear \u0013 MERGEFIELD FirstName \u0014«FirstName»\u0015 \u0013 MERGEFIELD LastName \u0014«LastName»\u0015,\rGreetings!",
                  doc->GetText().Trim());

        doc->get_MailMerge()->DeleteFields();

        ASSERT_EQ(u"Dear  ,\rGreetings!", doc->GetText().Trim());
        //ExEnd
    }

    void GetFieldNames()
    {
        //ExStart
        //ExFor:FieldAddressBlock
        //ExFor:FieldAddressBlock.GetFieldNames
        //ExSummary:Shows how to get mail merge field names used by a field.
        auto doc = MakeObject<Document>(MyDir + u"Field sample - ADDRESSBLOCK.docx");

        ArrayPtr<String> addressFieldsExpect = MakeArray<String>(
            {u"Company", u"First Name", u"Middle Name", u"Last Name", u"Suffix", u"Address 1", u"City", u"State", u"Country or Region", u"Postal Code"});

        auto addressBlockField = System::DynamicCast<FieldAddressBlock>(doc->get_Range()->get_Fields()->idx_get(0));
        ArrayPtr<String> addressBlockFieldNames = addressBlockField->GetFieldNames();
        //ExEnd

        ASPOSE_ASSERT_EQ(addressFieldsExpect, addressBlockFieldNames);

        ArrayPtr<String> greetingFieldsExpect = MakeArray<String>({u"Courtesy Title", u"Last Name"});

        auto greetingLineField = System::DynamicCast<FieldGreetingLine>(doc->get_Range()->get_Fields()->idx_get(1));
        ArrayPtr<String> greetingLineFieldNames = greetingLineField->GetFieldNames();

        ASPOSE_ASSERT_EQ(greetingFieldsExpect, greetingLineFieldNames);
    }

    void TestMailMergeGetRegionsHierarchy()
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
        //ExSummary:Shows how to verify mail merge regions.
        auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

        // Returns a full hierarchy of merge regions that contain MERGEFIELDs available in the document.
        SharedPtr<MailMergeRegionInfo> regionInfo = doc->get_MailMerge()->GetRegionsHierarchy();

        // Get top regions in the document.
        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> topRegions = regionInfo->get_Regions();

        ASSERT_EQ(2, topRegions->get_Count());
        ASSERT_EQ(u"Region1", topRegions->idx_get(0)->get_Name());
        ASSERT_EQ(u"Region2", topRegions->idx_get(1)->get_Name());
        ASSERT_EQ(1, topRegions->idx_get(0)->get_Level());
        ASSERT_EQ(1, topRegions->idx_get(1)->get_Level());

        // Get nested region in first top region.
        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> nestedRegions = topRegions->idx_get(0)->get_Regions();

        ASSERT_EQ(2, nestedRegions->get_Count());
        ASSERT_EQ(u"NestedRegion1", nestedRegions->idx_get(0)->get_Name());
        ASSERT_EQ(u"NestedRegion2", nestedRegions->idx_get(1)->get_Name());
        ASSERT_EQ(2, nestedRegions->idx_get(0)->get_Level());
        ASSERT_EQ(2, nestedRegions->idx_get(1)->get_Level());

        // Get list of fields inside the first top region.
        SharedPtr<System::Collections::Generic::IList<SharedPtr<Field>>> fieldList = topRegions->idx_get(0)->get_Fields();

        ASSERT_EQ(4, fieldList->get_Count());

        SharedPtr<FieldMergeField> startFieldMergeField = nestedRegions->idx_get(0)->get_StartField();

        ASSERT_EQ(u"TableStart:NestedRegion1", startFieldMergeField->get_FieldName());

        SharedPtr<FieldMergeField> endFieldMergeField = nestedRegions->idx_get(0)->get_EndField();

        ASSERT_EQ(u"TableEnd:NestedRegion1", endFieldMergeField->get_FieldName());
        //ExEnd
    }

    void GetRegionsByName()
    {
        auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"Region1");
        ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region1")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"Region1", region->get_Name());
        }

        regions = doc->get_MailMerge()->GetRegionsByName(u"Region2");
        ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region2")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"Region2", region->get_Name());
        }

        regions = doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1");
        ASSERT_EQ(2, doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"NestedRegion1", region->get_Name());
        }
    }

    void RestartListsAtEachSection()
    {
        //ExStart
        //ExFor:MailMerge.RestartListsAtEachSection
        //ExSummary:Shows how to control whether or not list numbering is restarted at each section when mail merge is performed.
        auto doc = MakeObject<Document>(MyDir + u"Section breaks with numbering.docx");

        doc->get_MailMerge()->set_RestartListsAtEachSection(false);
        doc->get_MailMerge()->Execute(MakeArray<String>(0), MakeArray<SharedPtr<System::Object>>(0));

        doc->Save(ArtifactsDir + u"MailMerge.RestartListsAtEachSection.pdf");
        //ExEnd
    }

    void RemoveLastEmptyParagraph()
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, HtmlInsertOptions)
        //ExSummary:Shows how to use options while inserting html.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" MERGEFIELD Name ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD EMAIL ");
        builder->InsertParagraph();

        // By default "DocumentBuilder.InsertHtml" inserts a HTML fragment that ends with a block-level HTML element,
        // it normally closes that block-level element and inserts a paragraph break.
        // As a result, a new empty paragraph appears after inserted document.
        // If we specify "HtmlInsertOptions.RemoveLastEmptyParagraph", those extra empty paragraphs will be removed.
        builder->MoveToMergeField(u"NAME");
        builder->InsertHtml(u"<p>John Smith</p>", HtmlInsertOptions::UseBuilderFormatting | HtmlInsertOptions::RemoveLastEmptyParagraph);
        builder->MoveToMergeField(u"EMAIL");
        builder->InsertHtml(u"<p>jsmith@example.com</p>", HtmlInsertOptions::UseBuilderFormatting);

        doc->Save(ArtifactsDir + u"MailMerge.RemoveLastEmptyParagraph.docx");
        //ExEnd

        ASSERT_EQ(4, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
    }
};

} // namespace ApiExamples
