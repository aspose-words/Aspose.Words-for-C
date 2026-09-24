#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/MailMerging/MailMergeRegionInfo.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>
#include <SQLiteCpp.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting {

class BaseOperations : public DocsExamplesBase
{
public:
    void SimpleMailMerge()
    {
        //ExStart:ExecuteSimpleMailMerge
        //GistId:533914df569fd0a1003bec94eff1d505
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        
        builder->InsertField(u" MERGEFIELD CustomerName ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Item ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Quantity ");

        auto boxHelper = [](const char16_t* value) { return System::ObjectExt::Box<String>(value); };

        // Fill the fields in the document with user data.
        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"CustomerName", u"Item", u"Quantity"}),
            MakeArray<SharedPtr<System::Object>>(
                { boxHelper(u"John Doe"), boxHelper(u"Hawaiian"), boxHelper(u"2")}));

        doc->Save(ArtifactsDir + u"BaseOperations.SimpleMailMerge.docx");
        //ExEnd:ExecuteSimpleMailMerge
    }

    void UseIfElseMustache()
    {
        //ExStart:UseIfElseMustache
        //GistId:c1bbd76f39074502f828f4ae24ee3681
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destinations - Mustache syntax.docx");

        doc->get_MailMerge()->set_UseNonMergeFields(true);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"GENDER"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"MALE")}));

        doc->Save(ArtifactsDir + u"BaseOperations.IfElseMustache.docx");
        //ExEnd:UseIfElseMustache
    }

    void GetRegionsByName()
    {
        //ExStart:GetRegionsByName
        //GistId:5c23f9f9a360347715de2a3123b0604f
        auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

        //ExStart:GetRegionsHierarchy
        //GistId:5c23f9f9a360347715de2a3123b0604f
        SharedPtr<MailMergeRegionInfo> regionInfo = doc->get_MailMerge()->GetRegionsHierarchy();
        //ExEnd:GetRegionsHierarchy

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
        //ExEnd:GetRegionsByName
    }

};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
