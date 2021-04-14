#pragma once
// CPPDEFECT: System.Data is not supported

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMergeRegionInfo.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

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
        //ExStart:SimpleMailMerge
        // Include the code for our template.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create Merge Fields.
        builder->InsertField(u" MERGEFIELD CustomerName ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Item ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Quantity ");

        // Fill the fields in the document with user data.
        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"CustomerName", u"Item", u"Quantity"}),
            MakeArray<SharedPtr<System::Object>>(
                {System::ObjectExt::Box<String>(u"John Doe"), System::ObjectExt::Box<String>(u"Hawaiian"), System::ObjectExt::Box<String>(u"2")}));

        doc->Save(ArtifactsDir + u"BaseOperations.SimpleMailMerge.docx");
        //ExEnd:SimpleMailMerge
    }

    void UseIfElseMustache()
    {
        //ExStart:UseOfifelseMustacheSyntax
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destinations - Mustache syntax.docx");

        doc->get_MailMerge()->set_UseNonMergeFields(true);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"GENDER"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"MALE")}));

        doc->Save(ArtifactsDir + u"BaseOperations.IfElseMustache.docx");
        //ExEnd:UseOfifelseMustacheSyntax
    }

    void GetRegionsByName()
    {
        //ExStart:GetRegionsByName
        auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"Region1");
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
        //ExEnd:GetRegionsByName
    }
};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
