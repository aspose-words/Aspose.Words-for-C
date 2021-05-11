#pragma once

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Comparing/CompareOptions.h>
#include <Aspose.Words.Cpp/Comparing/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Comparing/Granularity.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/RevisionCollection.h>
#include <system/date_time.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Comparing;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class CompareDocument : public DocsExamplesBase
{
public:
    void CompareForEqual()
    {
        //ExStart:CompareForEqual
        auto docA = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Document> docB = docA->Clone();

        // DocA now contains changes as revisions.
        docA->Compare(docB, u"user", System::DateTime::get_Now());

        std::cout << (docA->get_Revisions()->get_Count() == 0 ? String(u"Documents are equal") : String(u"Documents are not equal")) << std::endl;
        //ExEnd:CompareForEqual
    }

    void CompareOptions_()
    {
        //ExStart:CompareOptions
        auto docA = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Document> docB = docA->Clone();

        auto options = MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        options->set_IgnoreHeadersAndFooters(true);
        options->set_IgnoreCaseChanges(true);
        options->set_IgnoreTables(true);
        options->set_IgnoreFields(true);
        options->set_IgnoreComments(true);
        options->set_IgnoreTextboxes(true);
        options->set_IgnoreFootnotes(true);

        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);

        std::cout << (docA->get_Revisions()->get_Count() == 0 ? String(u"Documents are equal") : String(u"Documents are not equal")) << std::endl;
        //ExEnd:CompareOptions
    }

    void ComparisonTarget()
    {
        //ExStart:ComparisonTarget
        auto docA = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Document> docB = docA->Clone();

        // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box.
        auto options = MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        options->set_Target(ComparisonTargetType::New);

        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);
        //ExEnd:ComparisonTarget
    }

    void ComparisonGranularity()
    {
        //ExStart:ComparisonGranularity
        auto builderA = MakeObject<DocumentBuilder>(MakeObject<Document>());
        auto builderB = MakeObject<DocumentBuilder>(MakeObject<Document>());

        builderA->Writeln(u"This is A simple word");
        builderB->Writeln(u"This is B simple words");

        auto compareOptions = MakeObject<CompareOptions>();
        compareOptions->set_Granularity(Granularity::CharLevel);

        builderA->get_Document()->Compare(builderB->get_Document(), u"author", System::DateTime::get_Now(), compareOptions);
        //ExEnd:ComparisonGranularity
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
