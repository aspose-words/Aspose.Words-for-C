#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Comparer/CompareOptions.h>
#include <Aspose.Words.Cpp/Model/Comparer/ComparisonTargetType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Revisions/RevisionCollection.h>

using namespace Aspose::Words;

namespace
{
    void NormalComparison(System::String const &inputDataDir)
    {
        // ExStart:NormalComparison
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");
        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now());
        // ExEnd:NormalComparison
    }

    void CompareForEqual(System::String const &inputDataDir)
    {
        // ExStart:CompareForEqual
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");
        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now());
        if (docA->get_Revisions()->get_Count() == 0)
        {
            std::cout << "Documents are equal" << std::endl;
        }
        else
        {
            std::cout << "Documents are not equal" << std::endl;
        }
        // ExEnd:CompareForEqual
    }

    void CompareDocumentWithCompareOptions(System::String const &inputDataDir)
    {
        // ExStart:CompareDocumentWithCompareOptions
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");

        System::SharedPtr<CompareOptions> options = System::MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        options->set_IgnoreHeadersAndFooters(true);
        options->set_IgnoreCaseChanges(true);
        options->set_IgnoreTables(true);
        options->set_IgnoreFields(true);
        options->set_IgnoreComments(true);
        options->set_IgnoreTextboxes(true);
        options->set_IgnoreFootnotes(true);

        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);
        if (docA->get_Revisions()->get_Count() == 0)
        {
            std::cout << "Documents are equal" << std::endl;
        }
        else
        {
            std::cout << "Documents are not equal" << std::endl;
        }
        // ExEnd:CompareDocumentWithCompareOptions
    }

    void CompareDocumentWithComparisonTarget(System::String const &inputDataDir)
    {
        // ExStart:CompareDocumentWithComparisonTarget
        System::SharedPtr<Document> docA = System::MakeObject<Document>(inputDataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(inputDataDir + u"TestFile - Copy.doc");

        System::SharedPtr<CompareOptions> options = System::MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box. 
        options->set_Target(ComparisonTargetType::New);

        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);
        // ExEnd:CompareDocumentWithComparisonTarget

        std::cout << "Documents have compared successfully." << std::endl;
    }

    void SpecifyComparisonGranularity(System::String const &inputDataDir)
    {
        // ExStart:SpecifyComparisonGranularity
        System::SharedPtr<DocumentBuilder> builderA = System::MakeObject<DocumentBuilder>(System::MakeObject<Document>());
        System::SharedPtr<DocumentBuilder> builderB = System::MakeObject<DocumentBuilder>(System::MakeObject<Document>());

        builderA->Writeln(u"This is A simple word");
        builderB->Writeln(u"This is B simple words");

        System::SharedPtr<CompareOptions> co = System::MakeObject<CompareOptions>();
        co->set_Granularity(Aspose::Words::Granularity::CharLevel);

        builderA->get_Document()->Compare(builderB->get_Document(), u"author", System::DateTime::get_Now(), co);
        // ExEnd:SpecifyComparisonGranularity
    }
}

void CompareDocument()
{
    std::cout << "CompareDocument example started." << std::endl;
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    NormalComparison(inputDataDir);
    CompareForEqual(inputDataDir);
    CompareDocumentWithCompareOptions(inputDataDir);
    CompareDocumentWithComparisonTarget(inputDataDir);
    SpecifyComparisonGranularity(inputDataDir);
    std::cout << "CompareDocument example finished." << std::endl << std::endl;
}