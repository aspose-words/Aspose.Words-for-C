#include "stdafx.h"
#include "examples.h"

#include <Model/Comparer/CompareOptions.h>
#include <Model/Comparer/ComparisonTargetType.h>
#include <Model/Document/Document.h>
#include <Model/Revisions/RevisionCollection.h>

using namespace Aspose::Words;

namespace
{
    void NormalComparison(System::String const &dataDir)
    {
        // ExStart:NormalComparison
        System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");
        // DocA now contains changes as revisions. 
        docA->Compare(docB, u"user", System::DateTime::get_Now());
        // ExEnd:NormalComparison
    }

    void CompareForEqual(System::String const &dataDir)
    {
        // ExStart:CompareForEqual
        System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");
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

    void CompareDocumentWithCompareOptions(System::String const &dataDir)
    {
        // ExStart:CompareDocumentWithCompareOptions
        System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");

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

    void CompareDocumentWithComparisonTarget(System::String const &dataDir)
    {
        // ExStart:CompareDocumentWithComparisonTarget
        System::SharedPtr<Document> docA = System::MakeObject<Document>(dataDir + u"TestFile.doc");
        System::SharedPtr<Document> docB = System::MakeObject<Document>(dataDir + u"TestFile - Copy.doc");

        System::SharedPtr<CompareOptions> options = System::MakeObject<CompareOptions>();
        options->set_IgnoreFormatting(true);
        // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box. 
        options->set_Target(ComparisonTargetType::New);

        docA->Compare(docB, u"user", System::DateTime::get_Now(), options);
        // ExEnd:CompareDocumentWithComparisonTarget

        std::cout << "Documents have compared successfully." << std::endl;
    }
}

void CompareDocument()
{
    std::cout << "CompareDocument example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    NormalComparison(dataDir);
    CompareForEqual(dataDir);
    CompareDocumentWithCompareOptions(dataDir);
    CompareDocumentWithComparisonTarget(dataDir);
    std::cout << "CompareDocument example finished." << std::endl << std::endl;
}