#include "stdafx.h"
#include "examples.h"
#include "MailMergeCommon.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMergeCleanupOptions.h>


using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace
{

    void RemoveEmptyParagraphs(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		//ExStart:RemoveEmptyParagraphs
        auto doc = System::MakeObject<Document>(inputDataDir + u"MailMergeCleanUp.RemoveRowfromTable.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyParagraphs);

        doc->get_MailMerge()->Execute(
            System::MakeArray<System::String>({ u"FullName", u"Company", u"Address", u"Address2", u"City" }),
            BoxVector<System::String>({ u"James Bond", u"MI5 Headquarters", u"Milbank", u"", u"London" })
            );
		
        doc->Save(outputDataDir + u"MailMergeCleanUp.RemoveEmptyParagraphs.docx");
		//ExEnd:RemoveEmptyParagraphs
    }

	void RemoveUnusedFields(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		//ExStart:RemoveUnusedFields
        auto doc = System::MakeObject<Document>(inputDataDir + u"MailMergeCleanUp.RemoveRowfromTable.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveUnusedFields);

        doc->get_MailMerge()->Execute(
            System::MakeArray<System::String>({ u"FullName", u"Company", u"Address", u"Address2", u"City" }),
            BoxVector<System::String>({ u"James Bond", u"MI5 Headquarters", u"Milbank", u"", u"London" })
            );
		
        doc->Save(outputDataDir + u"MailMergeCleanUp.RemoveUnusedFields.docx");
		//ExEnd:RemoveUnusedFields
    }

	void RemoveContainingFields(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		//ExStart:RemoveContainingFields
        auto doc = System::MakeObject<Document>(inputDataDir + u"MailMergeCleanUp.RemoveRowfromTable.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveContainingFields);

        doc->get_MailMerge()->Execute(
            System::MakeArray<System::String>({ u"FullName", u"Company", u"Address", u"Address2", u"City" }),
            BoxVector<System::String>({ u"James Bond", u"MI5 Headquarters", u"Milbank", u"", u"London" })
            );
		
        doc->Save(outputDataDir + u"MailMergeCleanUp.RemoveContainingFields.docx");
		//ExEnd:RemoveContainingFields
    }

	void RemoveEmptyTableRows(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		//ExStart:RemoveEmptyTableRows
        auto doc = System::MakeObject<Document>(inputDataDir + u"MailMergeCleanUp.RemoveRowfromTable.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyTableRows);

        doc->get_MailMerge()->Execute(
            System::MakeArray<System::String>({ u"FullName", u"Company", u"Address", u"Address2", u"City" }),
            BoxVector<System::String>({ u"James Bond", u"MI5 Headquarters", u"Milbank", u"", u"London" })
            );
		
        doc->Save(outputDataDir + u"MailMergeCleanUp.RemoveEmptyTableRows.docx");
		//ExEnd:RemoveEmptyTableRows
    }

    void CleanupParagraphsWithPunctuationMarks(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:CleanupParagraphsWithPunctuationMarks
        // Open the document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"MailMerge.CleanupPunctuationMarks.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyParagraphs);
        doc->get_MailMerge()->set_CleanupParagraphsWithPunctuationMarks(false);

        doc->get_MailMerge()->Execute(
            System::MakeArray<System::String>({u"field1", u"field2"}),
            BoxVector<System::String>({u"", u""})
        );

        System::String outputPath = outputDataDir + u"MailMergeCleanUp.CleanupParagraphsWithPunctuationMarks.docx";
        // Save the output document to disk.
        doc->Save(outputPath);
        // ExEnd:CleanupParagraphsWithPunctuationMarks
        std::cout << "Mail merge performed with cleanup paragraphs having punctuation marks successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void MailMergeCleanUp()
{
    std::cout << "MailMergeCleanUp example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_MailMergeAndReporting();
    System::String outputDataDir = GetOutputDataDir_MailMergeAndReporting();

	RemoveEmptyParagraphs(inputDataDir, outputDataDir);
	RemoveUnusedFields(inputDataDir, outputDataDir);
	RemoveContainingFields(inputDataDir, outputDataDir);
	RemoveEmptyTableRows(inputDataDir, outputDataDir);

    CleanupParagraphsWithPunctuationMarks(inputDataDir, outputDataDir);
    std::cout << "MailMergeCleanUp example finished." << std::endl << std::endl;
}