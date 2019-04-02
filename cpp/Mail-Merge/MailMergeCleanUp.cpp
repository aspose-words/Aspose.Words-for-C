#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/MailMerge/MailMerge.h>
#include <Model/MailMerge/MailMergeCleanupOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace
{
    void CleanupParagraphsWithPunctuationMarks(System::String const &dataDir)
    {
        // ExStart:CleanupParagraphsWithPunctuationMarks
        System::String fileName = u"MailMerge.CleanupPunctuationMarks.docx";
        // Open the document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + fileName);

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyParagraphs);
        doc->get_MailMerge()->set_CleanupParagraphsWithPunctuationMarks(false);

        typedef System::SharedPtr<System::Object> TObjectPtr;
        doc->get_MailMerge()->Execute(System::MakeArray<System::String>({u"field1", u"field2"}),
                                      System::MakeArray<TObjectPtr>({System::ObjectExt::Box<System::String>(u""), System::ObjectExt::Box<System::String>(u"")}));

        System::String outputPath = dataDir + GetOutputFilePath(u"MailMergeCleanUp.CleanupParagraphsWithPunctuationMarks.docx");
        // Save the output document to disk.
        doc->Save(outputPath);
        // ExEnd:CleanupParagraphsWithPunctuationMarks
        std::cout << "Mail merge performed with cleanup paragraphs having punctuation marks successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void MailMergeCleanUp()
{
    std::cout << "MailMergeCleanUp example started." << std::endl;
    // The path to the documents directory.
    System::String dataDir = GetDataDir_MailMergeAndReporting();
    CleanupParagraphsWithPunctuationMarks(dataDir);
    std::cout << "MailMergeCleanUp example finished." << std::endl << std::endl;
}