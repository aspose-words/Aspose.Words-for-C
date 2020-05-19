#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/CleanupOptions.h>

using namespace Aspose::Words;

namespace
{
    void CleanupUnusedStylesandLists(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		// ExStart:CleanupUnusedStylesandLists
	    auto doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

        auto cleanupoptions = System::MakeObject<CleanupOptions>();
        cleanupoptions->set_UnusedLists(false);
        cleanupoptions->set_UnusedStyles(true);

		// Cleans unused styles and lists from the document depending on given CleanupOptions. 
        doc->Cleanup(cleanupoptions);

        auto outputPath = outputDataDir + u"CleanupUnusedStylesandLists.docx";
		doc->Save(outputPath);
		// ExEnd:CleanupUnusedStylesandLists

        std::cout << "All revisions accepted.\nFile saved at " << outputPath.ToUtf8String() << '\n';
    }

    void CleanupDuplicateStyle(System::String const &inputDataDir, System::String const &outputDataDir)
    {
		// ExStart:CleanupDuplicateStyle
		auto doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");

		auto options  = System::MakeObject<CleanupOptions>();
        options->set_DuplicateStyle(true);

		// Cleans duplicate styles from the document. 
        doc->Cleanup(options);

        auto outputPath = outputDataDir + u"CleanupDuplicateStyle.docx";
		doc->Save(outputPath);
		// ExEnd:CleanupDuplicateStyle

        std::cout << "All revisions accepted.\nFile saved at " << outputPath.ToUtf8String() << '\n';
    }
}

void WorkWithCleanupOptions()
{
    std::cout << "WorkWithCleanupOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();

    CleanupUnusedStylesandLists(inputDataDir, outputDataDir);
    CleanupDuplicateStyle(inputDataDir, outputDataDir);

    std::cout << "WorkWithCleanupOptions  example finished." << std::endl << std::endl;
}
