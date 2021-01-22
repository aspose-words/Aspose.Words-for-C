#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Range.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceOptions.h>
#include <Aspose.Words.Cpp/Model/FindReplace/FindReplaceDirection.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <drawing/color.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Replacing;

namespace 
{
    void Replace(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // ExStart:ReplaceWithString
        // Load a Word Docx document by creating an instance of the Document class.
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello _CustomerName_, ");

        // Specify the search string and replace string using the Replace method.
        doc->get_Range()->Replace(u"_CustomerName_", u"James Bond", System::MakeObject<FindReplaceOptions>());

        // Save the result.
        System::String outputPath = outputDataDir + u"Range.ReplaceSimple.docx";
        doc->Save(outputPath);
        // ExEnd:ReplaceWithString
    }

    void HighlightColor(System::String const& inputDataDir, System::String const& outputDataDir)
    {
        // Load a Word Docx document by creating an instance of the Document class.
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello _CustomerName_,");

        // ExStart:HighlightColor
        // Highlight word "the" with yellow color.
        auto options = System::MakeObject<FindReplaceOptions>();
        options->get_ApplyFont()->set_HighlightColor(System::Drawing::Color::get_Yellow());

        // Replace highlighted text.
        doc->get_Range()->Replace(u"Hello", u"Hello", options);
        // ExEnd:HighlightColor
    }

}

void ReplaceWithString()
{
    std::cout << "ReplaceWithString example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_FindAndReplace();
    System::String outputDataDir = GetOutputDataDir_FindAndReplace();

    Replace(inputDataDir, outputDataDir);
    HighlightColor(inputDataDir, outputDataDir);

    std::cout << "ReplaceWithString example finished." << std::endl << std::endl;
}
