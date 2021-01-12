#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Saving/MarkdownSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOptions.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <system/io/memory_stream.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void SaveAsMD(const System::String& inputDataDir, const System::String& outputDataDir)
    {
        // ExStart:SaveAsMD
        // Load the document from disk.
        auto doc = System::MakeObject<Document>(inputDataDir + u"Test.docx");

        // Save the document to Markdown format.
        doc->Save(outputDataDir + u"SpecifyMarkdownSaveOptions.SaveAsMD.md");
        // ExEnd:SaveAsMD
    }

    void SpecifySaveOptionsAndSaveAsMD(const System::String& inputDataDir, const System::String& outputDataDir)
    {
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Some text!");

        // specify MarkDownSaveOptions
        auto saveOptions = SaveOptions::CreateSaveOptions(SaveFormat::Markdown);
        doc->Save(outputDataDir + u"SpecifyMarkdownSaveOptions.SpecifySaveOptionsAndSaveAsMD.md");
    }

    void SupportedMarkdownFeatures(const System::String& inputDataDir, const System::String& outputDataDir)
    {
        // ExStart:SupportedMarkdownFeatures
        auto doc = System::MakeObject<Document>();
        auto builder = System::MakeObject<DocumentBuilder>(doc);


        // Specify the "Heading 1" style for the paragraph.
        builder->InsertParagraph();
        builder->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        builder->Write(u"Heading 1");

        // Specify the Italic emphasis for the paragraph.
        builder->InsertParagraph();
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->get_ParagraphFormat()->set_StyleName(u"Normal");
        builder->get_Font()->set_Italic(true);
        builder->Write(u"Italic Text");
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder->set_Italic(false);

        // Specify a Hyperlink for the desired text.
        builder->InsertParagraph();
        builder->InsertHyperlink(u"Aspose", u"https://www.aspose.com", false);
        builder->Write(u"Aspose");

        // Save your document as a Markdown file.
        doc->Save(outputDataDir + u"SpecifyMarkdownSaveOptions.SupportedMarkdownFeatures.md");
        // ExEnd:SupportedMarkdownFeatures
    }

    void SetImagesFolder(const System::String& inputDataDir, const System::String& outputDataDir)
    {
        // ExStart:SetImagesFolder
        // Load the document from disk.
        auto doc = System::MakeObject<Document>(inputDataDir + u"Test.docx");

        auto so = System::MakeObject<MarkdownSaveOptions>();
        so->set_ImagesFolder(outputDataDir + u"\\Images");

        auto stream = System::MakeObject<System::IO::MemoryStream>();
        doc->Save(stream, so);
        // ExEnd:SetImagesFolder
    }
}

void SpecifyMarkdownSaveOptions()
{
    std::cout << "SpecifyMarkdownSaveOptions example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();

#if 0
    // Source document is missing
    SaveAsMD(inputDataDir, outputDataDir);
#endif 
    SpecifySaveOptionsAndSaveAsMD(inputDataDir, outputDataDir);
    SupportedMarkdownFeatures(inputDataDir, outputDataDir);
    SetImagesFolder(inputDataDir, outputDataDir);

    std::cout << "SpecifyMarkdownSaveOptions example finished." << std::endl << std::endl;
}
