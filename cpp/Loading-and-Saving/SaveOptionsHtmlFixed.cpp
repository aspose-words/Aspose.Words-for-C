#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Saving/HtmlFixedSaveOptions.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace
{
    void UseFontFromTargetMachine(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:UseFontFromTargetMachine
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (doc).doc");

        System::SharedPtr<HtmlFixedSaveOptions> options = System::MakeObject<HtmlFixedSaveOptions>();
        options->set_UseTargetMachineFonts(true);

        System::String outputPath = outputDataDir + u"SaveOptionsHtmlFixed.UseFontFromTargetMachine.html";
        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:UseFontFromTargetMachine
        std::cout << "Fonts from target machine are used in saved HtmlFixed file." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void WriteAllCSSrulesinSingleFile(System::String const &inputDataDir, System::String const &outputDataDir)
    {
        // ExStart:WriteAllCSSrulesinSingleFile
        // Load the document from disk.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Test File (doc).doc");

        System::SharedPtr<HtmlFixedSaveOptions> options = System::MakeObject<HtmlFixedSaveOptions>();
        //Setting this property to true restores the old behavior (separate files) for compatibility with legacy code. 
        //Default value is false.
        //All CSS rules are written into single file "styles.css
        options->set_SaveFontFaceCssSeparately(false);

        System::String outputPath = outputDataDir + u"SaveOptionsHtmlFixed.WriteAllCSSrulesinSingleFile.html";
        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:WriteAllCSSrulesinSingleFile
        std::cout << "Write all CSS rules in single file successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }
}

void SaveOptionsHtmlFixed()
{
    std::cout << "SaveOptionsHtmlFixed example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_LoadingAndSaving();
    System::String outputDataDir = GetOutputDataDir_LoadingAndSaving();
    UseFontFromTargetMachine(inputDataDir, outputDataDir);
    WriteAllCSSrulesinSingleFile(inputDataDir, outputDataDir);
    std::cout << "SaveOptionsHtmlFixed example finished." << std::endl << std::endl;
}