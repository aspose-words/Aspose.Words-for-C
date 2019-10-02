#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Layout/Hyphenation/Hyphenation.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void HyphenateWordsOfLanguages()
{
    std::cout << "HyphenateWordsOfLanguages example started." << std::endl;
    // ExStart:HyphenateWordsOfLanguages
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_RenderingAndPrinting();
    System::String outputDataDir = GetOutputDataDir_RenderingAndPrinting();

    // Load the documents which store the shapes we want to render.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"TestFile RenderShape.doc");
    Hyphenation::RegisterDictionary(u"en-US", inputDataDir + u"hyph_en_US.dic");
    Hyphenation::RegisterDictionary(u"de-CH", inputDataDir + u"hyph_de_CH.dic");

    System::String outputPath = outputDataDir + u"HyphenateWordsOfLanguages.pdf";
    doc->Save(outputPath);
    // ExEnd:HyphenateWordsOfLanguages
    std::cout << "Words of special languages hyphenate successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "HyphenateWordsOfLanguages example finished." << std::endl << std::endl;
}