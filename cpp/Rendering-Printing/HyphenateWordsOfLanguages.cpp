#include "stdafx.h"
#include "examples.h"

#include <Layout/Hyphenation/Hyphenation.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void HyphenateWordsOfLanguages()
{
    std::cout << "HyphenateWordsOfLanguages example started." << std::endl;
    // ExStart:HyphenateWordsOfLanguages
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();

    // Load the documents which store the shapes we want to render.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile RenderShape.doc");
    Hyphenation::RegisterDictionary(u"en-US", dataDir + u"hyph_en_US.dic");
    Hyphenation::RegisterDictionary(u"de-CH", dataDir + u"hyph_de_CH.dic");

    System::String outputPath = dataDir + GetOutputFilePath(u"HyphenateWordsOfLanguages.pdf");
    doc->Save(outputPath);
    // ExEnd:HyphenateWordsOfLanguages
    std::cout << "Words of special languages hyphenate successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "HyphenateWordsOfLanguages example finished." << std::endl << std::endl;
}