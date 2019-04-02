#include "stdafx.h"
#include "examples.h"

#include <Layout/Hyphenation/Hyphenation.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void LoadHyphenationDictionaryForLanguage()
{
    std::cout << "LoadHyphenationDictionaryForLanguage example started." << std::endl;
    // ExStart:LoadHyphenationDictionaryForLanguage
    // The path to the documents directory.
    System::String dataDir = GetDataDir_RenderingAndPrinting();

    // Load the documents which store the shapes we want to render.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"TestFile RenderShape.doc");
    System::SharedPtr<System::IO::Stream> stream = System::IO::File::OpenRead(dataDir + u"hyph_de_CH.dic");
    Hyphenation::RegisterDictionary(u"de-CH", stream);

    System::String outputPath = dataDir + GetOutputFilePath(u"LoadHyphenationDictionaryForLanguage.pdf");
    doc->Save(outputPath);
    // ExEnd:LoadHyphenationDictionaryForLanguage
    std::cout << "Hyphenation dictionary for special language loaded successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "LoadHyphenationDictionaryForLanguage example finished." << std::endl << std::endl;
}