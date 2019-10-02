#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/TextDmlEffect.h>

using namespace Aspose::Words;

void CheckDMLTextEffect()
{
    std::cout << "CheckDMLTextEffect example started." << std::endl;
    // ExStart:CheckDMLTextEffect
    // The path to the documents directory.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();

    // Initialize document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"Document.doc");
    System::SharedPtr<RunCollection> runs = doc->get_FirstSection()->get_Body()->get_FirstParagraph()->get_Runs();
    System::SharedPtr<Font> runFont = runs->idx_get(0)->get_Font();

    // One run might have several Dml text effects applied.
    std::cout << runFont->HasDmlEffect(TextDmlEffect::Shadow) << std::endl;
    std::cout << runFont->HasDmlEffect(TextDmlEffect::Effect3D) << std::endl;
    std::cout << runFont->HasDmlEffect(TextDmlEffect::Reflection) << std::endl;
    std::cout << runFont->HasDmlEffect(TextDmlEffect::Outline) << std::endl;
    std::cout << runFont->HasDmlEffect(TextDmlEffect::Fill) << std::endl;
    // ExEnd:CheckDMLTextEffect
    std::cout << "CheckDMLTextEffect example finished." << std::endl << std::endl;
}