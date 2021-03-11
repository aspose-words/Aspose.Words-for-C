#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumnCollection.h>

using namespace Aspose::Words;

void ModifyPageSetupInAllSectionsOfDocument()
{
    System::String inputDataDir = GetInputDataDir_WorkingWithSections();
    System::String outputDataDir = GetOutputDataDir_WorkingWithSections();
    //ExStart:ModifyPageSetupInAllSections
    auto doc = System::MakeObject<Document>();
    auto builder = System::MakeObject<DocumentBuilder>(doc);

    builder->Writeln(u"Hello1");
    doc->AppendChild(System::MakeObject<Section>(doc));
    builder->Writeln(u"Hello22");
    doc->AppendChild(System::MakeObject<Section>(doc));
    builder->Writeln(u"Hello3");
    doc->AppendChild(System::MakeObject<Section>(doc));
    builder->Writeln(u"Hello45");

    // It is important to understand that a document can contain many sections,
    // and each section has its page setup. In this case, we want to modify them all.
    for(auto section : System::IterateOver<Section>(doc))
    {
        section->get_PageSetup()->set_PaperSize(PaperSize::Letter);
    }

    doc->Save(outputDataDir + u"WorkingWithSection.ModifyPageSetupInAllSections.doc");
    //ExEnd:ModifyPageSetupInAllSections
}
