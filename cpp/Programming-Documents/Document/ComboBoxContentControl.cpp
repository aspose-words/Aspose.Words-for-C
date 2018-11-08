#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Markup/MarkupLevel.h>
#include <Model/Markup/Sdt/SdtListItem.h>
#include <Model/Markup/Sdt/SdtListItemCollection.h>
#include <Model/Markup/Sdt/SdtType.h>
#include <Model/Markup/Sdt/StructuredDocumentTag.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

void ComboBoxContentControl()
{
    std::cout << "ComboBoxContentControl example started." << std::endl;
    // ExStart:ComboBoxContentControl
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<StructuredDocumentTag> sdt = System::MakeObject<StructuredDocumentTag>(doc, SdtType::ComboBox, MarkupLevel::Block);

    sdt->get_ListItems()->Add(System::MakeObject<SdtListItem>(u"Choose an item", u"-1"));
    sdt->get_ListItems()->Add(System::MakeObject<SdtListItem>(u"Item 1", u"1"));
    sdt->get_ListItems()->Add(System::MakeObject<SdtListItem>(u"Item 2", u"2"));
    doc->get_FirstSection()->get_Body()->AppendChild(sdt);

    System::String outputPath = dataDir + GetOutputFilePath(u"ComboBoxContentControl.docx");
    doc->Save(outputPath);
    // ExEnd:ComboBoxContentControl
    std::cout << "Combo box type content control created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ComboBoxContentControl example finished." << std::endl << std::endl;
}