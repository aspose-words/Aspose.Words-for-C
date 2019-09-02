#include "stdafx.h"
#include "examples.h"

#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/SaveFormat.h>
#include <Model/Markup/MarkupLevel.h>
#include <Model/Markup/Sdt/SdtType.h>
#include <Model/Markup/Sdt/StructuredDocumentTag.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Markup;

void CheckBoxTypeContentControl()
{
    std::cout << "CheckBoxTypeContentControl example started." << std::endl;
    // ExStart:CheckBoxTypeContentControl
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    // Open the empty document
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    System::SharedPtr<StructuredDocumentTag> SdtCheckBox = System::MakeObject<StructuredDocumentTag>(doc, SdtType::Checkbox, MarkupLevel::Inline);

    // Insert content control into the document
    builder->InsertNode(SdtCheckBox);
    System::String outputPath = outputDataDir + u"CheckBoxTypeContentControl.docx";

    doc->Save(outputPath, SaveFormat::Docx);
    // ExEnd:CheckBoxTypeContentControl
    std::cout << "CheckBox type content control created successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "CheckBoxTypeContentControl example finished." << std::endl << std::endl;
}