#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/special_casts.h>
#include <Model/Document/Document.h>
#include <Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/Section.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertIncludeTextFieldWithoutDocumentBuilder()
{
    std::cout << "InsertIncludeTextFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertIncludeTextFieldWithoutDocumentBuilder
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"in.doc");
    // Get paragraph you want to append this INCLUDETEXT field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert an INCLUDETEXT field like this:
    // { INCLUDETEXT  "file path" }

    // Create instance of FieldAsk class and lets build the above field code
    System::SharedPtr<FieldIncludeText> fieldIncludeText = System::DynamicCast<FieldIncludeText>(para->AppendField(FieldType::FieldIncludeText, false));
    fieldIncludeText->set_BookmarkName(u"bookmark");
    fieldIncludeText->set_SourceFullName(dataDir + u"IncludeText.docx");

    doc->get_FirstSection()->get_Body()->AppendChild(para);

    // Finally update this IncludeText field
    fieldIncludeText->Update();

    System::String outputPath = dataDir + GetOutputFilePath(u"InsertIncludeTextFieldWithoutDocumentBuilder.doc");
    doc->Save(outputPath);

    // ExEnd:InsertIncludeTextFieldWithoutDocumentBuilder
    std::cout << "IncludeText field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertIncludeTextFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}