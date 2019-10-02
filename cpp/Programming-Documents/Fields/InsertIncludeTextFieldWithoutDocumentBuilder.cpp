#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Fields/Fields/LinksAndReferences/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Model/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertIncludeTextFieldWithoutDocumentBuilder()
{
    std::cout << "InsertIncludeTextFieldWithoutDocumentBuilder example started." << std::endl;
    // ExStart:InsertIncludeTextFieldWithoutDocumentBuilder
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithFields();
    System::String outputDataDir = GetOutputDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"in.doc");
    // Get paragraph you want to append this INCLUDETEXT field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert an INCLUDETEXT field like this:
    // { INCLUDETEXT  "file path" }

    // Create instance of FieldAsk class and lets build the above field code
    System::SharedPtr<FieldIncludeText> fieldIncludeText = System::DynamicCast<FieldIncludeText>(para->AppendField(FieldType::FieldIncludeText, false));
    fieldIncludeText->set_BookmarkName(u"bookmark");
    fieldIncludeText->set_SourceFullName(inputDataDir + u"IncludeText.docx");

    doc->get_FirstSection()->get_Body()->AppendChild(para);

    // Finally update this IncludeText field
    fieldIncludeText->Update();

    System::String outputPath = outputDataDir + u"InsertIncludeTextFieldWithoutDocumentBuilder.doc";
    doc->Save(outputPath);

    // ExEnd:InsertIncludeTextFieldWithoutDocumentBuilder
    std::cout << "IncludeText field without using document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertIncludeTextFieldWithoutDocumentBuilder example finished." << std::endl << std::endl;
}