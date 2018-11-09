#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <Model/Document/Document.h>
#include <Model/Fields/Fields/DocumentInformation/FieldAuthor.h>
#include <Model/Fields/FieldType.h>
#include <Model/Nodes/NodeCollection.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Text/Paragraph.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

void InsertAuthorField()
{
    std::cout << "InsertAuthorField example started." << std::endl;
    // ExStart:InsertAuthorField
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithFields();
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"in.doc");
    // Get paragraph you want to append this AUTHOR field to
    System::SharedPtr<Paragraph> para = System::DynamicCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(1));

    // We want to insert an AUTHOR field like this:
    // { AUTHOR Test1 }

    // Create instance of FieldAuthor class and lets build the above field code
    System::SharedPtr<FieldAuthor> field = System::DynamicCast<FieldAuthor>(para->AppendField(FieldType::FieldAuthor, false));

    // { AUTHOR Test1 }
    field->set_AuthorName(u"Test1");

    // Finally update this AUTHOR field
    field->Update();

    System::String outputPath = dataDir + GetOutputFilePath(u"InsertAuthorField.doc");
    doc->Save(outputPath);
    // ExEnd:InsertAuthorField
    std::cout << "Author field without document builder inserted successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "InsertAuthorField example finished." << std::endl << std::endl;
}