#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/date_time.h>
#include <Model/Text/RunCollection.h>
#include <Model/Text/Run.h>
#include <Model/Text/ParagraphCollection.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/Comment.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>

using namespace Aspose::Words;

void AddComments()
{
    std::cout << "AddComments example started." << std::endl;
    // ExStart:AddComments
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithComments();
    // ExStart:CreateSimpleDocumentUsingDocumentBuilder
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->Write(u"Some text is added.");
    // ExEnd:CreateSimpleDocumentUsingDocumentBuilder
    System::SharedPtr<Comment> comment = System::MakeObject<Comment>(doc, u"Awais Hafeez", u"AH", System::DateTime::get_Today());
    builder->get_CurrentParagraph()->AppendChild(comment);
    comment->get_Paragraphs()->Add(System::MakeObject<Paragraph>(doc));
    comment->get_FirstParagraph()->get_Runs()->Add(System::MakeObject<Aspose::Words::Run>(doc, u"Comment text."));

    System::String outputPath = dataDir + GetOutputFilePath(u"AddComments.doc");
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:AddComments
    std::cout << "Comments added successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AddComments example finished." << std::endl << std::endl;
}
