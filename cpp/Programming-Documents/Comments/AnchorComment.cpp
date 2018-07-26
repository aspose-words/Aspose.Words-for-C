#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/date_time.h>
#include <Model/Text/RunCollection.h>
#include <Model/Text/Run.h>
#include <Model/Text/ParagraphCollection.h>
#include <Model/Text/Paragraph.h>
#include <Model/Text/CommentRangeStart.h>
#include <Model/Text/CommentRangeEnd.h>
#include <Model/Text/Comment.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Document/Document.h>
#include <cstdint>

using namespace Aspose::Words;

void AnchorComment()
{
    // ExStart:AnchorComment
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithComments();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    
    System::SharedPtr<Paragraph> para1 = System::MakeObject<Paragraph>(doc);
    System::SharedPtr<Aspose::Words::Run> run1 = System::MakeObject<Aspose::Words::Run>(doc, u"Some ");
    System::SharedPtr<Aspose::Words::Run> run2 = System::MakeObject<Aspose::Words::Run>(doc, u"text ");
    para1->AppendChild(run1);
    para1->AppendChild(run2);
    doc->get_FirstSection()->get_Body()->AppendChild(para1);
    
    System::SharedPtr<Paragraph> para2 = System::MakeObject<Paragraph>(doc);
    System::SharedPtr<Aspose::Words::Run> run3 = System::MakeObject<Aspose::Words::Run>(doc, u"is ");
    System::SharedPtr<Aspose::Words::Run> run4 = System::MakeObject<Aspose::Words::Run>(doc, u"added ");
    para2->AppendChild(run3);
    para2->AppendChild(run4);
    doc->get_FirstSection()->get_Body()->AppendChild(para2);
    
    System::SharedPtr<Comment> comment = System::MakeObject<Comment>(doc, u"Awais Hafeez", u"AH", System::DateTime::get_Today());
    comment->get_Paragraphs()->Add(System::MakeObject<Paragraph>(doc));
    comment->get_FirstParagraph()->get_Runs()->Add(System::MakeObject<Aspose::Words::Run>(doc, u"Comment text."));
    
    System::SharedPtr<CommentRangeStart> commentRangeStart = System::MakeObject<CommentRangeStart>(doc, comment->get_Id());
    System::SharedPtr<CommentRangeEnd> commentRangeEnd = System::MakeObject<CommentRangeEnd>(doc, comment->get_Id());
    
    run1->get_ParentNode()->InsertAfter(commentRangeStart, run1);
    run3->get_ParentNode()->InsertAfter(commentRangeEnd, run3);
    commentRangeEnd->get_ParentNode()->InsertAfter(comment, commentRangeEnd);
    
    dataDir = dataDir + GetOutputFilePath(u"AnchorComment.doc");
    // Save the document.
    doc->Save(dataDir);
    // ExEnd:AnchorComment
    std::cout << "\nComment anchored successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
