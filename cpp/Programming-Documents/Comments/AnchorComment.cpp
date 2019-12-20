#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeStart.h>
#include <Aspose.Words.Cpp/Model/Text/CommentRangeEnd.h>
#include <Aspose.Words.Cpp/Model/Text/Comment.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/CompositeNode.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

using namespace Aspose::Words;

void AnchorComment()
{
    std::cout << "AnchorComment example started." << std::endl;
    // ExStart:AnchorComment
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithComments();
    System::SharedPtr<Document> doc = System::MakeObject<Document>();

    System::SharedPtr<Paragraph> para1 = System::MakeObject<Paragraph>(doc);
    System::SharedPtr<Run> run1 = System::MakeObject<Run>(doc, u"Some ");
    System::SharedPtr<Run> run2 = System::MakeObject<Run>(doc, u"text ");
    para1->AppendChild(run1);
    para1->AppendChild(run2);
    doc->get_FirstSection()->get_Body()->AppendChild(para1);

    System::SharedPtr<Paragraph> para2 = System::MakeObject<Paragraph>(doc);
    System::SharedPtr<Run> run3 = System::MakeObject<Run>(doc, u"is ");
    System::SharedPtr<Run> run4 = System::MakeObject<Run>(doc, u"added ");
    para2->AppendChild(run3);
    para2->AppendChild(run4);
    doc->get_FirstSection()->get_Body()->AppendChild(para2);

    System::SharedPtr<Comment> comment = System::MakeObject<Comment>(doc, u"Awais Hafeez", u"AH", System::DateTime::get_Today());
    comment->get_Paragraphs()->Add(System::MakeObject<Paragraph>(doc));
    comment->get_FirstParagraph()->get_Runs()->Add(System::MakeObject<Run>(doc, u"Comment text."));
    
    System::SharedPtr<CommentRangeStart> commentRangeStart = System::MakeObject<CommentRangeStart>(doc, comment->get_Id());
    System::SharedPtr<CommentRangeEnd> commentRangeEnd = System::MakeObject<CommentRangeEnd>(doc, comment->get_Id());

    run1->get_ParentNode()->InsertAfter(commentRangeStart, run1);
    run3->get_ParentNode()->InsertAfter(commentRangeEnd, run3);
    commentRangeEnd->get_ParentNode()->InsertAfter(comment, commentRangeEnd);

    System::String outputPath = outputDataDir + u"AnchorComment.doc";
    // Save the document.
    doc->Save(outputPath);
    // ExEnd:AnchorComment
    std::cout << "Comment anchored successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "AnchorComment example finished." << std::endl << std::endl;
}
