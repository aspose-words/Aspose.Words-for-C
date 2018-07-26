#include "../../examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/HeaderFooterType.h>
#include <Model/Sections/Body.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Document/Document.h>
#include <Model/Document/BreakType.h>

using namespace Aspose::Words;

namespace
{

void CursorPosition(System::String dataDir)
{
    // ExStart:DocumentBuilderCursorPosition
    // Shows how to access the current node in a document builder.
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    System::SharedPtr<Node> curNode = builder->get_CurrentNode();
    System::SharedPtr<Paragraph> curParagraph = builder->get_CurrentParagraph();
    // ExEnd:DocumentBuilderCursorPosition
    std::cout << "\nCursor move to paragraph: " << curParagraph->GetText().ToUtf8String() << '\n';
}

void MoveToNode(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToNode
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    builder->MoveTo(doc->get_FirstSection()->get_Body()->get_LastParagraph());
    // ExEnd:DocumentBuilderMoveToNode   
    std::cout << "\nCursor move to required node.\n";
}

void MoveToDocumentStartEnd(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToDocumentStartEnd
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->MoveToDocumentEnd();
    std::cout << "\nThis is the end of the document.\n";
    
    builder->MoveToDocumentStart();
    std::cout << "\nThis is the beginning of the document.\n";
    // ExEnd:DocumentBuilderMoveToDocumentStartEnd
}

void MoveToSection(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToSection
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Parameters are 0-index. Moves to third section.
    builder->MoveToSection(2);
    builder->Writeln(u"This is the 3rd section.");
    // ExEnd:DocumentBuilderMoveToSection
}

void HeadersAndFooters(System::String dataDir)
{
    // ExStart:DocumentBuilderHeadersAndFooters
    // Create a blank document.
    System::SharedPtr<Document> doc = System::MakeObject<Document>();
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Specify that we want headers and footers different for first, even and odd pages.
    builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);
    builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);
    
    // Create the headers.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderFirst);
    builder->Write(u"Header First");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderEven);
    builder->Write(u"Header Even");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Header Odd");
    
    // Create three pages in the document.
    builder->MoveToSection(0);
    builder->Writeln(u"Page1");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page2");
    builder->InsertBreak(Aspose::Words::BreakType::PageBreak);
    builder->Writeln(u"Page3");
    
    dataDir = dataDir + GetOutputFilePath(u"DocumentBuilderMovingCursor.HeadersAndFooters.doc");
    doc->Save(dataDir);
    // ExEnd:DocumentBuilderHeadersAndFooters   
    std::cout << "\nHeaders and footers created successfully using DocumentBuilder.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}

void MoveToParagraph(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToParagraph
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // Parameters are 0-index. Moves to third paragraph.
    builder->MoveToParagraph(2, 0);
    builder->Writeln(u"This is the 3rd paragraph.");
    // ExEnd:DocumentBuilderMoveToParagraph
}

void MoveToTableCell(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToTableCell
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
    builder->MoveToCell(1, 2, 4, 0);
    builder->Writeln(u"Hello World!");
    // ExEnd:DocumentBuilderMoveToTableCell
}

void MoveToBookmark(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToBookmark
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->MoveToBookmark(u"CoolBookmark");
    builder->Writeln(u"This is a very cool bookmark.");
    // ExEnd:DocumentBuilderMoveToBookmark
}

void MoveToBookmarkEnd(System::String dataDir)
{
    // ExStart:DocumentBuilderMoveToBookmarkEnd
    System::SharedPtr<Document> doc = System::MakeObject<Document>(dataDir + u"DocumentBuilder.doc");
    System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
    
    builder->MoveToBookmark(u"CoolBookmark", false, true);
    builder->Writeln(u"This is a very cool bookmark.");
    // ExEnd:DocumentBuilderMoveToBookmarkEnd
}
}

void DocumentBuilderMovingCursor()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithDocument();
    CursorPosition(dataDir);
    MoveToNode(dataDir);
    MoveToDocumentStartEnd(dataDir);
    MoveToSection(dataDir);
    HeadersAndFooters(dataDir);
    MoveToParagraph(dataDir);
    MoveToTableCell(dataDir);
    MoveToBookmark(dataDir);
    MoveToBookmarkEnd(dataDir);
    
}
