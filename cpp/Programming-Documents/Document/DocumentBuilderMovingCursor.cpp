#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

using namespace Aspose::Words;

namespace
{
    void CursorPosition(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderCursorPosition
        // Shows how to access the current node in a document builder.
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        System::SharedPtr<Node> curNode = builder->get_CurrentNode();
        System::SharedPtr<Paragraph> curParagraph = builder->get_CurrentParagraph();
        // ExEnd:DocumentBuilderCursorPosition
        std::cout << "Cursor move to paragraph: " << curParagraph->GetText().ToUtf8String() << std::endl;
    }

    void MoveToNode(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToNode
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->MoveTo(doc->get_FirstSection()->get_Body()->get_LastParagraph());
        // ExEnd:DocumentBuilderMoveToNode
        std::cout << "Cursor move to required node." << std::endl;
    }

    void MoveToDocumentStartEnd(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToDocumentStartEnd
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->MoveToDocumentEnd();
        std::cout << "This is the end of the document." << std::endl;

        builder->MoveToDocumentStart();
        std::cout << "This is the beginning of the document." << std::endl;
        // ExEnd:DocumentBuilderMoveToDocumentStartEnd
    }

    void MoveToSection(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToSection
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Parameters are 0-index. Moves to third section.
        builder->MoveToSection(2);
        builder->Writeln(u"This is the 3rd section.");
        // ExEnd:DocumentBuilderMoveToSection
    }

    void HeadersAndFooters(System::String const &outputDataDir)
    {
        // ExStart:DocumentBuilderHeadersAndFooters
        // Create a blank document.
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Specify that we want headers and footers different for first, even and odd pages.
        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);
        builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);

        // Create the headers.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->Write(u"Header First");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderEven);
        builder->Write(u"Header Even");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Header Odd");

        // Create three pages in the document.
        builder->MoveToSection(0);
        builder->Writeln(u"Page1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page3");

        System::String outputPath = outputDataDir + u"DocumentBuilderMovingCursor.HeadersAndFooters.doc";
        doc->Save(outputPath);
        // ExEnd:DocumentBuilderHeadersAndFooters
        std::cout << "Headers and footers created successfully using DocumentBuilder." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void MoveToParagraph(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToParagraph
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Parameters are 0-index. Moves to third paragraph.
        builder->MoveToParagraph(2, 0);
        builder->Writeln(u"This is the 3rd paragraph.");
        // ExEnd:DocumentBuilderMoveToParagraph
    }

    void MoveToTableCell(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToTableCell
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
        builder->MoveToCell(1, 2, 4, 0);
        builder->Writeln(u"Hello World!");
        // ExEnd:DocumentBuilderMoveToTableCell
    }

    void MoveToBookmark(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToBookmark
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->MoveToBookmark(u"CoolBookmark");
        builder->Writeln(u"This is a very cool bookmark.");
        // ExEnd:DocumentBuilderMoveToBookmark
    }

    void MoveToBookmarkEnd(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToBookmarkEnd
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->MoveToBookmark(u"CoolBookmark", false, true);
        builder->Writeln(u"This is a very cool bookmark.");
        // ExEnd:DocumentBuilderMoveToBookmarkEnd
    }

    void MoveToMergeField(System::String const &inputDataDir)
    {
        // ExStart:DocumentBuilderMoveToMergeField
        System::SharedPtr<Document> doc = System::MakeObject<Document>(inputDataDir + u"DocumentBuilder.doc");
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->MoveToMergeField(u"NiceMergeField");
        builder->Writeln(u"This is a very nice merge field.");
        // ExEnd:DocumentBuilderMoveToMergeField
    }
}

void DocumentBuilderMovingCursor()
{
    std::cout << "DocumentBuilderMovingCursor example started." << std::endl;
    // The path to the documents directories.
    System::String inputDataDir = GetInputDataDir_WorkingWithDocument();
    System::String outputDataDir = GetOutputDataDir_WorkingWithDocument();
    CursorPosition(inputDataDir);
    MoveToNode(inputDataDir);
    MoveToDocumentStartEnd(inputDataDir);
    MoveToSection(inputDataDir);
    HeadersAndFooters(outputDataDir);
    MoveToParagraph(inputDataDir);
    MoveToTableCell(inputDataDir);
    MoveToBookmark(inputDataDir);
    MoveToBookmarkEnd(inputDataDir);
    MoveToMergeField(inputDataDir);
    std::cout << "DocumentBuilderMovingCursor example finished." << std::endl << std::endl;
}
