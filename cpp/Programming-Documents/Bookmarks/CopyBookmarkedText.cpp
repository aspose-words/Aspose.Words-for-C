#include "stdafx.h"
#include "examples.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <Model/Text/Range.h>
#include <Model/Text/Paragraph.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/Body.h>
#include <Model/Saving/SaveOutputParameters.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/CompositeNode.h>
#include <Model/Importing/NodeImporter.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Document/Document.h>
#include <Model/Bookmarks/BookmarkStart.h>
#include <Model/Bookmarks/BookmarkEnd.h>
#include <Model/Bookmarks/BookmarkCollection.h>
#include <Model/Bookmarks/Bookmark.h>

using namespace Aspose::Words;
namespace {

void AppendBookmarkedText(const System::SharedPtr<NodeImporter>& importer, const System::SharedPtr<Bookmark>& srcBookmark, const System::SharedPtr<CompositeNode>& dstNode)
{
    // This is the paragraph that contains the beginning of the bookmark.
    System::SharedPtr<Paragraph> startPara = System::DynamicCast<Aspose::Words::Paragraph>(srcBookmark->get_BookmarkStart()->get_ParentNode());
    
    // This is the paragraph that contains the end of the bookmark.
    System::SharedPtr<Paragraph> endPara = System::DynamicCast<Aspose::Words::Paragraph>(srcBookmark->get_BookmarkEnd()->get_ParentNode());
    
    if ((startPara == nullptr) || (endPara == nullptr))
    {
        throw System::InvalidOperationException(u"Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");
    }
    
    // Limit ourselves to a reasonably simple scenario.
    if (startPara->get_ParentNode() != endPara->get_ParentNode())
    {
        throw System::InvalidOperationException(u"Start and end paragraphs have different parents, cannot handle this scenario yet.");
    }
    
    // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
    // Therefore the node at which we stop is one after the end paragraph.
    System::SharedPtr<Node> endNode = endPara->get_NextSibling();
    
    // This is the loop to go through all paragraph-level nodes in the bookmark.
    for (System::SharedPtr<Node> curNode = startPara; curNode != endNode; curNode = curNode->get_NextSibling())
    {
        // This creates a copy of the current node and imports it (makes it valid) in the context
        // Of the destination document. Importing means adjusting styles and list identifiers correctly.
        System::SharedPtr<Node> newNode = importer->ImportNode(curNode, true);
        
        // Now we simply append the new node to the destination.
        dstNode->AppendChild(newNode);
    }
}

}

void CopyBookmarkedText()
{
    // The path to the documents directory.
    System::String dataDir = GetDataDir_WorkingWithBookmarks();
    
    // Load the source document.
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"Template.doc");
    
    // This is the bookmark whose content we want to copy.
    System::SharedPtr<Bookmark> srcBookmark = srcDoc->get_Range()->get_Bookmarks()->idx_get(u"ntf010145060");
    
    // We will be adding to this document.
    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>();
    
    // Let's say we will be appending to the end of the body of the last section.
    System::SharedPtr<CompositeNode> dstNode = dstDoc->get_LastSection()->get_Body();
    
    // It is a good idea to use this import context object because multiple nodes are being imported.
    // If you import multiple times without a single context, it will result in many styles created.
    auto importer = System::MakeObject<NodeImporter>(srcDoc, dstDoc, Aspose::Words::ImportFormatMode::KeepSourceFormatting);
    
    // Do it once.
    AppendBookmarkedText(importer, srcBookmark, dstNode);
    
    // Do it one more time for fun.
    AppendBookmarkedText(importer, srcBookmark, dstNode);
    
    dataDir = dataDir + GetOutputFilePath(u"CopyBookmarkedText.doc");
    // Save the finished document.
    dstDoc->Save(dataDir);
    
    std::cout << "\nBookmark copied successfully.\nFile saved at " << dataDir.ToUtf8String() << '\n';
}
