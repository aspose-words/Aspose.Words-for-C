#include "stdafx.h"
#include "examples.h"

#include <system/enumerator_adapter.h>
#include <Model/Document/Document.h>
#include <Model/Document/DocumentBuilder.h>
#include <Model/Fields/Field.h>
#include <Model/Fields/Nodes/FieldStart.h>
#include <Model/Importing/ImportFormatMode.h>
#include <Model/Nodes/Node.h>
#include <Model/Nodes/NodeType.h>
#include <Model/Sections/Body.h>
#include <Model/Sections/PageSetup.h>
#include <Model/Sections/Section.h>
#include <Model/Sections/SectionCollection.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace
{
    void RemoveField(System::SharedPtr<FieldStart> fieldStart)
    {
        System::SharedPtr<Node> currentNode = fieldStart;
        bool isRemoving = true;
        while (currentNode != nullptr && isRemoving)
        {
            if (currentNode->get_NodeType() == NodeType::FieldEnd)
            {
                isRemoving = false;
            }

            System::SharedPtr<Node> nextNode = currentNode->NextPreOrder(currentNode->get_Document());
            currentNode->Remove();
            currentNode = nextNode;
        }
    }

    System::String GetFieldCode(System::SharedPtr<FieldStart> fieldStart)
    {
        System::SharedPtr<System::Text::StringBuilder> builder = System::MakeObject<System::Text::StringBuilder>();

        for (System::SharedPtr<Node> node = fieldStart; node != nullptr && node->get_NodeType() != NodeType::FieldSeparator && node->get_NodeType() != NodeType::FieldEnd; node = node->NextPreOrder(node->get_Document()))
        {
            // Use text only of Run nodes to avoid duplication.
            if (node->get_NodeType() == NodeType::Run)
            {
                builder->Append(node->GetText());
            }
        }
        return builder->ToString();
    }

    void ConvertNumPageFieldsToPageRef(System::SharedPtr<Document> doc)
    {
        // This is the prefix for each bookmark which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        const System::String bookmarkPrefix = u"_SubDocumentEnd";
        // Field name of the NUMPAGES field.
        const System::String numPagesFieldName = u"NUMPAGES";
        // Field name of the PAGEREF field.
        const System::String pageRefFieldName = u"PAGEREF";

        // Create a new DocumentBuilder which is used to insert the bookmarks and replacement fields.
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        // Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
        // Found within this document.
        int32_t subDocumentCount = 0;

        // Iterate through all sections in the document.
        for (System::SharedPtr<Section> section : System::IterateOver<Section>(doc->get_Sections()))
        {
            // This section has it's page numbering restarted so we will treat this as the start of a sub document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section->get_PageSetup()->get_RestartPageNumbering())
            {
                // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                // The end of the previous sub document so therefore it is not applicable for first section in the document.
                if (!System::ObjectExt::Equals(section, doc->get_FirstSection()))
                {
                    // Get the previous section and the last node within the body of that section.
                    System::SharedPtr<Section> prevSection = System::DynamicCast<Section>(section->get_PreviousSibling());
                    System::SharedPtr<Node> lastNode = prevSection->get_Body()->get_LastChild();

                    // Use the DocumentBuilder to move to this node and insert the bookmark there.
                    // This bookmark represents the end of the sub document.
                    builder->MoveTo(lastNode);
                    builder->StartBookmark(bookmarkPrefix + subDocumentCount);
                    builder->EndBookmark(bookmarkPrefix + subDocumentCount);

                    // Increase the subdocument count to insert the correct bookmarks.
                    subDocumentCount++;
                }
            }

            // The last section simply needs the ending bookmark to signal that it is the end of the current sub document.
            if (System::ObjectExt::Equals(section, doc->get_LastSection()))
            {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                System::SharedPtr<Node> lastNode = doc->get_LastSection()->get_Body()->get_LastChild();
                builder->MoveTo(lastNode);
                builder->StartBookmark(bookmarkPrefix + subDocumentCount);
                builder->EndBookmark(bookmarkPrefix + subDocumentCount);
            }

            // Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
            // This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last 
            // Section is encountered.
            for (System::SharedPtr<Node> node : System::IterateOver(section->GetChildNodes(NodeType::FieldStart, true)))
            {
                System::SharedPtr<FieldStart> fieldStart = System::DynamicCast<FieldStart>(node);
                if (fieldStart->get_FieldType() == FieldType::FieldNumPages)
                {
                    // Get the field code.
                    System::String fieldCode = GetFieldCode(fieldStart);
                    // Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                    // Code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    System::String fieldSwitches = fieldCode.Replace(numPagesFieldName, u"").Trim();

                    // Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                    // Not pick up the formatting of the original field. To counter this insert the field just before the original field
                    System::SharedPtr<Node> previousNode = fieldStart->get_PreviousSibling();

                    // If a previous run cannot be found then we are forced to use the FieldStart node.
                    if (previousNode == nullptr)
                    {
                        previousNode = fieldStart;
                    }

                    // Insert a PAGEREF field at the same position as the field.
                    builder->MoveTo(previousNode);
                    // This will insert a new field with a code like " PAGEREF _SubDocumentEnd0 *\MERGEFORMAT ".
                    System::SharedPtr<Field> newField = builder->InsertField(System::String::Format(u" {0} {1}{2} {3} ", pageRefFieldName, bookmarkPrefix, subDocumentCount, fieldSwitches));

                    // The field will be inserted before the referenced node. Move the node before the field instead.
                    previousNode->get_ParentNode()->InsertBefore(previousNode, newField->get_FieldStart());

                    // Remove the original NUMPAGES field from the document.
                    RemoveField(fieldStart);
                }
            }
        }
    }
}

void ConvertNumPageFields()
{
    std::cout << "ConvertNumPageFields example started." << std::endl;
    // ExStart:ConvertNumPageFields
    // The path to the documents directory.
    System::String dataDir = GetDataDir_JoiningAndAppending();

    System::SharedPtr<Document> dstDoc = System::MakeObject<Document>(dataDir + u"TestFile.Destination.doc");
    System::SharedPtr<Document> srcDoc = System::MakeObject<Document>(dataDir + u"TestFile.Source.doc");

    // Restart the page numbering on the start of the source document.
    srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);
    srcDoc->get_FirstSection()->get_PageSetup()->set_PageStartingNumber(1);

    // Append the source document to the end of the destination document.
    dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

    // After joining the documents the NUMPAGE fields will now display the total number of pages which 
    // Is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
    ConvertNumPageFieldsToPageRef(dstDoc);

    // This needs to be called in order to update the new fields with page numbers.
    dstDoc->UpdatePageLayout();

    System::String outputPath = dataDir + GetOutputFilePath(u"ConvertNumPageFields.doc");
    dstDoc->Save(outputPath);
    // ExEnd:ConvertNumPageFields
    std::cout << "Document appended successfully with conversion of NUMPAGE fields with PAGEREF fields." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    std::cout << "ConvertNumPageFields example finished." << std::endl << std::endl;
}