#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BookmarkEnd.h>
#include <Aspose.Words.Cpp/BookmarkStart.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/ImportFormatOptions.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeImporter.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Orientation.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/SectionStart.h>
#include <system/array.h>
#include <system/collections/dictionary.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/object_ext.h>
#include <system/text/string_builder.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Programming_with_Documents { namespace Working_with_Document {

class JoinAndAppendDocuments : public DocsExamplesBase
{
public:
    void SimpleAppendDocument()
    {
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Append the source document to the destination document using no extra options.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.SimpleAppendDocument.docx");
    }

    void AppendDocument()
    {
        //ExStart:AppendDocumentManually
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Loop through all sections in the source document.
        // Section nodes are immediate children of the Document node so we can just enumerate the Document.
        for (const auto& srcSection : System::IterateOver<Section>(srcDoc))
        {
            // Because we are copying a section from one document to another,
            // it is required to import the Section node into the destination document.
            // This adjusts any document-specific references to styles, lists, etc.
            // Importing a node creates a copy of the original node, but the copy
            // ss ready to be inserted into the destination document.
            SharedPtr<Node> dstSection = dstDoc->ImportNode(srcSection, true, ImportFormatMode::KeepSourceFormatting);

            // Now the new section node can be appended to the destination document.
            dstDoc->AppendChild(dstSection);
        }

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.AppendDocument.docx");
        //ExEnd:AppendDocumentManually
    }

    void AppendDocumentToBlank()
    {
        //ExStart:AppendDocumentToBlank
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>();

        // The destination document is not empty, often causing a blank page to appear before the appended document.
        // This is due to the base document having an empty section and the new document being started on the next page.
        // Remove all content from the destination document before appending.
        dstDoc->RemoveAllChildren();
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.AppendDocumentToBlank.docx");
        //ExEnd:AppendDocumentToBlank
    }

    void AppendWithImportFormatOptions()
    {
        //ExStart:AppendWithImportFormatOptions
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source with list.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Document destination with list.docx");

        // Specify that if numbering clashes in source and destination documents,
        // then numbering from the source document will be used.
        auto options = MakeObject<ImportFormatOptions>();
        options->set_KeepSourceNumbering(true);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles, options);
        //ExEnd:AppendWithImportFormatOptions
    }

    void ConvertNumPageFields()
    {
        //ExStart:ConvertNumPageFields
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Restart the page numbering on the start of the source document.
        srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);
        srcDoc->get_FirstSection()->get_PageSetup()->set_PageStartingNumber(1);

        // Append the source document to the end of the destination document.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        // After joining the documents the NUMPAGE fields will now display the total number of pages which
        // is undesired behavior. Call this method to fix them by replacing them with PAGEREF fields.
        ConvertNumPageFieldsToPageRef(dstDoc);

        // This needs to be called in order to update the new fields with page numbers.
        dstDoc->UpdatePageLayout();

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.ConvertNumPageFields.docx");
        //ExEnd:ConvertNumPageFields
    }

    //ExStart:ConvertNumPageFieldsToPageRef
    void ConvertNumPageFieldsToPageRef(SharedPtr<Document> doc)
    {
        // This is the prefix for each bookmark, which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        const String bookmarkPrefix = u"_SubDocumentEnd";
        const String numPagesFieldName = u"NUMPAGES";
        const String pageRefFieldName = u"PAGEREF";

        // Defines the number of page restarts encountered and, therefore,
        // the number of "sub" documents found within this document.
        int subDocumentCount = 0;

        auto builder = MakeObject<DocumentBuilder>(doc);

        for (const auto& section : System::IterateOver<Section>(doc->get_Sections()))
        {
            // This section has its page numbering restarted to treat this as the start of a sub-document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if (section->get_PageSetup()->get_RestartPageNumbering())
            {
                // Don't do anything if this is the first section of the document.
                // This part of the code will insert the bookmark marking the end of the previous sub-document so,
                // therefore, it does not apply to the first section in the document.
                if (!System::ObjectExt::Equals(section, doc->get_FirstSection()))
                {
                    // Get the previous section and the last node within the body of that section.
                    auto prevSection = System::ExplicitCast<Section>(section->get_PreviousSibling());
                    SharedPtr<Node> lastNode = prevSection->get_Body()->get_LastChild();

                    builder->MoveTo(lastNode);

                    // This bookmark represents the end of the sub-document.
                    builder->StartBookmark(bookmarkPrefix + subDocumentCount);
                    builder->EndBookmark(bookmarkPrefix + subDocumentCount);

                    // Increase the sub-document count to insert the correct bookmarks.
                    subDocumentCount++;
                }
            }

            // The last section needs the ending bookmark to signal that it is the end of the current sub-document.
            if (System::ObjectExt::Equals(section, doc->get_LastSection()))
            {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                SharedPtr<Node> lastNode = doc->get_LastSection()->get_Body()->get_LastChild();

                builder->MoveTo(lastNode);
                builder->StartBookmark(bookmarkPrefix + subDocumentCount);
                builder->EndBookmark(bookmarkPrefix + subDocumentCount);
            }

            // Iterate through each NUMPAGES field in the section and replace it with a PAGEREF field
            // referring to the bookmark of the current sub-document. This bookmark is positioned at the end
            // of the sub-document but does not exist yet. It is inserted when a section with restart page numbering
            // or the last section is encountered.
            ArrayPtr<SharedPtr<Node>> nodes = section->GetChildNodes(NodeType::FieldStart, true)->ToArray();

            for (const auto& it_ : nodes)
            {
                SharedPtr<FieldStart> fieldStart = System::ExplicitCast<FieldStart>(it_);
                {
                    if (fieldStart->get_FieldType() == FieldType::FieldNumPages)
                    {
                        String fieldCode = GetFieldCode(fieldStart);
                        // Since the NUMPAGES field does not take any additional parameters,
                        // we can assume the field's remaining part. Code after the field name is the switches.
                        // We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                        String fieldSwitches = fieldCode.Replace(numPagesFieldName, u"").Trim();

                        // Inserting the new field directly at the FieldStart node of the original field will cause
                        // the new field not to pick up the original field's formatting. To counter this,
                        // insert the field just before the original field if a previous run cannot be found,
                        // we are forced to use the FieldStart node.
                        SharedPtr<Node> previousNode = fieldStart->get_PreviousSibling() != nullptr ? fieldStart->get_PreviousSibling() : fieldStart;

                        // Insert a PAGEREF field at the same position as the field.
                        builder->MoveTo(previousNode);

                        SharedPtr<Field> newField =
                            builder->InsertField(String::Format(u" {0} {1}{2} {3} ", pageRefFieldName, bookmarkPrefix, subDocumentCount, fieldSwitches));

                        // The field will be inserted before the referenced node. Move the node before the field instead.
                        previousNode->get_ParentNode()->InsertBefore(previousNode, newField->get_Start());

                        // Remove the original NUMPAGES field from the document.
                        RemoveField(fieldStart);
                    }
                }
            }
        }
    }
    //ExEnd:ConvertNumPageFieldsToPageRef

    //ExStart:GetRemoveField
    void RemoveField(SharedPtr<FieldStart> fieldStart)
    {
        bool isRemoving = true;

        SharedPtr<Node> currentNode = fieldStart;
        while (currentNode != nullptr && isRemoving)
        {
            if (currentNode->get_NodeType() == NodeType::FieldEnd)
            {
                isRemoving = false;
            }

            SharedPtr<Node> nextNode = currentNode->NextPreOrder(currentNode->get_Document());
            currentNode->Remove();
            currentNode = nextNode;
        }
    }

    String GetFieldCode(SharedPtr<FieldStart> fieldStart)
    {
        auto builder = MakeObject<System::Text::StringBuilder>();

        for (SharedPtr<Node> node = fieldStart;
             node != nullptr && node->get_NodeType() != NodeType::FieldSeparator && node->get_NodeType() != NodeType::FieldEnd;
             node = node->NextPreOrder(node->get_Document()))
        {
            // Use text only of Run nodes to avoid duplication.
            if (node->get_NodeType() == NodeType::Run)
            {
                builder->Append(node->GetText());
            }
        }

        return builder->ToString();
    }
    //ExEnd:GetRemoveField

    void DifferentPageSetup()
    {
        //ExStart:DifferentPageSetup
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

        // Restart the page numbering on the start of the source document.
        srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);
        srcDoc->get_FirstSection()->get_PageSetup()->set_PageStartingNumber(1);

        // To ensure this does not happen when the source document has different page setup settings, make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document,
        // this will need to be repeated for those sections.
        srcDoc->get_FirstSection()->get_PageSetup()->set_PageWidth(dstDoc->get_LastSection()->get_PageSetup()->get_PageWidth());
        srcDoc->get_FirstSection()->get_PageSetup()->set_PageHeight(dstDoc->get_LastSection()->get_PageSetup()->get_PageHeight());
        srcDoc->get_FirstSection()->get_PageSetup()->set_Orientation(dstDoc->get_LastSection()->get_PageSetup()->get_Orientation());

        // Iterate through all sections in the source document.
        for (const auto& para : System::IterateOver<Paragraph>(srcDoc->GetChildNodes(NodeType::Paragraph, true)))
        {
            para->get_ParagraphFormat()->set_KeepWithNext(true);
        }

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.DifferentPageSetup.docx");
        //ExEnd:DifferentPageSetup
    }

    void JoinContinuous()
    {
        //ExStart:JoinContinuous
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Make the document appear straight after the destination documents content.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);
        // Append the source document using the original styles found in the source document.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.JoinContinuous.docx");
        //ExEnd:JoinContinuous
    }

    void JoinNewPage()
    {
        //ExStart:JoinNewPage
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Set the appended document to start on a new page.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        // Append the source document using the original styles found in the source document.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.JoinNewPage.docx");
        //ExEnd:JoinNewPage
    }

    void KeepSourceFormatting()
    {
        //ExStart:KeepSourceFormatting
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto dstDoc = MakeObject<Document>();
        dstDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Destination document text. ");

        auto srcDoc = MakeObject<Document>();
        srcDoc->get_FirstSection()->get_Body()->AppendParagraph(u"Source document text. ");

        // Append the source document to the destination document.
        // Pass format mode to retain the original formatting of the source document when importing it.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.KeepSourceFormatting.docx");
        //ExEnd:KeepSourceFormatting
    }

    void KeepSourceTogether()
    {
        //ExStart:KeepSourceTogether
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Document destination with list.docx");

        // Set the source document to appear straight after the destination document's content.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

        for (const auto& para : System::IterateOver<Paragraph>(srcDoc->GetChildNodes(NodeType::Paragraph, true)))
        {
            para->get_ParagraphFormat()->set_KeepWithNext(true);
        }

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.KeepSourceTogether.docx");
        //ExEnd:KeepSourceTogether
    }

    void ListKeepSourceFormatting()
    {
        //ExStart:ListKeepSourceFormatting
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Document destination with list.docx");

        // Append the content of the document so it flows continuously.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.ListKeepSourceFormatting.docx");
        //ExEnd:ListKeepSourceFormatting
    }

    void ListUseDestinationStyles()
    {
        //ExStart:ListUseDestinationStyles
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Document destination with list.docx");

        // Set the source document to continue straight after the end of the destination document.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::Continuous);

        // Keep track of the lists that are created.
        auto newLists = MakeObject<System::Collections::Generic::Dictionary<int, SharedPtr<Aspose::Words::Lists::List>>>();

        for (const auto& para : System::IterateOver<Paragraph>(srcDoc->GetChildNodes(NodeType::Paragraph, true)))
        {
            if (para->get_IsListItem())
            {
                int listId = para->get_ListFormat()->get_List()->get_ListId();

                // Check if the destination document contains a list with this ID already. If it does, then this may
                // cause the two lists to run together. Create a copy of the list in the source document instead.
                if (dstDoc->get_Lists()->GetListByListId(listId) != nullptr)
                {
                    SharedPtr<Aspose::Words::Lists::List> currentList;
                    // A newly copied list already exists for this ID, retrieve the stored list,
                    // and use it on the current paragraph.
                    if (newLists->ContainsKey(listId))
                    {
                        currentList = newLists->idx_get(listId);
                    }
                    else
                    {
                        // Add a copy of this list to the document and store it for later reference.
                        currentList = srcDoc->get_Lists()->AddCopy(para->get_ListFormat()->get_List());
                        newLists->Add(listId, currentList);
                    }

                    // Set the list of this paragraph to the copied list.
                    para->get_ListFormat()->set_List(currentList);
                }
            }
        }

        // Append the source document to end of the destination document.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.ListUseDestinationStyles.docx");
        //ExEnd:ListUseDestinationStyles
    }

    void RestartPageNumbering()
    {
        //ExStart:RestartPageNumbering
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        srcDoc->get_FirstSection()->get_PageSetup()->set_RestartPageNumbering(true);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.RestartPageNumbering.docx");
        //ExEnd:RestartPageNumbering
    }

    void UpdatePageLayout()
    {
        //ExStart:UpdatePageLayout
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // If the destination document is rendered to PDF, image etc.
        // or UpdatePageLayout is called before the source document. Is appended,
        // then any changes made after will not be reflected in the rendered output
        dstDoc->UpdatePageLayout();

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        // For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        // If not called again, the appended document will not appear in the output of the next rendering.
        dstDoc->UpdatePageLayout();

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.UpdatePageLayout.docx");
        //ExEnd:UpdatePageLayout
    }

    void UseDestinationStyles()
    {
        //ExStart:UseDestinationStyles
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Append the source document using the styles of the destination document.
        dstDoc->AppendDocument(srcDoc, ImportFormatMode::UseDestinationStyles);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.UseDestinationStyles.docx");
        //ExEnd:UseDestinationStyles
    }

    void SmartStyleBehavior()
    {
        //ExStart:SmartStyleBehavior
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");
        auto builder = MakeObject<DocumentBuilder>(dstDoc);

        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);

        auto options = MakeObject<ImportFormatOptions>();
        options->set_SmartStyleBehavior(true);

        builder->InsertDocument(srcDoc, ImportFormatMode::UseDestinationStyles, options);
        builder->get_Document()->Save(ArtifactsDir + u"JoinAndAppendDocuments.SmartStyleBehavior.docx");
        //ExEnd:SmartStyleBehavior
    }

    void InsertDocument()
    {
        //ExStart:InsertDocumentWithBuilder
        //GistId:db2dfc4150d7c714bcac3782ae241d03
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");
        auto builder = MakeObject<DocumentBuilder>(dstDoc);

        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::PageBreak);

        builder->InsertDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);
        builder->get_Document()->Save(ArtifactsDir + u"JoinAndAppendDocuments.InsertDocument.docx");
        //ExEnd:InsertDocumentWithBuilder
    }

    void KeepSourceNumbering()
    {
        //ExStart:KeepSourceNumbering
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Keep source list formatting when importing numbered paragraphs.
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_KeepSourceNumbering(true);

        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting, importFormatOptions);

        SharedPtr<ParagraphCollection> srcParas = srcDoc->get_FirstSection()->get_Body()->get_Paragraphs();
        for (const auto& srcPara : System::IterateOver<Paragraph>(srcParas))
        {
            SharedPtr<Node> importedNode = importer->ImportNode(srcPara, false);
            dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
        }

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.KeepSourceNumbering.docx");
        //ExEnd:KeepSourceNumbering
    }

    void IgnoreTextBoxes()
    {
        //ExStart:IgnoreTextBoxes
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Keep the source text boxes formatting when importing.
        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_IgnoreTextBoxes(false);

        auto importer = MakeObject<NodeImporter>(srcDoc, dstDoc, ImportFormatMode::KeepSourceFormatting, importFormatOptions);

        SharedPtr<ParagraphCollection> srcParas = srcDoc->get_FirstSection()->get_Body()->get_Paragraphs();
        for (const auto& srcPara : System::IterateOver<Paragraph>(srcParas))
        {
            SharedPtr<Node> importedNode = importer->ImportNode(srcPara, true);
            dstDoc->get_FirstSection()->get_Body()->AppendChild(importedNode);
        }

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.IgnoreTextBoxes.docx");
        //ExEnd:IgnoreTextBoxes
    }

    void IgnoreHeaderFooter()
    {
        //ExStart:IgnoreHeaderFooter
        auto srcDocument = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDocument = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        auto importFormatOptions = MakeObject<ImportFormatOptions>();
        importFormatOptions->set_IgnoreHeaderFooter(false);

        dstDocument->AppendDocument(srcDocument, ImportFormatMode::KeepSourceFormatting, importFormatOptions);

        dstDocument->Save(ArtifactsDir + u"JoinAndAppendDocuments.IgnoreHeaderFooter.docx");
        //ExEnd:IgnoreHeaderFooter
    }

    void LinkHeadersFooters()
    {
        //ExStart:LinkHeadersFooters
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Set the appended document to appear on a new page.
        srcDoc->get_FirstSection()->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        // Link the headers and footers in the source document to the previous section.
        // This will override any headers or footers already found in the source document.
        srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(true);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.LinkHeadersFooters.docx");
        //ExEnd:LinkHeadersFooters
    }

    void RemoveSourceHeadersFooters()
    {
        //ExStart:RemoveSourceHeadersFooters
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Remove the headers and footers from each of the sections in the source document.
        for (const auto& section : System::IterateOver<Section>(srcDoc->get_Sections()))
        {
            section->ClearHeadersFooters();
        }

        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        // document. This should set to false to avoid this behavior.
        srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(false);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.RemoveSourceHeadersFooters.docx");
        //ExEnd:RemoveSourceHeadersFooters
    }

    void UnlinkHeadersFooters()
    {
        //ExStart:UnlinkHeadersFooters
        auto srcDoc = MakeObject<Document>(MyDir + u"Document source.docx");
        auto dstDoc = MakeObject<Document>(MyDir + u"Northwind traders.docx");

        // Unlink the headers and footers in the source document to stop this
        // from continuing the destination document's headers and footers.
        srcDoc->get_FirstSection()->get_HeadersFooters()->LinkToPrevious(false);

        dstDoc->AppendDocument(srcDoc, ImportFormatMode::KeepSourceFormatting);

        dstDoc->Save(ArtifactsDir + u"JoinAndAppendDocuments.UnlinkHeadersFooters.docx");
        //ExEnd:UnlinkHeadersFooters
    }
};

}}} // namespace DocsExamples::Programming_with_Documents::Working_with_Document
