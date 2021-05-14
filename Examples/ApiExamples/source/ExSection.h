#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/ControlChar.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/Fields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/PaperSize.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/ProtectionType.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/SectionStart.h>
#include <Aspose.Words.Cpp/TextColumnCollection.h>
#include <drawing/color.h>
#include <system/collections/ienumerable.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/globalization/culture_info.h>
#include <system/linq/enumerable.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/threading/thread.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "DocumentHelper.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;

namespace ApiExamples {

class ExSection : public ApiExampleBase
{
public:
    void Protect()
    {
        //ExStart
        //ExFor:Document.Protect(ProtectionType)
        //ExFor:ProtectionType
        //ExFor:Section.ProtectedForForms
        //ExSummary:Shows how to turn off protection for a section.
        auto doc = MakeObject<Document>();

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Section 1. Hello world!");
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        builder->Writeln(u"Section 2. Hello again!");
        builder->Write(u"Please enter text here: ");
        builder->InsertTextInput(u"TextInput1", TextFormFieldType::Regular, u"", u"Placeholder text", 0);

        // Apply write protection to every section in the document.
        doc->Protect(ProtectionType::AllowOnlyFormFields);

        // Turn off write protection for the first section.
        doc->get_Sections()->idx_get(0)->set_ProtectedForForms(false);

        // In this output document, we will be able to edit the first section freely,
        // and we will only be able to edit the contents of the form field in the second section.
        doc->Save(ArtifactsDir + u"Section.Protect.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Section.Protect.docx");

        ASSERT_FALSE(doc->get_Sections()->idx_get(0)->get_ProtectedForForms());
        ASSERT_TRUE(doc->get_Sections()->idx_get(1)->get_ProtectedForForms());
    }

    void AddRemove()
    {
        //ExStart
        //ExFor:Document.Sections
        //ExFor:Section.Clone
        //ExFor:SectionCollection
        //ExFor:NodeCollection.RemoveAt(Int32)
        //ExSummary:Shows how to add and remove sections in a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");

        ASSERT_EQ(u"Section 1\x000c"
                  u"Section 2",
                  doc->GetText().Trim());

        // Delete the first section from the document.
        doc->get_Sections()->RemoveAt(0);

        ASSERT_EQ(u"Section 2", doc->GetText().Trim());

        // Append a copy of what is now the first section to the end of the document.
        int lastSectionIdx = doc->get_Sections()->get_Count() - 1;
        SharedPtr<Section> newSection = doc->get_Sections()->idx_get(lastSectionIdx)->Clone();
        doc->get_Sections()->Add(newSection);

        ASSERT_EQ(u"Section 2\x000c"
                  u"Section 2",
                  doc->GetText().Trim());
        //ExEnd
    }

    void FirstAndLast()
    {
        //ExStart
        //ExFor:Document.FirstSection
        //ExFor:Document.LastSection
        //ExSummary:Shows how to create a new section with a document builder.
        auto doc = MakeObject<Document>();

        // A blank document contains one section by default,
        // which contains child nodes that we can edit.
        ASSERT_EQ(1, doc->get_Sections()->get_Count());

        // Use a document builder to add text to the first section.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Create a second section by inserting a section break.
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        ASSERT_EQ(2, doc->get_Sections()->get_Count());

        // Each section has its own page setup settings.
        // We can split the text in the second section into two columns.
        // This will not affect the text in the first section.
        doc->get_LastSection()->get_PageSetup()->get_TextColumns()->SetCount(2);
        builder->Writeln(u"Column 1.");
        builder->InsertBreak(BreakType::ColumnBreak);
        builder->Writeln(u"Column 2.");

        ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
        ASSERT_EQ(2, doc->get_LastSection()->get_PageSetup()->get_TextColumns()->get_Count());

        doc->Save(ArtifactsDir + u"Section.Create.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"Section.Create.docx");

        ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
        ASSERT_EQ(2, doc->get_LastSection()->get_PageSetup()->get_TextColumns()->get_Count());
    }

    void CreateManually()
    {
        //ExStart
        //ExFor:Node.GetText
        //ExFor:CompositeNode.RemoveAllChildren
        //ExFor:CompositeNode.AppendChild
        //ExFor:Section
        //ExFor:Section.#ctor
        //ExFor:Section.PageSetup
        //ExFor:PageSetup.SectionStart
        //ExFor:PageSetup.PaperSize
        //ExFor:SectionStart
        //ExFor:PaperSize
        //ExFor:Body
        //ExFor:Body.#ctor
        //ExFor:Paragraph
        //ExFor:Paragraph.#ctor
        //ExFor:Paragraph.ParagraphFormat
        //ExFor:ParagraphFormat
        //ExFor:ParagraphFormat.StyleName
        //ExFor:ParagraphFormat.Alignment
        //ExFor:ParagraphAlignment
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase)
        //ExFor:Run.Text
        //ExFor:Inline.Font
        //ExSummary:Shows how to construct an Aspose.Words document by hand.
        auto doc = MakeObject<Document>();

        // A blank document contains one section, one body and one paragraph.
        // Call the "RemoveAllChildren" method to remove all those nodes,
        // and end up with a document node with no children.
        doc->RemoveAllChildren();

        // This document now has no composite child nodes that we can add content to.
        // If we wish to edit it, we will need to repopulate its node collection.
        // First, create a new section, and then append it as a child to the root document node.
        auto section = MakeObject<Section>(doc);
        doc->AppendChild(section);

        // Set some page setup properties for the section.
        section->get_PageSetup()->set_SectionStart(SectionStart::NewPage);
        section->get_PageSetup()->set_PaperSize(PaperSize::Letter);

        // A section needs a body, which will contain and display all its contents
        // on the page between the section's header and footer.
        auto body = MakeObject<Body>(doc);
        section->AppendChild(body);

        // Create a paragraph, set some formatting properties, and then append it as a child to the body.
        auto para = MakeObject<Paragraph>(doc);

        para->get_ParagraphFormat()->set_StyleName(u"Heading 1");
        para->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);

        body->AppendChild(para);

        // Finally, add some content to do the document. Create a run,
        // set its appearance and contents, and then append it as a child to the paragraph.
        auto run = MakeObject<Run>(doc);
        run->set_Text(u"Hello World!");
        run->get_Font()->set_Color(System::Drawing::Color::get_Red());
        para->AppendChild(run);

        ASSERT_EQ(u"Hello World!", doc->GetText().Trim());

        doc->Save(ArtifactsDir + u"Section.CreateManually.docx");
        //ExEnd
    }

    void EnsureMinimum()
    {
        //ExStart
        //ExFor:NodeCollection.Add
        //ExFor:Section.EnsureMinimum
        //ExFor:SectionCollection.Item(Int32)
        //ExSummary:Shows how to prepare a new section node for editing.
        auto doc = MakeObject<Document>();

        // A blank document comes with a section, which has a body, which in turn has a paragraph.
        // We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
        ASSERT_EQ(NodeType::Section, doc->GetChild(NodeType::Any, 0, true)->get_NodeType());
        ASSERT_EQ(NodeType::Body, doc->get_Sections()->idx_get(0)->GetChild(NodeType::Any, 0, true)->get_NodeType());
        ASSERT_EQ(NodeType::Paragraph, doc->get_Sections()->idx_get(0)->get_Body()->GetChild(NodeType::Any, 0, true)->get_NodeType());

        // If we add a new section like this, it will not have a body, or any other child nodes.
        doc->get_Sections()->Add(MakeObject<Section>(doc));

        ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->GetChildNodes(NodeType::Any, true)->get_Count());

        // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
        doc->get_LastSection()->EnsureMinimum();

        ASSERT_EQ(NodeType::Body, doc->get_Sections()->idx_get(1)->GetChild(NodeType::Any, 0, true)->get_NodeType());
        ASSERT_EQ(NodeType::Paragraph, doc->get_Sections()->idx_get(1)->get_Body()->GetChild(NodeType::Any, 0, true)->get_NodeType());

        doc->get_Sections()->idx_get(0)->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }

    void BodyEnsureMinimum()
    {
        //ExStart
        //ExFor:Section.Body
        //ExFor:Body.EnsureMinimum
        //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
        auto doc = MakeObject<Document>();

        // A blank document contains one section, one body and one paragraph.
        // Call the "RemoveAllChildren" method to remove all those nodes,
        // and end up with a document node with no children.
        doc->RemoveAllChildren();

        // This document now has no composite child nodes that we can add content to.
        // If we wish to edit it, we will need to repopulate its node collection.
        // First, create a new section, and then append it as a child to the root document node.
        auto section = MakeObject<Section>(doc);
        doc->AppendChild(section);

        // A section needs a body, which will contain and display all its contents
        // on the page between the section's header and footer.
        auto body = MakeObject<Body>(doc);
        section->AppendChild(body);

        // This body has no children, so we cannot add runs to it yet.
        ASSERT_EQ(0, doc->get_FirstSection()->get_Body()->GetChildNodes(NodeType::Any, true)->get_Count());

        // Call the "EnsureMinimum" to make sure that this body contains at least one empty paragraph.
        body->EnsureMinimum();

        // Now, we can add runs to the body, and get the document to display them.
        body->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        //ExEnd
    }

    void BodyChildNodes()
    {
        //ExStart
        //ExFor:Body.NodeType
        //ExFor:HeaderFooter.NodeType
        //ExFor:Document.FirstSection
        //ExSummary:Shows how to iterate through the children of a composite node.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Primary header");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"Primary footer");

        SharedPtr<Section> section = doc->get_FirstSection();

        // A Section is a composite node and can contain child nodes,
        // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
        for (auto node : System::IterateOver(section))
        {
            switch (node->get_NodeType())
            {
            case NodeType::Body: {
                auto body = System::DynamicCast<Body>(node);

                std::cout << "Body:" << std::endl;
                std::cout << "\t\"" << body->GetText().Trim() << "\"" << std::endl;
                break;
            }

            case NodeType::HeaderFooter: {
                auto headerFooter = System::DynamicCast<HeaderFooter>(node);

                std::cout << String::Format(u"HeaderFooter type: {0}:", headerFooter->get_HeaderFooterType()) << std::endl;
                std::cout << "\t\"" << headerFooter->GetText().Trim() << "\"" << std::endl;
                break;
            }

            default: {
                throw System::Exception(u"Unexpected node type in a section.");
            }
            }
        }
        //ExEnd
    }

    void Clear()
    {
        //ExStart
        //ExFor:NodeCollection.Clear
        //ExSummary:Shows how to remove all sections from a document.
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        // This document has one section with a few child nodes containing and displaying all the document's contents.
        ASSERT_EQ(1, doc->get_Sections()->get_Count());
        ASSERT_EQ(19, doc->get_Sections()->idx_get(0)->GetChildNodes(NodeType::Any, true)->get_Count());
        ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());

        // Clear the collection of sections, which will remove all of the document's children.
        doc->get_Sections()->Clear();

        ASSERT_EQ(0, doc->GetChildNodes(NodeType::Any, true)->get_Count());
        ASSERT_EQ(String::Empty, doc->GetText().Trim());
        //ExEnd
    }

    void PrependAppendContent()
    {
        //ExStart
        //ExFor:Section.AppendContent
        //ExFor:Section.PrependContent
        //ExSummary:Shows how to append the contents of a section to another section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 3");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(2);

        ASSERT_EQ(String(u"Section 3") + ControlChar::SectionBreak(), section->GetText());

        // Insert the contents of the first section to the beginning of the third section.
        SharedPtr<Section> sectionToPrepend = doc->get_Sections()->idx_get(0);
        section->PrependContent(sectionToPrepend);

        // Insert the contents of the second section to the end of the third section.
        SharedPtr<Section> sectionToAppend = doc->get_Sections()->idx_get(1);
        section->AppendContent(sectionToAppend);

        // The "PrependContent" and "AppendContent" methods did not create any new sections.
        ASSERT_EQ(3, doc->get_Sections()->get_Count());
        ASSERT_EQ(String(u"Section 1") + ControlChar::ParagraphBreak() + u"Section 3" + ControlChar::ParagraphBreak() + u"Section 2" +
                      ControlChar::SectionBreak(),
                  section->GetText());
        //ExEnd
    }

    void ClearContent()
    {
        //ExStart
        //ExFor:Section.ClearContent
        //ExSummary:Shows how to clear the contents of a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Hello world!");

        ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());

        // Running the "ClearContent" method will remove all the section contents
        // but leave a blank paragraph to add content again.
        doc->get_FirstSection()->ClearContent();

        ASSERT_EQ(String::Empty, doc->GetText().Trim());
        ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
        //ExEnd
    }

    void ClearHeadersFooters()
    {
        //ExStart
        //ExFor:Section.ClearHeadersFooters
        //ExSummary:Shows how to clear the contents of all headers and footers in a section.
        auto doc = MakeObject<Document>();

        ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->get_Count());

        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"This is the primary header.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Writeln(u"This is the primary footer.");

        ASSERT_EQ(2, doc->get_FirstSection()->get_HeadersFooters()->get_Count());

        ASSERT_EQ(u"This is the primary header.", doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->GetText().Trim());
        ASSERT_EQ(u"This is the primary footer.", doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->GetText().Trim());

        // Empty all the headers and footers in this section of all their contents.
        // The headers and footers themselves will still be present but will have nothing to display.
        doc->get_FirstSection()->ClearHeadersFooters();

        ASSERT_EQ(2, doc->get_FirstSection()->get_HeadersFooters()->get_Count());

        ASSERT_EQ(String::Empty, doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->GetText().Trim());
        ASSERT_EQ(String::Empty, doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->GetText().Trim());
        //ExEnd
    }

    void DeleteHeaderFooterShapes()
    {
        //ExStart
        //ExFor:Section.DeleteHeaderFooterShapes
        //ExSummary:Shows how to remove all shapes from all headers footers in a section.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a primary header with a shape.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->InsertShape(ShapeType::Rectangle, 100, 100);

        // Create a primary footer with an image.
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->InsertImage(ImageDir + u"Logo Icon.ico");

        ASSERT_EQ(1,
                  doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->GetChildNodes(NodeType::Shape, true)->get_Count());
        ASSERT_EQ(1,
                  doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->GetChildNodes(NodeType::Shape, true)->get_Count());

        // Remove all shapes from the headers and footers in the first section.
        doc->get_FirstSection()->DeleteHeaderFooterShapes();

        ASSERT_EQ(0,
                  doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->GetChildNodes(NodeType::Shape, true)->get_Count());
        ASSERT_EQ(0,
                  doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->GetChildNodes(NodeType::Shape, true)->get_Count());
        //ExEnd
    }

    void SectionsCloneSection()
    {
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
    }

    void SectionsImportSection()
    {
        auto srcDoc = MakeObject<Document>(MyDir + u"Document.docx");
        auto dstDoc = MakeObject<Document>();

        SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
        auto newSection = System::DynamicCast<Section>(dstDoc->ImportNode(sourceSection, true));
        dstDoc->get_Sections()->Add(newSection);
    }

    void MigrateFrom2XImportSection()
    {
        auto srcDoc = MakeObject<Document>();
        auto dstDoc = MakeObject<Document>();

        SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
        auto newSection = System::DynamicCast<Section>(dstDoc->ImportNode(sourceSection, true));
        dstDoc->get_Sections()->Add(newSection);
    }

    void ModifyPageSetupInAllSections()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->InsertBreak(BreakType::SectionBreakNewPage);
        builder->Write(u"Section 2");

        // It is important to understand that a document can contain many sections,
        // and each section has its page setup. In this case, we want to modify them all.
        for (auto section : System::IterateOver(doc->LINQ_OfType<SharedPtr<Section>>()))
        {
            section->get_PageSetup()->set_PaperSize(PaperSize::Letter);
        }

        doc->Save(ArtifactsDir + u"Section.ModifyPageSetupInAllSections.doc");
    }

    void CultureInfoPageSetupDefaults()
    {
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"en-us"));

        auto docEn = MakeObject<Document>();

        // Assert that page defaults comply with current culture info.
        SharedPtr<Section> sectionEn = docEn->get_Sections()->idx_get(0);
        ASPOSE_ASSERT_EQ(72.0, sectionEn->get_PageSetup()->get_LeftMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(72.0, sectionEn->get_PageSetup()->get_RightMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(72.0, sectionEn->get_PageSetup()->get_TopMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(72.0, sectionEn->get_PageSetup()->get_BottomMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(36.0, sectionEn->get_PageSetup()->get_HeaderDistance());
        // 1.27 cm
        ASPOSE_ASSERT_EQ(36.0, sectionEn->get_PageSetup()->get_FooterDistance());
        // 1.27 cm
        ASPOSE_ASSERT_EQ(36.0, sectionEn->get_PageSetup()->get_TextColumns()->get_Spacing());
        // 1.27 cm

        // Change the culture and assert that the page defaults are changed.
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"de-de"));

        auto docDe = MakeObject<Document>();

        SharedPtr<Section> sectionDe = docDe->get_Sections()->idx_get(0);
        ASPOSE_ASSERT_EQ(70.85, sectionDe->get_PageSetup()->get_LeftMargin());
        // 2.5 cm
        ASPOSE_ASSERT_EQ(70.85, sectionDe->get_PageSetup()->get_RightMargin());
        // 2.5 cm
        ASPOSE_ASSERT_EQ(70.85, sectionDe->get_PageSetup()->get_TopMargin());
        // 2.5 cm
        ASPOSE_ASSERT_EQ(56.7, sectionDe->get_PageSetup()->get_BottomMargin());
        // 2 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDe->get_PageSetup()->get_HeaderDistance());
        // 1.25 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDe->get_PageSetup()->get_FooterDistance());
        // 1.25 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDe->get_PageSetup()->get_TextColumns()->get_Spacing());
        // 1.25 cm

        // Change page defaults.
        sectionDe->get_PageSetup()->set_LeftMargin(90);
        // 3.17 cm
        sectionDe->get_PageSetup()->set_RightMargin(90);
        // 3.17 cm
        sectionDe->get_PageSetup()->set_TopMargin(72);
        // 2.54 cm
        sectionDe->get_PageSetup()->set_BottomMargin(72);
        // 2.54 cm
        sectionDe->get_PageSetup()->set_HeaderDistance(35.4);
        // 1.25 cm
        sectionDe->get_PageSetup()->set_FooterDistance(35.4);
        // 1.25 cm
        sectionDe->get_PageSetup()->get_TextColumns()->set_Spacing(35.4);
        // 1.25 cm

        docDe = DocumentHelper::SaveOpen(docDe);

        SharedPtr<Section> sectionDeAfter = docDe->get_Sections()->idx_get(0);
        ASPOSE_ASSERT_EQ(90.0, sectionDeAfter->get_PageSetup()->get_LeftMargin());
        // 3.17 cm
        ASPOSE_ASSERT_EQ(90.0, sectionDeAfter->get_PageSetup()->get_RightMargin());
        // 3.17 cm
        ASPOSE_ASSERT_EQ(72.0, sectionDeAfter->get_PageSetup()->get_TopMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(72.0, sectionDeAfter->get_PageSetup()->get_BottomMargin());
        // 2.54 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDeAfter->get_PageSetup()->get_HeaderDistance());
        // 1.25 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDeAfter->get_PageSetup()->get_FooterDistance());
        // 1.25 cm
        ASPOSE_ASSERT_EQ(35.4, sectionDeAfter->get_PageSetup()->get_TextColumns()->get_Spacing());
        // 1.25 cm
    }
};

} // namespace ApiExamples
