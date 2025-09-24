// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSection.h"

#include <testing/test_predicates.h>
#include <system/threading/thread.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/globalization/culture_info.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <iostream>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ControlChar.h>
#include <Aspose.Words.Cpp/Model/Sections/TextColumnCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionStart.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/PaperSize.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/Model/Sections/HeaderFooter.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/TextFormFieldType.h>
#include <Aspose.Words.Cpp/Model/Fields/FormFields/FormField.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/WatermarkType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Watermark/Watermark.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "DocumentHelper.h"


using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(2885541077u, ::Aspose::Words::ApiExamples::ExSection, ThisTypeBaseTypesInfo);

namespace gtest_test
{

class ExSection : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExSection> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExSection>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExSection> ExSection::s_instance;

} // namespace gtest_test

void ExSection::Protect()
{
    //ExStart
    //ExFor:Document.Protect(ProtectionType)
    //ExFor:ProtectionType
    //ExFor:Section.ProtectedForForms
    //ExSummary:Shows how to turn off protection for a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Section 1. Hello world!");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    
    builder->Writeln(u"Section 2. Hello again!");
    builder->Write(u"Please enter text here: ");
    builder->InsertTextInput(u"TextInput1", Aspose::Words::Fields::TextFormFieldType::Regular, u"", u"Placeholder text", 0);
    
    // Apply write protection to every section in the document.
    doc->Protect(Aspose::Words::ProtectionType::AllowOnlyFormFields);
    
    // Turn off write protection for the first section.
    doc->get_Sections()->idx_get(0)->set_ProtectedForForms(false);
    
    // In this output document, we will be able to edit the first section freely,
    // and we will only be able to edit the contents of the form field in the second section.
    doc->Save(get_ArtifactsDir() + u"Section.Protect.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Section.Protect.docx");
    
    ASSERT_FALSE(doc->get_Sections()->idx_get(0)->get_ProtectedForForms());
    ASSERT_TRUE(doc->get_Sections()->idx_get(1)->get_ProtectedForForms());
}

namespace gtest_test
{

TEST_F(ExSection, Protect)
{
    s_instance->Protect();
}

} // namespace gtest_test

void ExSection::AddRemove()
{
    //ExStart
    //ExFor:Document.Sections
    //ExFor:Section.Clone
    //ExFor:SectionCollection
    //ExFor:NodeCollection.RemoveAt(Int32)
    //ExSummary:Shows how to add and remove sections in a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    
    ASSERT_EQ(u"Section 1\x000c" u"Section 2", doc->GetText().Trim());
    
    // Delete the first section from the document.
    doc->get_Sections()->RemoveAt(0);
    
    ASSERT_EQ(u"Section 2", doc->GetText().Trim());
    
    // Append a copy of what is now the first section to the end of the document.
    int32_t lastSectionIdx = doc->get_Sections()->get_Count() - 1;
    System::SharedPtr<Aspose::Words::Section> newSection = doc->get_Sections()->idx_get(lastSectionIdx)->Clone();
    doc->get_Sections()->Add(newSection);
    
    ASSERT_EQ(u"Section 2\x000c" u"Section 2", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, AddRemove)
{
    s_instance->AddRemove();
}

} // namespace gtest_test

void ExSection::FirstAndLast()
{
    //ExStart
    //ExFor:Document.FirstSection
    //ExFor:Document.LastSection
    //ExSummary:Shows how to create a new section with a document builder.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A blank document contains one section by default,
    // which contains child nodes that we can edit.
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    
    // Use a document builder to add text to the first section.
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    // Create a second section by inserting a section break.
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    
    ASSERT_EQ(2, doc->get_Sections()->get_Count());
    
    // Each section has its own page setup settings.
    // We can split the text in the second section into two columns.
    // This will not affect the text in the first section.
    doc->get_LastSection()->get_PageSetup()->get_TextColumns()->SetCount(2);
    builder->Writeln(u"Column 1.");
    builder->InsertBreak(Aspose::Words::BreakType::ColumnBreak);
    builder->Writeln(u"Column 2.");
    
    ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
    ASSERT_EQ(2, doc->get_LastSection()->get_PageSetup()->get_TextColumns()->get_Count());
    
    doc->Save(get_ArtifactsDir() + u"Section.Create.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"Section.Create.docx");
    
    ASSERT_EQ(1, doc->get_FirstSection()->get_PageSetup()->get_TextColumns()->get_Count());
    ASSERT_EQ(2, doc->get_LastSection()->get_PageSetup()->get_TextColumns()->get_Count());
}

namespace gtest_test
{

TEST_F(ExSection, FirstAndLast)
{
    s_instance->FirstAndLast();
}

} // namespace gtest_test

void ExSection::CreateManually()
{
    //ExStart
    //ExFor:Node.GetText
    //ExFor:CompositeNode.RemoveAllChildren
    //ExFor:CompositeNode.AppendChild``1(``0)
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
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A blank document contains one section, one body and one paragraph.
    // Call the "RemoveAllChildren" method to remove all those nodes,
    // and end up with a document node with no children.
    doc->RemoveAllChildren();
    
    // This document now has no composite child nodes that we can add content to.
    // If we wish to edit it, we will need to repopulate its node collection.
    // First, create a new section, and then append it as a child to the root document node.
    auto section = System::MakeObject<Aspose::Words::Section>(doc);
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(section);
    
    // Set some page setup properties for the section.
    section->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::NewPage);
    section->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Letter);
    
    // A section needs a body, which will contain and display all its contents
    // on the page between the section's header and footer.
    auto body = System::MakeObject<Aspose::Words::Body>(doc);
    section->AppendChild<System::SharedPtr<Aspose::Words::Body>>(body);
    
    // Create a paragraph, set some formatting properties, and then append it as a child to the body.
    auto para = System::MakeObject<Aspose::Words::Paragraph>(doc);
    
    para->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    para->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);
    
    body->AppendChild<System::SharedPtr<Aspose::Words::Paragraph>>(para);
    
    // Finally, add some content to do the document. Create a run,
    // set its appearance and contents, and then append it as a child to the paragraph.
    auto run = System::MakeObject<Aspose::Words::Run>(doc);
    run->set_Text(u"Hello World!");
    run->get_Font()->set_Color(System::Drawing::Color::get_Red());
    para->AppendChild<System::SharedPtr<Aspose::Words::Run>>(run);
    
    ASSERT_EQ(u"Hello World!", doc->GetText().Trim());
    
    doc->Save(get_ArtifactsDir() + u"Section.CreateManually.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, CreateManually)
{
    s_instance->CreateManually();
}

} // namespace gtest_test

void ExSection::EnsureMinimum()
{
    //ExStart
    //ExFor:NodeCollection.Add
    //ExFor:Section.EnsureMinimum
    //ExFor:SectionCollection.Item(Int32)
    //ExSummary:Shows how to prepare a new section node for editing.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A blank document comes with a section, which has a body, which in turn has a paragraph.
    // We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
    ASSERT_EQ(Aspose::Words::NodeType::Section, doc->GetChild(Aspose::Words::NodeType::Any, 0, true)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Body, doc->get_Sections()->idx_get(0)->GetChild(Aspose::Words::NodeType::Any, 0, true)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, doc->get_Sections()->idx_get(0)->get_Body()->GetChild(Aspose::Words::NodeType::Any, 0, true)->get_NodeType());
    
    // If we add a new section like this, it will not have a body, or any other child nodes.
    doc->get_Sections()->Add(System::MakeObject<Aspose::Words::Section>(doc));
    
    ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
    doc->get_LastSection()->EnsureMinimum();
    
    ASSERT_EQ(Aspose::Words::NodeType::Body, doc->get_Sections()->idx_get(1)->GetChild(Aspose::Words::NodeType::Any, 0, true)->get_NodeType());
    ASSERT_EQ(Aspose::Words::NodeType::Paragraph, doc->get_Sections()->idx_get(1)->get_Body()->GetChild(Aspose::Words::NodeType::Any, 0, true)->get_NodeType());
    
    doc->get_Sections()->idx_get(0)->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, EnsureMinimum)
{
    s_instance->EnsureMinimum();
}

} // namespace gtest_test

void ExSection::BodyEnsureMinimum()
{
    //ExStart
    //ExFor:Section.Body
    //ExFor:Body.EnsureMinimum
    //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    // A blank document contains one section, one body and one paragraph.
    // Call the "RemoveAllChildren" method to remove all those nodes,
    // and end up with a document node with no children.
    doc->RemoveAllChildren();
    
    // This document now has no composite child nodes that we can add content to.
    // If we wish to edit it, we will need to repopulate its node collection.
    // First, create a new section, and then append it as a child to the root document node.
    auto section = System::MakeObject<Aspose::Words::Section>(doc);
    doc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(section);
    
    // A section needs a body, which will contain and display all its contents
    // on the page between the section's header and footer.
    auto body = System::MakeObject<Aspose::Words::Body>(doc);
    section->AppendChild<System::SharedPtr<Aspose::Words::Body>>(body);
    
    // This body has no children, so we cannot add runs to it yet.
    ASSERT_EQ(0, doc->get_FirstSection()->get_Body()->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    
    // Call the "EnsureMinimum" to make sure that this body contains at least one empty paragraph. 
    body->EnsureMinimum();
    
    // Now, we can add runs to the body, and get the document to display them.
    body->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(doc, u"Hello world!"));
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, BodyEnsureMinimum)
{
    s_instance->BodyEnsureMinimum();
}

} // namespace gtest_test

void ExSection::BodyChildNodes()
{
    //ExStart
    //ExFor:Body.NodeType
    //ExFor:HeaderFooter.NodeType
    //ExFor:Document.FirstSection
    //ExSummary:Shows how to iterate through the children of a composite node.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Primary header");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Write(u"Primary footer");
    
    System::SharedPtr<Aspose::Words::Section> section = doc->get_FirstSection();
    
    // A Section is a composite node and can contain child nodes,
    // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
    for (auto&& node : System::IterateOver(section))
    {
        switch (node->get_NodeType())
        {
            case Aspose::Words::NodeType::Body:
                {
                    auto body = System::ExplicitCast<Aspose::Words::Body>(node);
                    
                    std::cout << "Body:" << std::endl;
                    std::cout << System::String::Format(u"\t\"{0}\"", body->GetText().Trim()) << std::endl;
                    break;
                }
            
            case Aspose::Words::NodeType::HeaderFooter:
                {
                    auto headerFooter = System::ExplicitCast<Aspose::Words::HeaderFooter>(node);
                    
                    std::cout << System::String::Format(u"HeaderFooter type: {0}:", headerFooter->get_HeaderFooterType()) << std::endl;
                    std::cout << System::String::Format(u"\t\"{0}\"", headerFooter->GetText().Trim()) << std::endl;
                    break;
                }
            
            default: 
                {
                    throw System::Exception(u"Unexpected node type in a section.");
                }
            
        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, BodyChildNodes)
{
    s_instance->BodyChildNodes();
}

} // namespace gtest_test

void ExSection::Clear()
{
    //ExStart
    //ExFor:NodeCollection.Clear
    //ExSummary:Shows how to remove all sections from a document.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    
    // This document has one section with a few child nodes containing and displaying all the document's contents.
    ASSERT_EQ(1, doc->get_Sections()->get_Count());
    ASSERT_EQ(17, doc->get_Sections()->idx_get(0)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    ASSERT_EQ(u"Hello World!\r\rHello Word!\r\r\rHello World!", doc->GetText().Trim());
    
    // Clear the collection of sections, which will remove all of the document's children.
    doc->get_Sections()->Clear();
    
    ASSERT_EQ(0, doc->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    ASSERT_EQ(System::String::Empty, doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, Clear)
{
    s_instance->Clear();
}

} // namespace gtest_test

void ExSection::PrependAppendContent()
{
    //ExStart
    //ExFor:Section.AppendContent
    //ExFor:Section.PrependContent
    //ExSummary:Shows how to append the contents of a section to another section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 3");
    
    System::SharedPtr<Aspose::Words::Section> section = doc->get_Sections()->idx_get(2);
    
    ASSERT_EQ(System::String(u"Section 3") + Aspose::Words::ControlChar::SectionBreak(), section->GetText());
    
    // Insert the contents of the first section to the beginning of the third section.
    System::SharedPtr<Aspose::Words::Section> sectionToPrepend = doc->get_Sections()->idx_get(0);
    section->PrependContent(sectionToPrepend);
    
    // Insert the contents of the second section to the end of the third section.
    System::SharedPtr<Aspose::Words::Section> sectionToAppend = doc->get_Sections()->idx_get(1);
    section->AppendContent(sectionToAppend);
    
    // The "PrependContent" and "AppendContent" methods did not create any new sections.
    ASSERT_EQ(3, doc->get_Sections()->get_Count());
    ASSERT_EQ(System::String(u"Section 1") + Aspose::Words::ControlChar::ParagraphBreak() + u"Section 3" + Aspose::Words::ControlChar::ParagraphBreak() + u"Section 2" + Aspose::Words::ControlChar::SectionBreak(), section->GetText());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, PrependAppendContent)
{
    s_instance->PrependAppendContent();
}

} // namespace gtest_test

void ExSection::ClearContent()
{
    //ExStart
    //ExFor:Section.ClearContent
    //ExSummary:Shows how to clear the contents of a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Hello world!");
    
    ASSERT_EQ(u"Hello world!", doc->GetText().Trim());
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
    
    // Running the "ClearContent" method will remove all the section contents
    // but leave a blank paragraph to add content again.
    doc->get_FirstSection()->ClearContent();
    
    ASSERT_EQ(System::String::Empty, doc->GetText().Trim());
    ASSERT_EQ(1, doc->get_FirstSection()->get_Body()->get_Paragraphs()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, ClearContent)
{
    s_instance->ClearContent();
}

} // namespace gtest_test

void ExSection::ClearHeadersFooters()
{
    //ExStart
    //ExFor:Section.ClearHeadersFooters
    //ExSummary:Shows how to clear the contents of all headers and footers in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->get_Count());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Writeln(u"This is the primary header.");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Writeln(u"This is the primary footer.");
    
    ASSERT_EQ(2, doc->get_FirstSection()->get_HeadersFooters()->get_Count());
    
    ASSERT_EQ(u"This is the primary header.", doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetText().Trim());
    ASSERT_EQ(u"This is the primary footer.", doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Trim());
    
    // Empty all the headers and footers in this section of all their contents.
    // The headers and footers themselves will still be present but will have nothing to display.
    doc->get_FirstSection()->ClearHeadersFooters();
    
    ASSERT_EQ(2, doc->get_FirstSection()->get_HeadersFooters()->get_Count());
    
    ASSERT_EQ(System::String::Empty, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetText().Trim());
    ASSERT_EQ(System::String::Empty, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, ClearHeadersFooters)
{
    s_instance->ClearHeadersFooters();
}

} // namespace gtest_test

void ExSection::DeleteHeaderFooterShapes()
{
    //ExStart
    //ExFor:Section.DeleteHeaderFooterShapes
    //ExSummary:Shows how to remove all shapes from all headers footers in a section.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Create a primary header with a shape.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->InsertShape(Aspose::Words::Drawing::ShapeType::Rectangle, 100, 100);
    
    // Create a primary footer with an image.
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->InsertImage(get_ImageDir() + u"Logo icon.ico");
    
    ASSERT_EQ(1, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_EQ(1, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    
    // Remove all shapes from the headers and footers in the first section.
    doc->get_FirstSection()->DeleteHeaderFooterShapes();
    
    ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    ASSERT_EQ(0, doc->get_FirstSection()->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, DeleteHeaderFooterShapes)
{
    s_instance->DeleteHeaderFooterShapes();
}

} // namespace gtest_test

void ExSection::SectionsCloneSection()
{
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    System::SharedPtr<Aspose::Words::Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
}

namespace gtest_test
{

TEST_F(ExSection, SectionsCloneSection)
{
    s_instance->SectionsCloneSection();
}

} // namespace gtest_test

void ExSection::SectionsImportSection()
{
    auto srcDoc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Document.docx");
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
    auto newSection = System::ExplicitCast<Aspose::Words::Section>(dstDoc->ImportNode(sourceSection, true));
    dstDoc->get_Sections()->Add(newSection);
}

namespace gtest_test
{

TEST_F(ExSection, SectionsImportSection)
{
    s_instance->SectionsImportSection();
}

} // namespace gtest_test

void ExSection::MigrateFrom2XImportSection()
{
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
    auto newSection = System::ExplicitCast<Aspose::Words::Section>(dstDoc->ImportNode(sourceSection, true));
    dstDoc->get_Sections()->Add(newSection);
}

namespace gtest_test
{

TEST_F(ExSection, MigrateFrom2XImportSection)
{
    s_instance->MigrateFrom2XImportSection();
}

} // namespace gtest_test

void ExSection::ModifyPageSetupInAllSections()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    
    // It is important to understand that a document can contain many sections,
    // and each section has its page setup. In this case, we want to modify them all.
    for (auto&& section : System::IterateOver<Aspose::Words::Section>(doc->GetChildNodes(Aspose::Words::NodeType::Section, true)))
    {
        section->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Letter);
    }
    
    doc->Save(get_ArtifactsDir() + u"Section.ModifyPageSetupInAllSections.doc");
}

namespace gtest_test
{

TEST_F(ExSection, ModifyPageSetupInAllSections)
{
    s_instance->ModifyPageSetupInAllSections();
}

} // namespace gtest_test

void ExSection::CultureInfoPageSetupDefaults()
{
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::MakeObject<System::Globalization::CultureInfo>(u"en-us"));
    
    auto docEn = System::MakeObject<Aspose::Words::Document>();
    
    // Assert that page defaults comply with current culture info.
    System::SharedPtr<Aspose::Words::Section> sectionEn = docEn->get_Sections()->idx_get(0);
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
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(System::MakeObject<System::Globalization::CultureInfo>(u"de-de"));
    
    auto docDe = System::MakeObject<Aspose::Words::Document>();
    
    System::SharedPtr<Aspose::Words::Section> sectionDe = docDe->get_Sections()->idx_get(0);
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
    
    docDe = Aspose::Words::ApiExamples::DocumentHelper::SaveOpen(docDe);
    
    System::SharedPtr<Aspose::Words::Section> sectionDeAfter = docDe->get_Sections()->idx_get(0);
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

namespace gtest_test
{

TEST_F(ExSection, CultureInfoPageSetupDefaults)
{
    s_instance->CultureInfoPageSetupDefaults();
}

} // namespace gtest_test

void ExSection::PreserveWatermarks()
{
    //ExStart:PreserveWatermarks
    //GistId:708ce40a68fac5003d46f6b4acfd5ff1
    //ExFor:Section.ClearHeadersFooters(bool)
    //ExSummary:Shows how to clear the contents of header and footer with or without a watermark.
    auto doc = System::MakeObject<Aspose::Words::Document>(get_MyDir() + u"Header and footer types.docx");
    
    // Add a plain text watermark.
    doc->get_Watermark()->SetText(u"Aspose Watermark");
    
    // Make sure the headers and footers have content.
    System::SharedPtr<Aspose::Words::HeaderFooterCollection> headersFooters = doc->get_FirstSection()->get_HeadersFooters();
    ASSERT_EQ(u"First header", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText().Trim());
    ASSERT_EQ(u"Second header", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderEven)->GetText().Trim());
    ASSERT_EQ(u"Third header", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetText().Trim());
    ASSERT_EQ(u"First footer", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterFirst)->GetText().Trim());
    ASSERT_EQ(u"Second footer", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterEven)->GetText().Trim());
    ASSERT_EQ(u"Third footer", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Trim());
    
    // Removes all header and footer content except watermarks.
    doc->get_FirstSection()->ClearHeadersFooters(true);
    
    headersFooters = doc->get_FirstSection()->get_HeadersFooters();
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText().Trim());
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderEven)->GetText().Trim());
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::HeaderPrimary)->GetText().Trim());
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterFirst)->GetText().Trim());
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterEven)->GetText().Trim());
    ASSERT_EQ(u"", headersFooters->idx_get(Aspose::Words::HeaderFooterType::FooterPrimary)->GetText().Trim());
    ASSERT_EQ(Aspose::Words::WatermarkType::Text, doc->get_Watermark()->get_Type());
    
    // Removes all header and footer content including watermarks.
    doc->get_FirstSection()->ClearHeadersFooters(false);
    ASSERT_EQ(Aspose::Words::WatermarkType::None, doc->get_Watermark()->get_Type());
    //ExEnd:PreserveWatermarks
}

namespace gtest_test
{

TEST_F(ExSection, PreserveWatermarks)
{
    s_instance->PreserveWatermarks();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
