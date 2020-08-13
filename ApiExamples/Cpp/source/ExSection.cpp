// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExSection.h"

#include <testing/test_predicates.h>
#include <system/threading/thread.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/globalization/culture_info.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/console.h>
#include <system/collections/ienumerable.h>
#include <gtest/gtest.h>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
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
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Document/ProtectionType.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>

#include "DocumentHelper.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
namespace ApiExamples {

namespace gtest_test
{

class ExSection : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExSection> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExSection>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExSection> ExSection::s_instance;

} // namespace gtest_test

void ExSection::Protect()
{
    //ExStart
    //ExFor:Document.Protect(ProtectionType)
    //ExFor:ProtectionType
    //ExFor:Section.ProtectedForForms
    //ExSummary:Shows how to protect a section so only editing in form fields is possible.
    auto doc = MakeObject<Document>();

    // Insert two sections with some text
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->Writeln(u"Section 1. Unprotected.");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakContinuous);
    builder->Writeln(u"Section 2. Protected.");

    // Section protection only works when document protection is turned and only editing in form fields is allowed
    doc->Protect(Aspose::Words::ProtectionType::AllowOnlyFormFields);

    // By default, all sections are protected, but we can selectively turn protection off
    doc->get_Sections()->idx_get(0)->set_ProtectedForForms(false);

    doc->Save(ArtifactsDir + u"Section.Protect.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"Section.Protect.docx");

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
    //ExSummary:Shows how to add/remove sections in a document.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");

    // This shows what is in the document originally. The document has two sections
    ASSERT_EQ(u"Section 1\x000c" u"Section 2", doc->GetText().Trim());

    // Delete the first section from the document
    doc->get_Sections()->RemoveAt(0);

    // Duplicate the last section and append the copy to the end of the document
    int lastSectionIdx = doc->get_Sections()->get_Count() - 1;
    SharedPtr<Section> newSection = doc->get_Sections()->idx_get(lastSectionIdx)->Clone();
    doc->get_Sections()->Add(newSection);

    // Check what the document contains after we changed it
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

void ExSection::CreateFromScratch()
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
    //ExSummary:Shows how to construct an Aspose Words document node by node.
    auto doc = MakeObject<Document>();

    // A newly created blank document still comes one section, one body and one paragraph
    // Calling this method will remove all those nodes to completely empty the document
    doc->RemoveAllChildren();

    // This document now has no composite nodes that content can be added to
    // If we wish to edit it, we will need to repopulate its node collection,
    // which we will start to do with by creating a new Section node
    auto section = MakeObject<Section>(doc);

    // Append the section to the document
    doc->AppendChild(section);

    // Lets set some properties for the section
    section->get_PageSetup()->set_SectionStart(Aspose::Words::SectionStart::NewPage);
    section->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Letter);

    // The section that we created is empty, lets populate it. The section needs at least the Body node
    auto body = MakeObject<Body>(doc);
    section->AppendChild(body);

    // The body needs to have at least one paragraph
    // Note that the paragraph has not yet been added to the document,
    // but we have to specify the parent document
    // The parent document is needed so the paragraph can correctly work
    // with styles and other document-wide information
    auto para = MakeObject<Paragraph>(doc);
    body->AppendChild(para);

    // We can set some formatting for the paragraph
    para->get_ParagraphFormat()->set_StyleName(u"Heading 1");
    para->get_ParagraphFormat()->set_Alignment(Aspose::Words::ParagraphAlignment::Center);

    // So far we have one empty paragraph in the document
    // The document is valid and can be saved, but lets add some text before saving
    // Create a new run of text and add it to our paragraph
    auto run = MakeObject<Run>(doc);
    run->set_Text(u"Hello World!");
    run->get_Font()->set_Color(System::Drawing::Color::get_Red());
    para->AppendChild(run);

    ASSERT_EQ(String(u"Hello World!") + ControlChar::SectionBreakChar, doc->GetText());

    doc->Save(ArtifactsDir + u"Section.CreateFromScratch.docx");
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, CreateFromScratch)
{
    s_instance->CreateFromScratch();
}

} // namespace gtest_test

void ExSection::EnsureSectionMinimum()
{
    //ExStart
    //ExFor:NodeCollection.Add
    //ExFor:Section.EnsureMinimum
    //ExFor:SectionCollection.Item(Int32)
    //ExSummary:Shows how to prepare a new section node for editing.
    auto doc = MakeObject<Document>();

    // A blank document comes with a section, which has a body, which in turn has a paragraph,
    // so we can edit the document by adding children to the paragraph like shapes or runs of text
    ASSERT_EQ(2, doc->get_Sections()->idx_get(0)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // If we add a new section like this, it will not have a body or a paragraph that we can edit
    doc->get_Sections()->Add(MakeObject<Section>(doc));

    ASSERT_EQ(0, doc->get_Sections()->idx_get(1)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    // Makes sure that the section contains a body with at least one paragraph
    doc->get_LastSection()->EnsureMinimum();

    // Now we can add content to this section
    ASSERT_EQ(2, doc->get_Sections()->idx_get(1)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, EnsureSectionMinimum)
{
    s_instance->EnsureSectionMinimum();
}

} // namespace gtest_test

void ExSection::BodyEnsureMinimum()
{
    //ExStart
    //ExFor:Section.Body
    //ExFor:Body.EnsureMinimum
    //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
    // Open a document
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");

    // This shows what is in the document originally
    // The document has two sections
    ASSERT_EQ(String::Format(u"Section 1{0}Section 2{1}",ControlChar::SectionBreak(),ControlChar::SectionBreak()), doc->GetText());

    // Loop through all sections in the document
    for (auto section : System::IterateOver(doc->get_Sections()->LINQ_OfType<SharedPtr<Section> >()))
    {
        // Each section has a Body node that contains main story (main text) of the section
        SharedPtr<Body> body = section->get_Body();

        // This clears all nodes from the body
        body->RemoveAllChildren();

        // Technically speaking, for the main story of a section to be valid, it needs to have
        // at least one empty paragraph. That's what the EnsureMinimum method does
        body->EnsureMinimum();
    }

    // Check how the content of the document looks now
    ASSERT_EQ(String::Format(u"{0}{1}",ControlChar::SectionBreak(),ControlChar::SectionBreak()), doc->GetText());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, BodyEnsureMinimum)
{
    s_instance->BodyEnsureMinimum();
}

} // namespace gtest_test

void ExSection::BodyNodeType()
{
    //ExStart
    //ExFor:Body.NodeType
    //ExFor:HeaderFooter.NodeType
    //ExFor:Document.FirstSection
    //ExSummary:Shows how you can enumerate through children of a composite node and detect types of the children nodes.
    // Open a document
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Section 1");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::HeaderPrimary);
    builder->Write(u"Primary header");
    builder->MoveToHeaderFooter(Aspose::Words::HeaderFooterType::FooterPrimary);
    builder->Write(u"Primary footer");

    // Get the first section in the document
    SharedPtr<Section> section = doc->get_FirstSection();

    // A Section is a composite node and therefore can contain child nodes
    // Section can contain only Body and HeaderFooter nodes
    for (auto node : System::IterateOver(section))
    {
        // Every node has the NodeType property
        switch (node->get_NodeType())
        {
            case Aspose::Words::NodeType::Body:
                {
                    // If the node type is Body, we can cast the node to the Body class
                    auto body = System::DynamicCast<Aspose::Words::Body>(node);

                    // Write the content of the main story of the section to the console
                    System::Console::WriteLine(u"*** Body ***");
                    System::Console::WriteLine(body->GetText());
                    break;
                }

            case Aspose::Words::NodeType::HeaderFooter:
                {
                    // If the node type is HeaderFooter, we can cast the node to the HeaderFooter class
                    auto headerFooter = System::DynamicCast<Aspose::Words::HeaderFooter>(node);

                    // Write the content of the header footer to the console
                    System::Console::WriteLine(u"*** HeaderFooter ***");
                    System::Console::WriteLine(System::ObjectExt::Box<HeaderFooterType>(headerFooter->get_HeaderFooterType()));
                    System::Console::WriteLine(headerFooter->GetText());
                    break;
                }

            default:
                {
                    // Other types of nodes never occur inside a Section node
                    throw System::Exception(u"Unexpected node type in a section.");
                }

        }
    }
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, BodyNodeType)
{
    s_instance->BodyNodeType();
}

} // namespace gtest_test

void ExSection::SectionsDeleteAllSections()
{
    //ExStart
    //ExFor:NodeCollection.Clear
    //ExSummary:Shows how to remove all sections from a document.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    // All of the document's content is stored in the child nodes of sections like this one
    ASSERT_EQ(u"Hello World!", doc->GetText().Trim());
    ASSERT_EQ(5, doc->get_Sections()->idx_get(0)->GetChildNodes(Aspose::Words::NodeType::Any, true)->get_Count());

    doc->get_Sections()->Clear();

    // Clearing the section collection effectively empties the document
    ASSERT_EQ(String::Empty, doc->GetText());
    ASSERT_EQ(0, doc->get_Sections()->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, SectionsDeleteAllSections)
{
    s_instance->SectionsDeleteAllSections();
}

} // namespace gtest_test

void ExSection::SectionsAppendSectionContent()
{
    //ExStart
    //ExFor:Section.AppendContent
    //ExFor:Section.PrependContent
    //ExSummary:Shows how to append content of an existing section. The number of sections in the document remains the same.
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 3");

    // This is the section that we will append and prepend to
    SharedPtr<Section> section = doc->get_Sections()->idx_get(2);

    // This copies content of the 1st section and inserts it at the beginning of the specified section
    SharedPtr<Section> sectionToPrepend = doc->get_Sections()->idx_get(0);
    section->PrependContent(sectionToPrepend);

    // This copies content of the 2nd section and inserts it at the end of the specified section
    SharedPtr<Section> sectionToAppend = doc->get_Sections()->idx_get(1);
    section->AppendContent(sectionToAppend);

    ASSERT_EQ(String(u"Section 1") + ControlChar::SectionBreak() + u"Section 2" + ControlChar::SectionBreak() + u"Section 1" + ControlChar::ParagraphBreak() + u"Section 3" + ControlChar::ParagraphBreak() + u"Section 2" + ControlChar::SectionBreak(), doc->GetText());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, SectionsAppendSectionContent)
{
    s_instance->SectionsAppendSectionContent();
}

} // namespace gtest_test

void ExSection::SectionsDeleteSectionContent()
{
    //ExStart
    //ExFor:Section.ClearContent
    //ExSummary:Shows how to clear the content of a section.
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");

    ASSERT_EQ(u"Hello World!", doc->GetText().Trim());

    doc->get_FirstSection()->ClearContent();

    ASSERT_EQ(String::Empty, doc->GetText().Trim());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, SectionsDeleteSectionContent)
{
    s_instance->SectionsDeleteSectionContent();
}

} // namespace gtest_test

void ExSection::SectionsDeleteHeaderFooter()
{
    //ExStart
    //ExFor:Section.ClearHeadersFooters
    //ExSummary:Clears content of all headers and footers in a section.
    auto doc = MakeObject<Document>(MyDir + u"Header and footer types.docx");

    SharedPtr<Section> section = doc->get_Sections()->idx_get(0);

    ASSERT_EQ(6, section->get_HeadersFooters()->get_Count());
    ASSERT_EQ(u"First header", section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText().Trim());

    section->ClearHeadersFooters();

    ASSERT_EQ(6, section->get_HeadersFooters()->get_Count());
    ASSERT_EQ(String::Empty, section->get_HeadersFooters()->idx_get(Aspose::Words::HeaderFooterType::HeaderFirst)->GetText());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, SectionsDeleteHeaderFooter)
{
    s_instance->SectionsDeleteHeaderFooter();
}

} // namespace gtest_test

void ExSection::SectionDeleteHeaderFooterShapes()
{
    //ExStart
    //ExFor:Section.DeleteHeaderFooterShapes
    //ExSummary:Removes all images and shapes from all headers footers in a section.
    auto doc = MakeObject<Document>();
    SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
    auto firstHeader = MakeObject<HeaderFooter>(doc, Aspose::Words::HeaderFooterType::HeaderFirst);

    section->get_HeadersFooters()->Add(firstHeader);

    firstHeader->AppendParagraph(u"This paragraph contains a shape: ");

    auto shape = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Arrow);
    firstHeader->get_FirstParagraph()->AppendChild(shape);

    ASSERT_EQ(1, firstHeader->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    section->DeleteHeaderFooterShapes();

    ASSERT_EQ(0, firstHeader->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExSection, SectionDeleteHeaderFooterShapes)
{
    s_instance->SectionDeleteHeaderFooterShapes();
}

} // namespace gtest_test

void ExSection::SectionsCloneSection()
{
    auto doc = MakeObject<Document>(MyDir + u"Document.docx");
    SharedPtr<Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
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
    auto srcDoc = MakeObject<Document>(MyDir + u"Document.docx");
    auto dstDoc = MakeObject<Document>();

    SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
    auto newSection = System::DynamicCast<Aspose::Words::Section>(dstDoc->ImportNode(sourceSection, true));
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
    auto srcDoc = MakeObject<Document>();
    auto dstDoc = MakeObject<Document>();

    SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
    auto newSection = System::DynamicCast<Aspose::Words::Section>(dstDoc->ImportNode(sourceSection, true));
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
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    builder->Write(u"Section 1");
    builder->InsertBreak(Aspose::Words::BreakType::SectionBreakNewPage);
    builder->Write(u"Section 2");

    // It is important to understand that a document can contain many sections and each
    // section has its own page setup. In this case we want to modify them all
    for (auto section : System::IterateOver(doc->LINQ_OfType<SharedPtr<Section> >()))
    {
        section->get_PageSetup()->set_PaperSize(Aspose::Words::PaperSize::Letter);
    }

    doc->Save(ArtifactsDir + u"Section.ModifyPageSetupInAllSections.doc");
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
    System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"en-us"));

    auto docEn = MakeObject<Document>();

    // Assert that page defaults comply current culture info
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

    // Change culture and assert that the page defaults are changed
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

    // Change page defaults
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

namespace gtest_test
{

TEST_F(ExSection, CultureInfoPageSetupDefaults)
{
    s_instance->CultureInfoPageSetupDefaults();
}

} // namespace gtest_test

} // namespace ApiExamples
