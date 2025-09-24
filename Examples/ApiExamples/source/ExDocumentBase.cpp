// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBase.h"

#include <testing/test_predicates.h>
#include <system/type_info.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/object_ext.h>
#include <system/io/file.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/color.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/RunCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Run.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleCollection.h>
#include <Aspose.Words.Cpp/Model/Styles/Style.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/Body.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/PdfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>


using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Loading;
namespace Aspose {

namespace Words {

namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3034527200u, ::Aspose::Words::ApiExamples::ExDocumentBase::ImageNameHandler, ThisTypeBaseTypesInfo);

Aspose::Words::Loading::ResourceLoadingAction ExDocumentBase::ImageNameHandler::ResourceLoading(System::SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args)
{
    // If this callback encounters one of the image shorthands while loading an image,
    // it will apply unique logic for each defined shorthand instead of treating it as a URI.
    if (args->get_ResourceType() == Aspose::Words::Loading::ResourceType::Image)
    {
        const System::String& switch_value_0 = args->get_OriginalUri();
        if (switch_value_0 == u"Aspose logo")
        {
            args->SetData(System::IO::File::ReadAllBytes(get_ImageDir() + u"Logo.jpg"));
            return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
        }
        else if (switch_value_0 == u"Watermark")
        {
            args->SetData(System::IO::File::ReadAllBytes(get_ImageDir() + u"Transparent background logo.png"));
            return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
        }
    }
    
    return Aspose::Words::Loading::ResourceLoadingAction::Default;
}


RTTI_INFO_IMPL_HASH(4225993084u, ::Aspose::Words::ApiExamples::ExDocumentBase, ThisTypeBaseTypesInfo);

void ExDocumentBase::TestResourceLoadingCallback(System::SharedPtr<Aspose::Words::Document> doc)
{
    for (auto&& shape : System::IterateOver<Aspose::Words::Drawing::Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)))
    {
        ASSERT_TRUE(shape->get_HasImage());
        ASSERT_FALSE(System::TestTools::IsEmpty(shape->get_ImageData()->get_ImageBytes()));
    }
}


namespace gtest_test
{

class ExDocumentBase : public ::testing::Test
{
protected:
    static System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBase> s_instance;
    
    void SetUp() override
    {
        s_instance->SetUp();
    };
    
    static void SetUpTestCase()
    {
        s_instance = System::MakeObject<::Aspose::Words::ApiExamples::ExDocumentBase>();
        s_instance->OneTimeSetUp();
    };
    
    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };
    
};

System::SharedPtr<::Aspose::Words::ApiExamples::ExDocumentBase> ExDocumentBase::s_instance;

} // namespace gtest_test

void ExDocumentBase::Constructor()
{
    //ExStart
    //ExFor:DocumentBase
    //ExSummary:Shows how to initialize the subclasses of DocumentBase.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<Aspose::Words::DocumentBase>(), System::ObjectExt::GetType(doc).get_BaseType());
    
    auto glossaryDoc = System::MakeObject<Aspose::Words::BuildingBlocks::GlossaryDocument>();
    doc->set_GlossaryDocument(glossaryDoc);
    
    ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<Aspose::Words::DocumentBase>(), System::ObjectExt::GetType(glossaryDoc).get_BaseType());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBase, Constructor)
{
    s_instance->Constructor();
}

} // namespace gtest_test

void ExDocumentBase::SetPageColor()
{
    //ExStart
    //ExFor:DocumentBase.PageColor
    //ExSummary:Shows how to set the background color for all pages of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    builder->Writeln(u"Hello world!");
    
    doc->set_PageColor(System::Drawing::Color::get_LightGray());
    
    doc->Save(get_ArtifactsDir() + u"DocumentBase.SetPageColor.docx");
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBase.SetPageColor.docx");
    
    ASSERT_EQ(System::Drawing::Color::get_LightGray().ToArgb(), doc->get_PageColor().ToArgb());
}

namespace gtest_test
{

TEST_F(ExDocumentBase, SetPageColor)
{
    s_instance->SetPageColor();
}

} // namespace gtest_test

void ExDocumentBase::ImportNode()
{
    //ExStart
    //ExFor:DocumentBase.ImportNode(Node, Boolean)
    //ExSummary:Shows how to import a node from one document to another.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    
    srcDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(srcDoc, u"Source document first paragraph text."));
    dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild<System::SharedPtr<Aspose::Words::Run>>(System::MakeObject<Aspose::Words::Run>(dstDoc, u"Destination document first paragraph text."));
    
    // Every node has a parent document, which is the document that contains the node.
    // Inserting a node into a document that the node does not belong to will throw an exception.
    ASPOSE_ASSERT_NE(dstDoc, srcDoc->get_FirstSection()->get_Document());
    ASSERT_THROW(static_cast<std::function<void()>>([&dstDoc, &srcDoc]() -> void
    {
        dstDoc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(srcDoc->get_FirstSection());
    })(), System::ArgumentException);
    
    // Use the ImportNode method to create a copy of a node, which will have the document
    // that called the ImportNode method set as its new owner document.
    auto importedSection = System::ExplicitCast<Aspose::Words::Section>(dstDoc->ImportNode(srcDoc->get_FirstSection(), true));
    
    ASPOSE_ASSERT_EQ(dstDoc, importedSection->get_Document());
    
    // We can now insert the node into the document.
    dstDoc->AppendChild<System::SharedPtr<Aspose::Words::Section>>(importedSection);
    
    ASSERT_EQ(u"Destination document first paragraph text.\r\nSource document first paragraph text.\r\n", dstDoc->ToString(Aspose::Words::SaveFormat::Text));
    //ExEnd
    
    ASPOSE_ASSERT_NE(importedSection, srcDoc->get_FirstSection());
    ASPOSE_ASSERT_NE(importedSection->get_Document(), srcDoc->get_FirstSection()->get_Document());
    ASSERT_EQ(importedSection->get_Body()->get_FirstParagraph()->GetText(), srcDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
}

namespace gtest_test
{

TEST_F(ExDocumentBase, ImportNode)
{
    s_instance->ImportNode();
}

} // namespace gtest_test

void ExDocumentBase::ImportNodeCustom()
{
    //ExStart
    //ExFor:DocumentBase.ImportNode(Node, Boolean, ImportFormatMode)
    //ExSummary:Shows how to import node from source document to destination document with specific options.
    // Create two documents and add a character style to each document.
    // Configure the styles to have the same name, but different text formatting.
    auto srcDoc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Style> srcStyle = srcDoc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"My style");
    srcStyle->get_Font()->set_Name(u"Courier New");
    auto srcBuilder = System::MakeObject<Aspose::Words::DocumentBuilder>(srcDoc);
    srcBuilder->get_Font()->set_Style(srcStyle);
    srcBuilder->Writeln(u"Source document text.");
    
    auto dstDoc = System::MakeObject<Aspose::Words::Document>();
    System::SharedPtr<Aspose::Words::Style> dstStyle = dstDoc->get_Styles()->Add(Aspose::Words::StyleType::Character, u"My style");
    dstStyle->get_Font()->set_Name(u"Calibri");
    auto dstBuilder = System::MakeObject<Aspose::Words::DocumentBuilder>(dstDoc);
    dstBuilder->get_Font()->set_Style(dstStyle);
    dstBuilder->Writeln(u"Destination document text.");
    
    // Import the Section from the destination document into the source document, causing a style name collision.
    // If we use destination styles, then the imported source text with the same style name
    // as destination text will adopt the destination style.
    auto importedSection = System::ExplicitCast<Aspose::Words::Section>(dstDoc->ImportNode(srcDoc->get_FirstSection(), true, Aspose::Words::ImportFormatMode::UseDestinationStyles));
    ASSERT_EQ(u"Source document text.", importedSection->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->GetText().Trim());
    //ExSkip
    ASSERT_TRUE(System::TestTools::IsNull(dstDoc->get_Styles()->idx_get(u"My style_0")));
    //ExSkip
    ASSERT_EQ(dstStyle->get_Font()->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    ASSERT_EQ(dstStyle->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_StyleName());
    
    // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
    // and the naming clash resolves by adding a suffix.
    dstDoc->ImportNode(srcDoc->get_FirstSection(), true, Aspose::Words::ImportFormatMode::KeepDifferentStyles);
    ASSERT_EQ(dstStyle->get_Font()->get_Name(), dstDoc->get_Styles()->idx_get(u"My style")->get_Font()->get_Name());
    ASSERT_EQ(srcStyle->get_Font()->get_Name(), dstDoc->get_Styles()->idx_get(u"My style_0")->get_Font()->get_Name());
    //ExEnd
}

namespace gtest_test
{

TEST_F(ExDocumentBase, ImportNodeCustom)
{
    s_instance->ImportNodeCustom();
}

} // namespace gtest_test

void ExDocumentBase::BackgroundShape()
{
    //ExStart
    //ExFor:DocumentBase.BackgroundShape
    //ExSummary:Shows how to set a background shape for every page of a document.
    auto doc = System::MakeObject<Aspose::Words::Document>();
    
    ASSERT_TRUE(System::TestTools::IsNull(doc->get_BackgroundShape()));
    
    // The only shape type that we can use as a background is a rectangle.
    auto shapeRectangle = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    
    // There are two ways of using this shape as a page background.
    // 1 -  A flat color:
    shapeRectangle->set_FillColor(System::Drawing::Color::get_LightBlue());
    doc->set_BackgroundShape(shapeRectangle);
    
    doc->Save(get_ArtifactsDir() + u"DocumentBase.BackgroundShape.FlatColor.docx");
    
    // 2 -  An image:
    shapeRectangle = System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    shapeRectangle->get_ImageData()->SetImage(get_ImageDir() + u"Transparent background logo.png");
    
    // Adjust the image's appearance to make it more suitable as a watermark.
    shapeRectangle->get_ImageData()->set_Contrast(0.2);
    shapeRectangle->get_ImageData()->set_Brightness(0.7);
    
    doc->set_BackgroundShape(shapeRectangle);
    
    ASSERT_TRUE(doc->get_BackgroundShape()->get_HasImage());
    
    auto saveOptions = System::MakeObject<Aspose::Words::Saving::PdfSaveOptions>();
    saveOptions->set_CacheBackgroundGraphics(false);
    
    // Microsoft Word does not support shapes with images as backgrounds,
    // but we can still see these backgrounds in other save formats such as .pdf.
    doc->Save(get_ArtifactsDir() + u"DocumentBase.BackgroundShape.Image.pdf", saveOptions);
    //ExEnd
    
    doc = System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBase.BackgroundShape.FlatColor.docx");
    
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), doc->get_BackgroundShape()->get_FillColor().ToArgb());
    ASSERT_THROW(static_cast<std::function<void()>>([&doc]() -> void
    {
        doc->set_BackgroundShape(System::MakeObject<Aspose::Words::Drawing::Shape>(doc, Aspose::Words::Drawing::ShapeType::Triangle));
    })(), System::ArgumentException);
}

namespace gtest_test
{

TEST_F(ExDocumentBase, BackgroundShape)
{
    s_instance->BackgroundShape();
}

} // namespace gtest_test

void ExDocumentBase::ResourceLoadingCallback()
{
    auto doc = System::MakeObject<Aspose::Words::Document>();
    doc->set_ResourceLoadingCallback(System::MakeObject<Aspose::Words::ApiExamples::ExDocumentBase::ImageNameHandler>());
    
    auto builder = System::MakeObject<Aspose::Words::DocumentBuilder>(doc);
    
    // Images usually are inserted using a URI, or a byte array.
    // Every instance of a resource load will call our callback's ResourceLoading method.
    builder->InsertImage(u"Aspose logo");
    builder->InsertImage(u"Watermark");
    
    doc->Save(get_ArtifactsDir() + u"DocumentBase.ResourceLoadingCallback.docx");
    TestResourceLoadingCallback(System::MakeObject<Aspose::Words::Document>(get_ArtifactsDir() + u"DocumentBase.ResourceLoadingCallback.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExDocumentBase, ResourceLoadingCallback)
{
    s_instance->ResourceLoadingCallback();
}

} // namespace gtest_test

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
