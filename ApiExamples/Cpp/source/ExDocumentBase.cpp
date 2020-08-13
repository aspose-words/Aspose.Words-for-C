// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExDocumentBase.h"

#include <testing/test_predicates.h>
#include <system/test_tools/test_tools.h>
#include <system/test_tools/compare.h>
#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/enumerator_adapter.h>
#include <system/details/dispose_guard.h>
#include <system/array.h>
#include <net/web_client.h>
#include <gtest/gtest.h>
#include <functional>
#include <drawing/image_converter.h>
#include <drawing/image.h>
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
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Model/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Model/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Model/Importing/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Model/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBase.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/BuildingBlocks/GlossaryDocument.h>

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Loading;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(4066832806u, ::ApiExamples::ExDocumentBase::ImageNameHandler, ThisTypeBaseTypesInfo);

Aspose::Words::Loading::ResourceLoadingAction ExDocumentBase::ImageNameHandler::ResourceLoading(SharedPtr<Aspose::Words::Loading::ResourceLoadingArgs> args)
{
    if (args->get_ResourceType() == Aspose::Words::Loading::ResourceType::Image)
    {
        // builder.InsertImage expects a uri so inputs like "Google Logo" would normally trigger a FileNotFoundException
        // We can still process those inputs and find an image any way we like, as long as an image byte array is passed to args.SetData()
        if (args->get_OriginalUri() == u"Google Logo")
        {
            {
                auto webClient = MakeObject<System::Net::WebClient>();
                ArrayPtr<uint8_t> imageBytes = webClient->DownloadData(u"http://www.google.com/images/logos/ps_logo2.png");
                args->SetData(imageBytes);
                // We need this return statement any time a resource is loaded in a custom manner
                return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
            }
        }

        if (args->get_OriginalUri() == u"Aspose Logo")
        {
            {
                auto webClient = MakeObject<System::Net::WebClient>();
                ArrayPtr<uint8_t> imageBytes = webClient->DownloadData(AsposeLogoUrl);
                args->SetData(imageBytes);
                return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
            }
        }

        // We can find and add an image any way we like, as long as args.SetData() is called with some image byte array as a parameter
        if (args->get_OriginalUri() == u"My Watermark")
        {
            SharedPtr<System::Drawing::Image> watermark = System::Drawing::Image::FromFile(ImageDir + u"Transparent background logo.png");

            auto converter = MakeObject<System::Drawing::ImageConverter>();
            ArrayPtr<uint8_t> imageBytes = System::DynamicCast<System::Array<uint8_t>>(converter->ConvertTo(watermark, System::ObjectExt::GetType<System::Array<uint8_t>>()));
            args->SetData(imageBytes);

            return Aspose::Words::Loading::ResourceLoadingAction::UserProvided;
        }
    }

    // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
    return Aspose::Words::Loading::ResourceLoadingAction::Default;
}

void ExDocumentBase::TestResourceLoadingCallback(SharedPtr<Aspose::Words::Document> doc)
{
    for (auto shape : System::IterateOver<Shape>(doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)))
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
    static SharedPtr<::ApiExamples::ExDocumentBase> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExDocumentBase>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExDocumentBase> ExDocumentBase::s_instance;

} // namespace gtest_test

void ExDocumentBase::Constructor()
{
    //ExStart
    //ExFor:DocumentBase
    //ExSummary:Shows how to initialize the subclasses of DocumentBase.
    // DocumentBase is the abstract base class for the Document and GlossaryDocument classes
    auto doc = MakeObject<Document>();

    auto glossaryDoc = MakeObject<GlossaryDocument>();
    doc->set_GlossaryDocument(glossaryDoc);
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
    //ExSummary:Shows how to set the page color.
    auto doc = MakeObject<Document>();

    doc->set_PageColor(System::Drawing::Color::get_LightGray());

    doc->Save(ArtifactsDir + u"DocumentBase.SetPageColor.docx");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBase.SetPageColor.docx");
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
    //ExSummary:Shows how to import node from source document to destination document.
    auto src = MakeObject<Document>();
    auto dst = MakeObject<Document>();

    // Add text to both documents
    src->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(src, u"Source document first paragraph text."));
    dst->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(dst, u"Destination document first paragraph text."));

    // In order for a child node to be successfully appended to another node in a document,
    // both nodes must have the same parent document, or an exception is thrown
    ASPOSE_ASSERT_NE(dst, src->get_FirstSection()->get_Document());

    auto appendSection = [&dst, &src]()
    {
        dst->AppendChild(src->get_FirstSection());
    };

    ASSERT_THROW(appendSection(), System::ArgumentException);

    // For that reason, we can't just append a section of the source document to the destination document using Node.AppendChild()
    // Document.ImportNode() lets us get around this by creating a clone of a node and sets its parent to the calling document
    auto importedSection = System::DynamicCast<Aspose::Words::Section>(dst->ImportNode(src->get_FirstSection(), true));

    // Now it is ready to be placed in the document
    dst->AppendChild(importedSection);

    // Our document now contains both the original and imported section
    ASSERT_EQ(u"Destination document first paragraph text.\r\nSource document first paragraph text.\r\n", dst->ToString(Aspose::Words::SaveFormat::Text));
    //ExEnd

    ASPOSE_ASSERT_NE(importedSection, src->get_FirstSection());
    ASPOSE_ASSERT_NE(importedSection->get_Document(), src->get_FirstSection()->get_Document());
    ASSERT_EQ(importedSection->get_Body()->get_FirstParagraph()->GetText(), src->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
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
    //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
    //ExSummary:Shows how to import node from source document to destination document with specific options.
    // Create two documents with two styles that differ in font but have the same name
    auto src = MakeObject<Document>();
    SharedPtr<Style> srcStyle = src->get_Styles()->Add(Aspose::Words::StyleType::Character, u"My style");
    srcStyle->get_Font()->set_Name(u"Courier New");
    auto srcBuilder = MakeObject<DocumentBuilder>(src);
    srcBuilder->get_Font()->set_Style(srcStyle);
    srcBuilder->Writeln(u"Source document text.");

    auto dst = MakeObject<Document>();
    SharedPtr<Style> dstStyle = dst->get_Styles()->Add(Aspose::Words::StyleType::Character, u"My style");
    dstStyle->get_Font()->set_Name(u"Calibri");
    auto dstBuilder = MakeObject<DocumentBuilder>(dst);
    dstBuilder->get_Font()->set_Style(dstStyle);
    dstBuilder->Writeln(u"Destination document text.");

    // Import the Section from the destination document into the source document, causing a style name collision
    // If we use destination styles then the imported source text with the same style name as destination text
    // will adopt the destination style
    auto importedSection = System::DynamicCast<Aspose::Words::Section>(dst->ImportNode(src->get_FirstSection(), true, Aspose::Words::ImportFormatMode::UseDestinationStyles));
    ASSERT_EQ(u"Source document text.", importedSection->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->GetText().Trim());
    //ExSkip
    ASSERT_TRUE(dst->get_Styles()->idx_get(u"My style_0") == nullptr);
    //ExSkip
    ASSERT_EQ(dstStyle->get_Font()->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
    ASSERT_EQ(dstStyle->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_StyleName());

    // If we use ImportFormatMode.KeepDifferentStyles,
    // the source style is preserved and the naming clash is resolved by adding a suffix
    dst->ImportNode(src->get_FirstSection(), true, Aspose::Words::ImportFormatMode::KeepDifferentStyles);
    ASSERT_EQ(dstStyle->get_Font()->get_Name(), dst->get_Styles()->idx_get(u"My style")->get_Font()->get_Name());
    ASSERT_EQ(srcStyle->get_Font()->get_Name(), dst->get_Styles()->idx_get(u"My style_0")->get_Font()->get_Name());
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
    //ExSummary:Shows how to set the background shape of a document.
    auto doc = MakeObject<Document>();
    ASSERT_TRUE(doc->get_BackgroundShape() == nullptr);

    // A background shape can only be a rectangle
    // We will set the color of this rectangle to light blue
    auto shapeRectangle = MakeObject<Shape>(doc, Aspose::Words::Drawing::ShapeType::Rectangle);
    doc->set_BackgroundShape(shapeRectangle);

    // This rectangle covers the entire page in the output document
    // We can also do this by setting doc.PageColor
    shapeRectangle->set_FillColor(System::Drawing::Color::get_LightBlue());
    doc->Save(ArtifactsDir + u"DocumentBase.BackgroundShapeFlatColor.docx");

    // Setting the image will override the flat background color with the image
    shapeRectangle->get_ImageData()->SetImage(ImageDir + u"Transparent background logo.png");
    ASSERT_TRUE(doc->get_BackgroundShape()->get_HasImage());

    // This image is a photo with a white background
    // To make it suitable as a watermark, we will need to do some image processing
    // The default values for these variables are 0.5, so here we are lowering the contrast and increasing the brightness
    shapeRectangle->get_ImageData()->set_Contrast(0.2);
    shapeRectangle->get_ImageData()->set_Brightness(0.7);

    // Microsoft Word does not support images in background shapes, so even though we set the background as an image,
    // the output will show a light blue background like before
    // However, we can see our watermark in an output pdf
    doc->Save(ArtifactsDir + u"DocumentBase.BackgroundShape.pdf");
    //ExEnd

    doc = MakeObject<Document>(ArtifactsDir + u"DocumentBase.BackgroundShapeFlatColor.docx");
    ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), doc->get_BackgroundShape()->get_FillColor().ToArgb());
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
    auto doc = MakeObject<Document>();

    // Enable our custom image loading
    doc->set_ResourceLoadingCallback(MakeObject<ExDocumentBase::ImageNameHandler>());

    auto builder = MakeObject<DocumentBuilder>(doc);

    // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
    // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
    builder->InsertImage(u"Google Logo");
    builder->InsertImage(u"Aspose Logo");
    builder->InsertImage(u"My Watermark");

    // Images belong to Shape objects, which are placed and scaled in the document
    ASSERT_EQ(3, doc->GetChildNodes(Aspose::Words::NodeType::Shape, true)->get_Count());

    doc->Save(ArtifactsDir + u"DocumentBase.ResourceLoadingCallback.docx");
    TestResourceLoadingCallback(MakeObject<Document>(ArtifactsDir + u"DocumentBase.ResourceLoadingCallback.docx"));
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
