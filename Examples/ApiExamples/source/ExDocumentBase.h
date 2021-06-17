#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <functional>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BuildingBlocks/GlossaryDocument.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBase.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/ShapeType.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/ImportFormatMode.h>
#include <Aspose.Words.Cpp/Loading/IResourceLoadingCallback.h>
#include <Aspose.Words.Cpp/Loading/ResourceLoadingAction.h>
#include <Aspose.Words.Cpp/Loading/ResourceLoadingArgs.h>
#include <Aspose.Words.Cpp/Loading/ResourceType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphCollection.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/RunCollection.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Style.h>
#include <Aspose.Words.Cpp/StyleCollection.h>
#include <Aspose.Words.Cpp/StyleType.h>
#include <drawing/color.h>
#include <net/http_status_code.h>
#include <net/web_client.h>
#include <system/array.h>
#include <system/details/dispose_guard.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/io/file.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/type_info.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::BuildingBlocks;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Loading;

namespace ApiExamples {

class ExDocumentBase : public ApiExampleBase
{
public:
    void Constructor()
    {
        //ExStart
        //ExFor:DocumentBase
        //ExSummary:Shows how to initialize the subclasses of DocumentBase.
        auto doc = MakeObject<Document>();

        ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<DocumentBase>(), System::ObjectExt::GetType(doc).get_BaseType());

        auto glossaryDoc = MakeObject<GlossaryDocument>();
        doc->set_GlossaryDocument(glossaryDoc);

        ASPOSE_ASSERT_EQ(System::ObjectExt::GetType<DocumentBase>(), System::ObjectExt::GetType(glossaryDoc).get_BaseType());
        //ExEnd
    }

    void SetPageColor()
    {
        //ExStart
        //ExFor:DocumentBase.PageColor
        //ExSummary:Shows how to set the background color for all pages of a document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        doc->set_PageColor(System::Drawing::Color::get_LightGray());

        doc->Save(ArtifactsDir + u"DocumentBase.SetPageColor.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBase.SetPageColor.docx");

        ASSERT_EQ(System::Drawing::Color::get_LightGray().ToArgb(), doc->get_PageColor().ToArgb());
    }

    void ImportNode()
    {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, Boolean)
        //ExSummary:Shows how to import a node from one document to another.
        auto srcDoc = MakeObject<Document>();
        auto dstDoc = MakeObject<Document>();

        srcDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(srcDoc, u"Source document first paragraph text."));
        dstDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(dstDoc, u"Destination document first paragraph text."));

        // Every node has a parent document, which is the document that contains the node.
        // Inserting a node into a document that the node does not belong to will throw an exception.
        ASPOSE_ASSERT_NE(dstDoc, srcDoc->get_FirstSection()->get_Document());
        ASSERT_THROW(dstDoc->AppendChild(srcDoc->get_FirstSection()), System::ArgumentException);

        // Use the ImportNode method to create a copy of a node, which will have the document
        // that called the ImportNode method set as its new owner document.
        auto importedSection = System::DynamicCast<Section>(dstDoc->ImportNode(srcDoc->get_FirstSection(), true));

        ASPOSE_ASSERT_EQ(dstDoc, importedSection->get_Document());

        // We can now insert the node into the document.
        dstDoc->AppendChild(importedSection);

        ASSERT_EQ(u"Destination document first paragraph text.\r\nSource document first paragraph text.\r\n", dstDoc->ToString(SaveFormat::Text));
        //ExEnd

        ASPOSE_ASSERT_NE(importedSection, srcDoc->get_FirstSection());
        ASPOSE_ASSERT_NE(importedSection->get_Document(), srcDoc->get_FirstSection()->get_Document());
        ASSERT_EQ(importedSection->get_Body()->get_FirstParagraph()->GetText(), srcDoc->get_FirstSection()->get_Body()->get_FirstParagraph()->GetText());
    }

    void ImportNodeCustom()
    {
        //ExStart
        //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
        //ExSummary:Shows how to import node from source document to destination document with specific options.
        // Create two documents and add a character style to each document.
        // Configure the styles to have the same name, but different text formatting.
        auto srcDoc = MakeObject<Document>();
        SharedPtr<Style> srcStyle = srcDoc->get_Styles()->Add(StyleType::Character, u"My style");
        srcStyle->get_Font()->set_Name(u"Courier New");
        auto srcBuilder = MakeObject<DocumentBuilder>(srcDoc);
        srcBuilder->get_Font()->set_Style(srcStyle);
        srcBuilder->Writeln(u"Source document text.");

        auto dstDoc = MakeObject<Document>();
        SharedPtr<Style> dstStyle = dstDoc->get_Styles()->Add(StyleType::Character, u"My style");
        dstStyle->get_Font()->set_Name(u"Calibri");
        auto dstBuilder = MakeObject<DocumentBuilder>(dstDoc);
        dstBuilder->get_Font()->set_Style(dstStyle);
        dstBuilder->Writeln(u"Destination document text.");

        // Import the Section from the destination document into the source document, causing a style name collision.
        // If we use destination styles, then the imported source text with the same style name
        // as destination text will adopt the destination style.
        auto importedSection = System::DynamicCast<Section>(dstDoc->ImportNode(srcDoc->get_FirstSection(), true, ImportFormatMode::UseDestinationStyles));
        ASSERT_EQ(u"Source document text.", importedSection->get_Body()->get_Paragraphs()->idx_get(0)->get_Runs()->idx_get(0)->GetText().Trim());
        //ExSkip
        ASSERT_TRUE(dstDoc->get_Styles()->idx_get(u"My style_0") == nullptr);
        //ExSkip
        ASSERT_EQ(dstStyle->get_Font()->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_Name());
        ASSERT_EQ(dstStyle->get_Name(), importedSection->get_Body()->get_FirstParagraph()->get_Runs()->idx_get(0)->get_Font()->get_StyleName());

        // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
        // and the naming clash resolves by adding a suffix.
        dstDoc->ImportNode(srcDoc->get_FirstSection(), true, ImportFormatMode::KeepDifferentStyles);
        ASSERT_EQ(dstStyle->get_Font()->get_Name(), dstDoc->get_Styles()->idx_get(u"My style")->get_Font()->get_Name());
        ASSERT_EQ(srcStyle->get_Font()->get_Name(), dstDoc->get_Styles()->idx_get(u"My style_0")->get_Font()->get_Name());
        //ExEnd
    }

    void BackgroundShape()
    {
        //ExStart
        //ExFor:DocumentBase.BackgroundShape
        //ExSummary:Shows how to set a background shape for every page of a document.
        auto doc = MakeObject<Document>();

        ASSERT_TRUE(doc->get_BackgroundShape() == nullptr);

        // The only shape type that we can use as a background is a rectangle.
        auto shapeRectangle = MakeObject<Shape>(doc, ShapeType::Rectangle);

        // There are two ways of using this shape as a page background.
        // 1 -  A flat color:
        shapeRectangle->set_FillColor(System::Drawing::Color::get_LightBlue());
        doc->set_BackgroundShape(shapeRectangle);

        doc->Save(ArtifactsDir + u"DocumentBase.BackgroundShape.FlatColor.docx");

        // 2 -  An image:
        shapeRectangle = MakeObject<Shape>(doc, ShapeType::Rectangle);
        shapeRectangle->get_ImageData()->SetImage(ImageDir + u"Transparent background logo.png");

        // Adjust the image's appearance to make it more suitable as a watermark.
        shapeRectangle->get_ImageData()->set_Contrast(0.2);
        shapeRectangle->get_ImageData()->set_Brightness(0.7);

        doc->set_BackgroundShape(shapeRectangle);

        ASSERT_TRUE(doc->get_BackgroundShape()->get_HasImage());

        // Microsoft Word does not support shapes with images as backgrounds,
        // but we can still see these backgrounds in other save formats such as .pdf.
        doc->Save(ArtifactsDir + u"DocumentBase.BackgroundShape.Image.pdf");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"DocumentBase.BackgroundShape.FlatColor.docx");

        ASSERT_EQ(System::Drawing::Color::get_LightBlue().ToArgb(), doc->get_BackgroundShape()->get_FillColor().ToArgb());
        ASSERT_THROW(doc->set_BackgroundShape(MakeObject<Shape>(doc, ShapeType::Triangle)), System::ArgumentException);
    }

    //ExStart
    //ExFor:DocumentBase.ResourceLoadingCallback
    //ExFor:IResourceLoadingCallback
    //ExFor:IResourceLoadingCallback.ResourceLoading(ResourceLoadingArgs)
    //ExFor:ResourceLoadingAction
    //ExFor:ResourceLoadingArgs
    //ExFor:ResourceLoadingArgs.OriginalUri
    //ExFor:ResourceLoadingArgs.ResourceType
    //ExFor:ResourceLoadingArgs.SetData(Byte[])
    //ExFor:ResourceType
    //ExSummary:Shows how to customize the process of loading external resources into a document.
    void ResourceLoadingCallback()
    {
        auto doc = MakeObject<Document>();
        doc->set_ResourceLoadingCallback(MakeObject<ExDocumentBase::ImageNameHandler>());

        auto builder = MakeObject<DocumentBuilder>(doc);

        // Images usually are inserted using a URI, or a byte array.
        // Every instance of a resource load will call our callback's ResourceLoading method.
        builder->InsertImage(u"Google logo");
        builder->InsertImage(u"Aspose logo");
        builder->InsertImage(u"Watermark");

        ASSERT_EQ(3, doc->GetChildNodes(NodeType::Shape, true)->get_Count());

        doc->Save(ArtifactsDir + u"DocumentBase.ResourceLoadingCallback.docx");
        TestResourceLoadingCallback(MakeObject<Document>(ArtifactsDir + u"DocumentBase.ResourceLoadingCallback.docx"));
        //ExSkip
    }

    /// <summary>
    /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
    /// This will separate image loading logic from the rest of the document construction.
    /// </summary>
    class ImageNameHandler : public IResourceLoadingCallback
    {
    public:
        ResourceLoadingAction ResourceLoading(SharedPtr<ResourceLoadingArgs> args) override
        {
            // If this callback encounters one of the image shorthands while loading an image,
            // it will apply unique logic for each defined shorthand instead of treating it as a URI.
            if (args->get_ResourceType() == ResourceType::Image)
            {
                String imageUri = args->get_OriginalUri();
                if (imageUri == u"Google logo")
                {
                    {
                        auto webClient = MakeObject<System::Net::WebClient>();
                        args->SetData(webClient->DownloadData(u"http://www.google.com/images/logos/ps_logo2.png"));
                    }
                    return ResourceLoadingAction::UserProvided;
                }
                else if (imageUri == u"Aspose logo")
                {
                    args->SetData(System::IO::File::ReadAllBytes(ImageDir + u"Logo.jpg"));
                    return ResourceLoadingAction::UserProvided;
                }
                else if (imageUri == u"Watermark")
                {
                    args->SetData(System::IO::File::ReadAllBytes(ImageDir + u"Transparent background logo.png"));
                    return ResourceLoadingAction::UserProvided;
                }
            }
            return ResourceLoadingAction::Default;
        }
    };
    //ExEnd

    void TestResourceLoadingCallback(SharedPtr<Document> doc)
    {
        for (const auto& shape : System::IterateOver<Shape>(doc->GetChildNodes(NodeType::Shape, true)))
        {
            ASSERT_TRUE(shape->get_HasImage());
            ASSERT_FALSE(System::TestTools::IsEmpty(shape->get_ImageData()->get_ImageBytes()));
        }

        TestUtil::VerifyWebResponseStatusCode(System::Net::HttpStatusCode::OK, u"http://www.google.com/images/logos/ps_logo2.png");
    }
};

} // namespace ApiExamples
