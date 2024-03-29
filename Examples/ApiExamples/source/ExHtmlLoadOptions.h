﻿#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/DigitalSignatures/DigitalSignatureUtil.h>
#include <Aspose.Words.Cpp/DigitalSignatures/SignOptions.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Fields/FormField.h>
#include <Aspose.Words.Cpp/IWarningCallback.h>
#include <Aspose.Words.Cpp/LoadFormat.h>
#include <Aspose.Words.Cpp/Loading/HtmlControlType.h>
#include <Aspose.Words.Cpp/Loading/HtmlLoadOptions.h>
#include <Aspose.Words.Cpp/Markup/SdtListItem.h>
#include <Aspose.Words.Cpp/Markup/SdtListItemCollection.h>
#include <Aspose.Words.Cpp/Markup/StructuredDocumentTag.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/WarningInfo.h>
#include <Aspose.Words.Cpp/WarningSource.h>
#include <Aspose.Words.Cpp/WarningType.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/io/memory_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::Loading;
using namespace Aspose::Words::Markup;

namespace ApiExamples {

class ExHtmlLoadOptions : public ApiExampleBase
{
public:
    void SupportVml(bool supportVml)
    {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor()
        //ExFor:HtmlLoadOptions.SupportVml
        //ExSummary:Shows how to support conditional comments while loading an HTML document.
        auto loadOptions = MakeObject<HtmlLoadOptions>();

        // If the value is true, then we take VML code into account while parsing the loaded document.
        loadOptions->set_SupportVml(supportVml);

        // This document contains a JPEG image within "<!--[if gte vml 1]>" tags,
        // and a different PNG image within "<![if !vml]>" tags.
        // If we set the "SupportVml" flag to "true", then Aspose.Words will load the JPEG.
        // If we set this flag to "false", then Aspose.Words will only load the PNG.
        auto doc = MakeObject<Document>(MyDir + u"VML conditional.htm", loadOptions);

        if (supportVml)
        {
            ASSERT_EQ(ImageType::Jpeg, (System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_ImageData()->get_ImageType());
        }
        else
        {
            ASSERT_EQ(ImageType::Png, (System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_ImageData()->get_ImageType());
        }
        //ExEnd

        auto imageShape = System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true));

        if (supportVml)
        {
            TestUtil::VerifyImageInShape(400, 400, ImageType::Jpeg, imageShape);
        }
        else
        {
            TestUtil::VerifyImageInShape(400, 400, ImageType::Png, imageShape);
        }
    }

    void EncryptedHtml()
    {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(String)
        //ExSummary:Shows how to encrypt an Html document, and then open it using a password.
        // Create and sign an encrypted HTML document from an encrypted .docx.
        SharedPtr<CertificateHolder> certificateHolder = CertificateHolder::Create(MyDir + u"morzal.pfx", u"aw");

        auto signOptions = MakeObject<SignOptions>();
        signOptions->set_Comments(u"Comment");
        signOptions->set_SignTime(System::DateTime::get_Now());
        signOptions->set_DecryptionPassword(u"docPassword");

        String inputFileName = MyDir + u"Encrypted.docx";
        String outputFileName = ArtifactsDir + u"HtmlLoadOptions.EncryptedHtml.html";
        DigitalSignatureUtil::Sign(inputFileName, outputFileName, certificateHolder, signOptions);

        // To load and read this document, we will need to pass its decryption
        // password using a HtmlLoadOptions object.
        auto loadOptions = MakeObject<HtmlLoadOptions>(u"docPassword");

        ASSERT_EQ(signOptions->get_DecryptionPassword(), loadOptions->get_Password());

        auto doc = MakeObject<Document>(outputFileName, loadOptions);

        ASSERT_EQ(u"Test encrypted document.", doc->GetText().Trim());
        //ExEnd
    }

    void BaseUri()
    {
        //ExStart
        //ExFor:HtmlLoadOptions.#ctor(LoadFormat,String,String)
        //ExFor:LoadOptions.#ctor(LoadFormat, String, String)
        //ExFor:LoadOptions.LoadFormat
        //ExFor:LoadFormat
        //ExSummary:Shows how to specify a base URI when opening an html document.
        // Suppose we want to load an .html document that contains an image linked by a relative URI
        // while the image is in a different location. In that case, we will need to resolve the relative URI into an absolute one.
        // We can provide a base URI using an HtmlLoadOptions object.
        auto loadOptions = MakeObject<HtmlLoadOptions>(LoadFormat::Html, u"", ImageDir);

        ASSERT_EQ(LoadFormat::Html, loadOptions->get_LoadFormat());

        auto doc = MakeObject<Document>(MyDir + u"Missing image.html", loadOptions);

        // While the image was broken in the input .html, our custom base URI helped us repair the link.
        auto imageShape = System::ExplicitCast<Shape>(doc->GetChildNodes(NodeType::Shape, true)->idx_get(0));
        ASSERT_TRUE(imageShape->get_IsImage());

        // This output document will display the image that was missing.
        doc->Save(ArtifactsDir + u"HtmlLoadOptions.BaseUri.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"HtmlLoadOptions.BaseUri.docx");

        ASSERT_TRUE((System::ExplicitCast<Shape>(doc->GetChild(NodeType::Shape, 0, true)))->get_ImageData()->get_ImageBytes()->get_Length() > 0);
    }

    void GetSelectAsSdt()
    {
        //ExStart
        //ExFor:HtmlLoadOptions.PreferredControlType
        //ExSummary:Shows how to set preferred type of document nodes that will represent imported <input> and <select> elements.
        const String html = u"\r\n                <html>\r\n                    <select name='ComboBox' size='1'>\r\n                        <option "
                            u"value='val1'>item1</option>\r\n                        <option value='val2'></option>                        \r\n                "
                            u"    </select>\r\n                </html>\r\n            ";

        auto htmlLoadOptions = MakeObject<HtmlLoadOptions>();
        htmlLoadOptions->set_PreferredControlType(HtmlControlType::StructuredDocumentTag);

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), htmlLoadOptions);
        SharedPtr<NodeCollection> nodes = doc->GetChildNodes(NodeType::StructuredDocumentTag, true);

        auto tag = System::ExplicitCast<StructuredDocumentTag>(nodes->idx_get(0));
        //ExEnd

        ASSERT_EQ(2, tag->get_ListItems()->get_Count());

        ASSERT_EQ(u"val1", tag->get_ListItems()->idx_get(0)->get_Value());
        ASSERT_EQ(u"val2", tag->get_ListItems()->idx_get(1)->get_Value());
    }

    void GetInputAsFormField()
    {
        const String html =
            u"\r\n                <html>\r\n                    <input type='text' value='Input value text' />\r\n                </html>\r\n            ";

        // By default, "HtmlLoadOptions.PreferredControlType" value is "HtmlControlType.FormField".
        // So, we do not set this value.
        auto htmlLoadOptions = MakeObject<HtmlLoadOptions>();

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_UTF8()->GetBytes(html)), htmlLoadOptions);
        SharedPtr<NodeCollection> nodes = doc->GetChildNodes(NodeType::FormField, true);

        ASSERT_EQ(1, nodes->get_Count());

        auto formField = System::ExplicitCast<FormField>(nodes->idx_get(0));
        ASSERT_EQ(u"Input value text", formField->get_Result());
    }
};

} // namespace ApiExamples
