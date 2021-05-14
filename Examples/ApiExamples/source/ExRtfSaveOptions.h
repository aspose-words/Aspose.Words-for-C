#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <gtest/gtest-spi.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExRtfSaveOptions : public ApiExampleBase
{
public:
    void ExportImages(bool exportImagesForOldReaders)
    {
        //ExStart
        //ExFor:RtfSaveOptions
        //ExFor:RtfSaveOptions.ExportCompactSize
        //ExFor:RtfSaveOptions.ExportImagesForOldReaders
        //ExFor:RtfSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to .rtf with custom options.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
        auto options = MakeObject<RtfSaveOptions>();

        ASSERT_EQ(SaveFormat::Rtf, options->get_SaveFormat());

        // Set the "ExportCompactSize" property to "true" to
        // reduce the saved document's size at the cost of right-to-left text compatibility.
        options->set_ExportCompactSize(true);

        // Set the "ExportImagesFotOldReaders" property to "true" to use extra keywords to ensure that our document is
        // compatible with pre-Microsoft Word 97 readers and WordPad.
        // Set the "ExportImagesFotOldReaders" property to "false" to reduce the size of the document,
        // but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
        options->set_ExportImagesForOldReaders(exportImagesForOldReaders);

        doc->Save(ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf", options);
        //ExEnd

        if (exportImagesForOldReaders)
        {
            TestUtil::FileContainsString(u"nonshppict", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
            TestUtil::FileContainsString(u"shprslt", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf");
        }
        else
        {
            if (!IsRunningOnMono())
            {
                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"nonshppict", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf"),
                                     "not found in the provided source");

                EXPECT_FATAL_FAILURE(TestUtil::FileContainsString(u"shprslt", ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf"),
                                     "not found in the provided source");
            }
        }
    }

    void SaveImagesAsWmf(bool saveImagesAsWmf)
    {
        //ExStart
        //ExFor:RtfSaveOptions.SaveImagesAsWmf
        //ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Jpeg image:");
        SharedPtr<Shape> imageShape = builder->InsertImage(ImageDir + u"Logo.jpg");

        ASSERT_EQ(ImageType::Jpeg, imageShape->get_ImageData()->get_ImageType());

        builder->InsertParagraph();
        builder->Writeln(u"Png image:");
        imageShape = builder->InsertImage(ImageDir + u"Transparent background logo.png");

        ASSERT_EQ(ImageType::Png, imageShape->get_ImageData()->get_ImageType());

        // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
        auto rtfSaveOptions = MakeObject<RtfSaveOptions>();

        // Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
        // Doing so will help readers such as WordPad to read our document.
        // Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
        // as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
        rtfSaveOptions->set_SaveImagesAsWmf(saveImagesAsWmf);

        doc->Save(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);

        doc = MakeObject<Document>(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf");

        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);

        if (saveImagesAsWmf)
        {
            ASSERT_EQ(ImageType::Wmf, (System::DynamicCast<Shape>(shapes->idx_get(0)))->get_ImageData()->get_ImageType());
            ASSERT_EQ(ImageType::Wmf, (System::DynamicCast<Shape>(shapes->idx_get(1)))->get_ImageData()->get_ImageType());
        }
        else
        {
            ASSERT_EQ(ImageType::Jpeg, (System::DynamicCast<Shape>(shapes->idx_get(0)))->get_ImageData()->get_ImageType());
            ASSERT_EQ(ImageType::Png, (System::DynamicCast<Shape>(shapes->idx_get(1)))->get_ImageData()->get_ImageType());
        }
        //ExEnd
    }
};

} // namespace ApiExamples
