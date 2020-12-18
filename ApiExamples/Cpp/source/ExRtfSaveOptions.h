#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <functional>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageData.h>
#include <Aspose.Words.Cpp/Model/Drawing/ImageType.h>
#include <Aspose.Words.Cpp/Model/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeCollection.h>
#include <Aspose.Words.Cpp/Model/Nodes/NodeType.h>
#include <Aspose.Words.Cpp/Model/Saving/RtfSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <system/exceptions.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/method_argument_tuple.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>
#include <gtest/gtest-spi.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExRtfSaveOptions : public ApiExampleBase
{
public:
    void ExportImages(bool doExportImagesForOldReaders)
    {
        //ExStart
        //ExFor:RtfSaveOptions
        //ExFor:RtfSaveOptions.ExportCompactSize
        //ExFor:RtfSaveOptions.ExportImagesForOldReaders
        //ExFor:RtfSaveOptions.SaveFormat
        //ExSummary:Shows how to save a document to .rtf with custom options.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Configure a RtfSaveOptions instance to make our output document more suitable for older devices
        auto options = MakeObject<RtfSaveOptions>();
        options->set_SaveFormat(SaveFormat::Rtf);
        options->set_ExportCompactSize(true);
        options->set_ExportImagesForOldReaders(doExportImagesForOldReaders);

        doc->Save(ArtifactsDir + u"RtfSaveOptions.ExportImages.rtf", options);
        //ExEnd

        if (doExportImagesForOldReaders)
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

    void SaveImagesAsWmf()
    {
        //ExStart
        //ExFor:RtfSaveOptions.SaveImagesAsWmf
        //ExSummary:Shows how to save all images as Wmf when saving to the Rtf document.
        // Open a document that contains images in the jpeg format
        auto doc = MakeObject<Document>(MyDir + u"Images.docx");

        SharedPtr<NodeCollection> shapes = doc->GetChildNodes(NodeType::Shape, true);
        auto shapeWithJpg = System::DynamicCast<Shape>(shapes->idx_get(0));
        ASSERT_EQ(ImageType::Jpeg, shapeWithJpg->get_ImageData()->get_ImageType());

        auto rtfSaveOptions = MakeObject<RtfSaveOptions>();
        rtfSaveOptions->set_SaveImagesAsWmf(true);
        doc->Save(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"RtfSaveOptions.SaveImagesAsWmf.rtf");

        shapes = doc->GetChildNodes(NodeType::Shape, true);
        auto shapeWithWmf = System::DynamicCast<Shape>(shapes->idx_get(0));
        ASSERT_EQ(ImageType::Wmf, shapeWithWmf->get_ImageData()->get_ImageType());
    }
};

} // namespace ApiExamples
