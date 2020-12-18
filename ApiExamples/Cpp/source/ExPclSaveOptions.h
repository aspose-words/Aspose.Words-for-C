#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/PclSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/linq/enumerable.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExPclSaveOptions : public ApiExampleBase
{
public:
    void RasterizeElements()
    {
        //ExStart
        //ExFor:PclSaveOptions
        //ExFor:PclSaveOptions.SaveFormat
        //ExFor:PclSaveOptions.RasterizeTransformedElements
        //ExSummary:Shows how to rasterize complex elements before saving.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Pcl);
        saveOptions->set_RasterizeTransformedElements(true);

        doc->Save(ArtifactsDir + u"PclSaveOptions.RasterizeElements.pcl", saveOptions);
        //ExEnd
    }

    void SetPrinterFont()
    {
        //ExStart
        //ExFor:PclSaveOptions.AddPrinterFont(string, string)
        //ExFor:PclSaveOptions.FallbackFontName
        //ExSummary:Shows how to add information about font that is uploaded to the printer and set the font that will be used if no expected font is found in printer and built-in fonts collections.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->AddPrinterFont(u"Courier", u"Courier");
        saveOptions->set_FallbackFontName(u"Times New Roman");

        doc->Save(ArtifactsDir + u"PclSaveOptions.SetPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    void GetPreservedPaperTrayInformation()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Paper tray information is now preserved when saving document to PCL format
        // Following information is transferred from document's model to PCL file
        for (auto section : System::IterateOver(doc->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
        {
            section->get_PageSetup()->set_FirstPageTray(15);
            section->get_PageSetup()->set_OtherPagesTray(12);
        }

        doc->Save(ArtifactsDir + u"PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
    }
};

} // namespace ApiExamples
