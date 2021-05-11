#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/PclSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <system/collections/ienumerable.h>
#include <system/enumerator_adapter.h>
#include <system/linq/enumerable.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

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
        //ExSummary:Shows how to rasterize complex elements while saving a document to PCL.
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Pcl);
        saveOptions->set_RasterizeTransformedElements(true);

        doc->Save(ArtifactsDir + u"PclSaveOptions.RasterizeElements.pcl", saveOptions);
        //ExEnd
    }

    void FallbackFontName()
    {
        //ExStart
        //ExFor:PclSaveOptions.FallbackFontName
        //ExSummary:Shows how to declare a font that a printer will apply to printed text as a substitute should its original font be unavailable.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Non-existent font");
        builder->Write(u"Hello world!");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->set_FallbackFontName(u"Times New Roman");

        // This document will instruct the printer to apply "Times New Roman" to the text with the missing font.
        // Should "Times New Roman" also be unavailable, the printer will default to the "Arial" font.
        doc->Save(ArtifactsDir + u"PclSaveOptions.SetPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    void AddPrinterFont()
    {
        //ExStart
        //ExFor:PclSaveOptions.AddPrinterFont(string, string)
        //ExSummary:Shows how to get a printer to substitute all instances of a specific font with a different font.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_Font()->set_Name(u"Courier");
        builder->Write(u"Hello world!");

        auto saveOptions = MakeObject<PclSaveOptions>();
        saveOptions->AddPrinterFont(u"Courier New", u"Courier");

        // When printing this document, the printer will use the "Courier New" font
        // to access places where our document used the "Courier" font.
        doc->Save(ArtifactsDir + u"PclSaveOptions.AddPrinterFont.pcl", saveOptions);
        //ExEnd
    }

    void GetPreservedPaperTrayInformation()
    {
        auto doc = MakeObject<Document>(MyDir + u"Rendering.docx");

        // Paper tray information is now preserved when saving document to PCL format.
        // Following information is transferred from document's model to PCL file.
        for (auto section : System::IterateOver(doc->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
        {
            section->get_PageSetup()->set_FirstPageTray(15);
            section->get_PageSetup()->set_OtherPagesTray(12);
        }

        doc->Save(ArtifactsDir + u"PclSaveOptions.GetPreservedPaperTrayInformation.pcl");
    }
};

} // namespace ApiExamples
