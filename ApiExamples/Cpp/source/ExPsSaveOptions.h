#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <system/enumerator_adapter.h>

#include "ApiExampleBase.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExPsSaveOptions : public ApiExampleBase
{
public:
    void UseBookFoldPrintingSettings()
    {
        //ExStart
        //ExFor:PsSaveOptions
        //ExFor:PsSaveOptions.SaveFormat
        //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to create a bookfold in the PostScript format.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Configure both page setup and PsSaveOptions to create a book fold
        for (auto s : System::IterateOver<Section>(doc->get_Sections()))
        {
            s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
        }

        auto saveOptions = MakeObject<PsSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Ps);
        saveOptions->set_UseBookFoldPrintingSettings(true);

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // in the order they come out of the printer and then folding down the middle
        doc->Save(ArtifactsDir + u"PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
        //ExEnd
    }
};

} // namespace ApiExamples
