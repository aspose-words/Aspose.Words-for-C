#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/PsSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/Settings/MultiplePagesType.h>
#include <system/enumerator_adapter.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Saving;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExPsSaveOptions : public ApiExampleBase
{
public:
    void UseBookFoldPrintingSettings(bool renderTextAsBookFold)
    {
        //ExStart
        //ExFor:PsSaveOptions
        //ExFor:PsSaveOptions.SaveFormat
        //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Create a "PsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to PostScript.
        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output Postscript document in a way that helps us make a booklet out of it.
        // Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
        auto saveOptions = MakeObject<PsSaveOptions>();
        saveOptions->set_SaveFormat(SaveFormat::Ps);
        saveOptions->set_UseBookFoldPrintingSettings(renderTextAsBookFold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        for (auto s : System::IterateOver<Section>(doc->get_Sections()))
        {
            s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
        }

        // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
        // and the contents will line up in a way that creates a booklet.
        doc->Save(ArtifactsDir + u"PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
        //ExEnd
    }
};

} // namespace ApiExamples
