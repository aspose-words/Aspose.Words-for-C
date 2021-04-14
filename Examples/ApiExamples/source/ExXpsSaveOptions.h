#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/BreakType.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>
#include <Aspose.Words.Cpp/Model/Nodes/Node.h>
#include <Aspose.Words.Cpp/Model/Saving/OutlineOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/PageSet/PageSet.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Saving/XpsSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Sections/PageSetup.h>
#include <Aspose.Words.Cpp/Model/Sections/Section.h>
#include <Aspose.Words.Cpp/Model/Sections/SectionCollection.h>
#include <Aspose.Words.Cpp/Model/Settings/MultiplePagesType.h>
#include <Aspose.Words.Cpp/Model/Styles/StyleIdentifier.h>
#include <Aspose.Words.Cpp/Model/Text/ParagraphFormat.h>
#include <system/array.h>
#include <system/enumerator_adapter.h>
#include <system/io/file_info.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

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

class ExXpsSaveOptions : public ApiExampleBase
{
public:
    void OutlineLevels()
    {
        //ExStart
        //ExFor:XpsSaveOptions
        //ExFor:XpsSaveOptions.#ctor
        //ExFor:XpsSaveOptions.OutlineOptions
        //ExFor:XpsSaveOptions.SaveFormat
        //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading1);

        ASSERT_TRUE(builder->get_ParagraphFormat()->get_IsHeading());

        builder->Writeln(u"Heading 1");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading2);

        builder->Writeln(u"Heading 1.1");
        builder->Writeln(u"Heading 1.2");

        builder->get_ParagraphFormat()->set_StyleIdentifier(StyleIdentifier::Heading3);

        builder->Writeln(u"Heading 1.2.1");
        builder->Writeln(u"Heading 1.2.2");

        // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        auto saveOptions = MakeObject<XpsSaveOptions>();

        ASSERT_EQ(SaveFormat::Xps, saveOptions->get_SaveFormat());

        // The output XPS document will contain an outline, a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        // The last two headings we have inserted above will not appear.
        saveOptions->get_OutlineOptions()->set_HeadingsOutlineLevels(2);

        doc->Save(ArtifactsDir + u"XpsSaveOptions.OutlineLevels.xps", saveOptions);
        //ExEnd
    }

    void BookFold(bool renderTextAsBookFold)
    {
        //ExStart
        //ExFor:XpsSaveOptions.#ctor(SaveFormat)
        //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
        auto doc = MakeObject<Document>(MyDir + u"Paragraphs.docx");

        // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        auto xpsOptions = MakeObject<XpsSaveOptions>(SaveFormat::Xps);

        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output XPS in a way that helps us use it to make a booklet.
        // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
        xpsOptions->set_UseBookFoldPrintingSettings(renderTextAsBookFold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        if (renderTextAsBookFold)
        {
            for (auto s : System::IterateOver<Section>(doc->get_Sections()))
            {
                s->get_PageSetup()->set_MultiplePages(MultiplePagesType::BookFoldPrinting);
            }
        }

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // to come out of the printer and folding down the middle.
        doc->Save(ArtifactsDir + u"XpsSaveOptions.BookFold.xps", xpsOptions);
        //ExEnd
    }

    void OptimizeOutput(bool optimizeOutput)
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        auto doc = MakeObject<Document>(MyDir + u"Unoptimized document.docx");

        // Create an "XpsSaveOptions" object to pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        auto saveOptions = MakeObject<XpsSaveOptions>();

        // Set the "OptimizeOutput" property to "true" to take measures such as removing nested or empty canvases
        // and concatenating adjacent runs with identical formatting to optimize the output document's content.
        // This may affect the appearance of the document.
        // Set the "OptimizeOutput" property to "false" to save the document normally.
        saveOptions->set_OptimizeOutput(optimizeOutput);

        doc->Save(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps", saveOptions);
        //ExEnd

        auto outFileInfo = MakeObject<System::IO::FileInfo>(ArtifactsDir + u"XpsSaveOptions.OptimizeOutput.xps");

        if (optimizeOutput)
        {
            ASSERT_GE(50000, outFileInfo->get_Length());
        }
        else
        {
            ASSERT_LT(60000, outFileInfo->get_Length());
        }
    }

    void ExportExactPages()
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExFor:PageSet.#ctor(int[])
        //ExSummary:Shows how to extract pages based on exact page indices.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add five pages to the document.
        for (int i = 1; i < 6; i++)
        {
            builder->Write(String(u"Page ") + i);
            builder->InsertBreak(BreakType::PageBreak);
        }

        // Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        auto xpsOptions = MakeObject<XpsSaveOptions>();

        // Use the "PageSet" property to select a set of the document's pages to save to output XPS.
        // In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
        xpsOptions->set_PageSet(MakeObject<PageSet>(MakeArray<int>({0, 1, 3})));

        doc->Save(ArtifactsDir + u"XpsSaveOptions.ExportExactPages.xps", xpsOptions);
        //ExEnd
    }
};

} // namespace ApiExamples
