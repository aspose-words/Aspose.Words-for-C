#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Settings/ViewOptions.h>
#include <Aspose.Words.Cpp/Settings/ViewType.h>
#include <Aspose.Words.Cpp/Settings/ZoomType.h>
#include <system/array.h>
#include <system/io/file.h>
#include <system/io/memory_stream.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <system/text/encoding.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Settings;

namespace ApiExamples {

class ExViewOptions : public ApiExampleBase
{
public:
    void SetZoomPercentage()
    {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ViewType
        //ExFor:ViewOptions.ZoomPercent
        //ExFor:ViewOptions.ZoomType
        //ExFor:ViewType
        //ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        doc->get_ViewOptions()->set_ViewType(ViewType::PageLayout);
        doc->get_ViewOptions()->set_ZoomPercent(50);

        ASSERT_EQ(ZoomType::Custom, doc->get_ViewOptions()->get_ZoomType());
        ASSERT_EQ(ZoomType::None, doc->get_ViewOptions()->get_ZoomType());

        doc->Save(ArtifactsDir + u"ViewOptions.SetZoomPercentage.doc");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.SetZoomPercentage.doc");

        ASSERT_EQ(ViewType::PageLayout, doc->get_ViewOptions()->get_ViewType());
        ASPOSE_ASSERT_EQ(50.0, doc->get_ViewOptions()->get_ZoomPercent());
        ASSERT_EQ(ZoomType::None, doc->get_ViewOptions()->get_ZoomType());
    }

    void SetZoomType(ZoomType zoomType)
    {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ZoomType
        //ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Set the "ZoomType" property to "ZoomType.PageWidth" to get Microsoft Word
        // to automatically zoom the document to fit the width of the page.
        // Set the "ZoomType" property to "ZoomType.FullPage" to get Microsoft Word
        // to automatically zoom the document to make the entire first page visible.
        // Set the "ZoomType" property to "ZoomType.TextFit" to get Microsoft Word
        // to automatically zoom the document to fit the inner text margins of the first page.
        doc->get_ViewOptions()->set_ZoomType(zoomType);

        doc->Save(ArtifactsDir + u"ViewOptions.SetZoomType.doc");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.SetZoomType.doc");

        ASSERT_EQ(zoomType, doc->get_ViewOptions()->get_ZoomType());
    }

    void DisplayBackgroundShape(bool displayBackgroundShape)
    {
        //ExStart
        //ExFor:ViewOptions.DisplayBackgroundShape
        //ExSummary:Shows how to hide/display document background images in view options.
        // Use an HTML string to create a new document with a flat background color.
        const String html = u"<html>\r\n                <body style='background-color: blue'>\r\n                    <p>Hello world!</p>\r\n                "
                            u"</body>\r\n            </html>";

        auto doc = MakeObject<Document>(MakeObject<System::IO::MemoryStream>(System::Text::Encoding::get_Unicode()->GetBytes(html)));

        // The source for the document has a flat color background,
        // the presence of which will set the "DisplayBackgroundShape" flag to "true".
        ASSERT_TRUE(doc->get_ViewOptions()->get_DisplayBackgroundShape());

        // Keep the "DisplayBackgroundShape" as "true" to get the document to display the background color.
        // This may affect some text colors to improve visibility.
        // Set the "DisplayBackgroundShape" to "false" to not display the background color.
        doc->get_ViewOptions()->set_DisplayBackgroundShape(displayBackgroundShape);

        doc->Save(ArtifactsDir + u"ViewOptions.DisplayBackgroundShape.docx");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.DisplayBackgroundShape.docx");

        ASPOSE_ASSERT_EQ(displayBackgroundShape, doc->get_ViewOptions()->get_DisplayBackgroundShape());
    }

    void DisplayPageBoundaries(bool doNotDisplayPageBoundaries)
    {
        //ExStart
        //ExFor:ViewOptions.DoNotDisplayPageBoundaries
        //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert content that spans across 3 pages.
        builder->Writeln(u"Paragraph 1, Page 1.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Paragraph 2, Page 2.");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Paragraph 3, Page 3.");

        // Insert a header and a footer.
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Writeln(u"This is the header.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Writeln(u"This is the footer.");

        // This document contains a small amount of content that takes up a few full pages worth of space.
        // Set the "DoNotDisplayPageBoundaries" flag to "true" to get older versions of Microsoft Word to omit headers,
        // footers, and much of the vertical whitespace when displaying our document.
        // Set the "DoNotDisplayPageBoundaries" flag to "false" to get older versions of Microsoft Word
        // to normally display our document.
        doc->get_ViewOptions()->set_DoNotDisplayPageBoundaries(doNotDisplayPageBoundaries);

        doc->Save(ArtifactsDir + u"ViewOptions.DisplayPageBoundaries.doc");
        //ExEnd

        doc = MakeObject<Document>(ArtifactsDir + u"ViewOptions.DisplayPageBoundaries.doc");

        ASPOSE_ASSERT_EQ(doNotDisplayPageBoundaries, doc->get_ViewOptions()->get_DoNotDisplayPageBoundaries());
    }

    void FormsDesign(bool useFormsDesign)
    {
        //ExStart
        //ExFor:ViewOptions.FormsDesign
        //ExSummary:Shows how to enable/disable forms design mode.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Hello world!");

        // Set the "FormsDesign" property to "false" to keep forms design mode disabled.
        // Set the "FormsDesign" property to "true" to enable forms design mode.
        doc->get_ViewOptions()->set_FormsDesign(useFormsDesign);

        doc->Save(ArtifactsDir + u"ViewOptions.FormsDesign.xml");

        ASPOSE_ASSERT_EQ(useFormsDesign, System::IO::File::ReadAllText(ArtifactsDir + u"ViewOptions.FormsDesign.xml").Contains(u"<w:formsDesign />"));
        //ExEnd
    }
};

} // namespace ApiExamples
