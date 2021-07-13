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
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/SaveFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Saving/TxtExportHeadersFootersMode.h>
#include <Aspose.Words.Cpp/Saving/TxtListIndentation.h>
#include <Aspose.Words.Cpp/Saving/TxtSaveOptions.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <system/array.h>
#include <system/io/file.h>
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
using namespace Aspose::Words::Saving;

namespace ApiExamples {

class ExTxtSaveOptions : public ApiExampleBase
{
public:
    void PageBreaks(bool forcePageBreaks)
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ForcePageBreaks
        //ExSummary:Shows how to specify whether to preserve page breaks when exporting a document to plaintext.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save"
        // method to modify how we save the document to plaintext.
        auto saveOptions = MakeObject<TxtSaveOptions>();

        // The Aspose.Words "Document" objects have page breaks, just like Microsoft Word documents.
        // Save formats such as ".txt" are one continuous body of text without page breaks.
        // Set the "ForcePageBreaks" property to "true" to preserve all page breaks in the form of '\f' characters.
        // Set the "ForcePageBreaks" property to "false" to discard all page breaks.
        saveOptions->set_ForcePageBreaks(forcePageBreaks);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt", saveOptions);

        // If we load a plaintext document with page breaks,
        // the "Document" object will use them to split the body into pages.
        doc = MakeObject<Document>(ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt");

        ASSERT_EQ(forcePageBreaks ? 3 : 1, doc->get_PageCount());
        //ExEnd

        TestUtil::FileContainsString(forcePageBreaks ? String(u"Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n") : String(u"Page 1\r\nPage 2\r\nPage 3\r\n\r\n"),
                                     ArtifactsDir + u"TxtSaveOptions.PageBreaks.txt");
    }

    void AddBidiMarks(bool addBidiMarks)
    {
        //ExStart
        //ExFor:TxtSaveOptions.AddBidiMarks
        //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello world!");
        builder->get_ParagraphFormat()->set_Bidi(true);
        builder->Writeln(u"שלום עולם!");
        builder->Writeln(u"مرحبا بالعالم!");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->set_Encoding(System::Text::Encoding::get_Unicode());

        // Set the "AddBidiMarks" property to "true" to add marks before runs
        // with right-to-left text to indicate the fact.
        // Set the "AddBidiMarks" property to "false" to write all left-to-right
        // and right-to-left run equally with nothing to indicate which is which.
        saveOptions->set_AddBidiMarks(addBidiMarks);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.AddBidiMarks.txt", saveOptions);

        String docText = System::Text::Encoding::get_Unicode()->GetString(System::IO::File::ReadAllBytes(ArtifactsDir + u"TxtSaveOptions.AddBidiMarks.txt"));

        if (addBidiMarks)
        {
            ASSERT_EQ(u"\ufeffHello world!‎\r\nשלום עולם!‏\r\nمرحبا بالعالم!‏\r\n\r\n", docText);
            ASSERT_TRUE(docText.Contains(u"\u200f"));
        }
        else
        {
            ASSERT_EQ(u"\ufeffHello world!\r\nשלום עולם!\r\nمرحبا بالعالم!\r\n\r\n", docText);
            ASSERT_FALSE(docText.Contains(u"\u200f"));
        }
        //ExEnd
    }

    void ExportHeadersFooters(TxtExportHeadersFootersMode txtExportHeadersFootersMode)
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
        //ExFor:TxtExportHeadersFootersMode
        //ExSummary:Shows how to specify how to export headers and footers to plain text format.
        auto doc = MakeObject<Document>();

        // Insert even and primary headers/footers into the document.
        // The primary header/footers will override the even headers/footers.
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderEven));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderEven)->AppendParagraph(u"Even header");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterEven));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterEven)->AppendParagraph(u"Even footer");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::HeaderPrimary));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::HeaderPrimary)->AppendParagraph(u"Primary header");
        doc->get_FirstSection()->get_HeadersFooters()->Add(MakeObject<HeaderFooter>(doc, HeaderFooterType::FooterPrimary));
        doc->get_FirstSection()->get_HeadersFooters()->idx_get(HeaderFooterType::FooterPrimary)->AppendParagraph(u"Primary footer");

        // Insert pages to display these headers and footers.
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Write(u"Page 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto saveOptions = MakeObject<TxtSaveOptions>();

        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.None"
        // to not export any headers/footers.
        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.PrimaryOnly"
        // to only export primary headers/footers.
        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.AllAtEnd"
        // to place all headers and footers for all section bodies at the end of the document.
        saveOptions->set_ExportHeadersFootersMode(txtExportHeadersFootersMode);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);

        String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.ExportHeadersFooters.txt");

        switch (txtExportHeadersFootersMode)
        {
        case TxtExportHeadersFootersMode::AllAtEnd:
            ASSERT_EQ(String(u"Page 1\r\n") + u"Page 2\r\n" + u"Page 3\r\n" + u"Even header\r\n\r\n" + u"Primary header\r\n\r\n" + u"Even footer\r\n\r\n" +
                          u"Primary footer\r\n\r\n",
                      docText);
            break;

        case TxtExportHeadersFootersMode::PrimaryOnly:
            ASSERT_EQ(String(u"Primary header\r\n") + u"Page 1\r\n" + u"Page 2\r\n" + u"Page 3\r\n" + u"Primary footer\r\n", docText);
            break;

        case TxtExportHeadersFootersMode::None:
            ASSERT_EQ(String(u"Page 1\r\n") + u"Page 2\r\n" + u"Page 3\r\n", docText);
            break;
        }
        //ExEnd
    }

    void TxtListIndentation_()
    {
        //ExStart
        //ExFor:TxtListIndentation
        //ExFor:TxtListIndentation.Count
        //ExFor:TxtListIndentation.Character
        //ExFor:TxtSaveOptions.ListIndentation
        //ExSummary:Shows how to configure list indenting when saving a document to plaintext.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list with three levels of indentation.
        builder->get_ListFormat()->ApplyNumberDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto txtSaveOptions = MakeObject<TxtSaveOptions>();

        // Set the "Character" property to assign a character to use
        // for padding that simulates list indentation in plaintext.
        txtSaveOptions->get_ListIndentation()->set_Character(u' ');

        // Set the "Count" property to specify the number of times
        // to place the padding character for each list indent level.
        txtSaveOptions->get_ListIndentation()->set_Count(3);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);

        String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.TxtListIndentation.txt");

        ASSERT_EQ(String(u"1. Item 1\r\n") + u"   a. Item 2\r\n" + u"      i. Item 3\r\n", docText);
        //ExEnd
    }

    void SimplifyListLabels(bool simplifyListLabels)
    {
        //ExStart
        //ExFor:TxtSaveOptions.SimplifyListLabels
        //ExSummary:Shows how to change the appearance of lists when saving a document to plaintext.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a bulleted list with five levels of indentation.
        builder->get_ListFormat()->ApplyBulletDefault();
        builder->Writeln(u"Item 1");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 3");
        builder->get_ListFormat()->ListIndent();
        builder->Writeln(u"Item 4");
        builder->get_ListFormat()->ListIndent();
        builder->Write(u"Item 5");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto txtSaveOptions = MakeObject<TxtSaveOptions>();

        // Set the "SimplifyListLabels" property to "true" to convert some list
        // symbols into simpler ASCII characters, such as '*', 'o', '+', '>', etc.
        // Set the "SimplifyListLabels" property to "false" to preserve as many original list symbols as possible.
        txtSaveOptions->set_SimplifyListLabels(simplifyListLabels);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);

        String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.SimplifyListLabels.txt");

        if (simplifyListLabels)
        {
            ASSERT_EQ(String(u"* Item 1\r\n") + u"  > Item 2\r\n" + u"    + Item 3\r\n" + u"      - Item 4\r\n" + u"        o Item 5\r\n", docText);
        }
        else
        {
            ASSERT_EQ(String(u"· Item 1\r\n") + u"o Item 2\r\n" + u"§ Item 3\r\n" + u"· Item 4\r\n" + u"o Item 5\r\n", docText);
        }
        //ExEnd
    }

    void ParagraphBreak()
    {
        //ExStart
        //ExFor:TxtSaveOptions
        //ExFor:TxtSaveOptions.SaveFormat
        //ExFor:TxtSaveOptionsBase
        //ExFor:TxtSaveOptionsBase.ParagraphBreak
        //ExSummary:Shows how to save a .txt document with a custom paragraph break.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Paragraph 1.");
        builder->Writeln(u"Paragraph 2.");
        builder->Write(u"Paragraph 3.");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto txtSaveOptions = MakeObject<TxtSaveOptions>();

        ASSERT_EQ(SaveFormat::Text, txtSaveOptions->get_SaveFormat());

        // Set the "ParagraphBreak" to a custom value that we wish to put at the end of every paragraph.
        txtSaveOptions->set_ParagraphBreak(u" End of paragraph.\n\n\t");

        doc->Save(ArtifactsDir + u"TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);

        String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.ParagraphBreak.txt");

        ASSERT_EQ(String(u"Paragraph 1. End of paragraph.\n\n\t") + u"Paragraph 2. End of paragraph.\n\n\t" + u"Paragraph 3. End of paragraph.\n\n\t", docText);
        //ExEnd
    }

    void Encoding_()
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.Encoding
        //ExSummary:Shows how to set encoding for a .txt output document.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Add some text with characters from outside the ASCII character set.
        builder->Write(u"À È Ì Ò Ù.");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto txtSaveOptions = MakeObject<TxtSaveOptions>();

        // Verify that the "Encoding" property contains the appropriate encoding for our document's contents.
        ASPOSE_ASSERT_EQ(System::Text::Encoding::get_UTF8(), txtSaveOptions->get_Encoding());

        doc->Save(ArtifactsDir + u"TxtSaveOptions.Encoding.UTF8.txt", txtSaveOptions);

        String docText = System::Text::Encoding::get_UTF8()->GetString(System::IO::File::ReadAllBytes(ArtifactsDir + u"TxtSaveOptions.Encoding.UTF8.txt"));

        ASSERT_EQ(u"\ufeffÀ È Ì Ò Ù.\r\n", docText);

        // Using an unsuitable encoding may result in a loss of document contents.
        txtSaveOptions->set_Encoding(System::Text::Encoding::get_ASCII());
        doc->Save(ArtifactsDir + u"TxtSaveOptions.Encoding.ASCII.txt", txtSaveOptions);
        docText = System::Text::Encoding::get_ASCII()->GetString(System::IO::File::ReadAllBytes(ArtifactsDir + u"TxtSaveOptions.Encoding.ASCII.txt"));

        ASSERT_EQ(u"? ? ? ? ?.\r\n", docText);
        //ExEnd
    }

    void PreserveTableLayout(bool preserveTableLayout)
    {
        //ExStart
        //ExFor:TxtSaveOptions.PreserveTableLayout
        //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->StartTable();
        builder->InsertCell();
        builder->Write(u"Row 1, cell 1");
        builder->InsertCell();
        builder->Write(u"Row 1, cell 2");
        builder->EndRow();
        builder->InsertCell();
        builder->Write(u"Row 2, cell 1");
        builder->InsertCell();
        builder->Write(u"Row 2, cell 2");
        builder->EndTable();

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        auto txtSaveOptions = MakeObject<TxtSaveOptions>();

        // Set the "PreserveTableLayout" property to "true" to apply whitespace padding to the contents
        // of the output plaintext document to preserve as much of the table's layout as possible.
        // Set the "PreserveTableLayout" property to "false" to save all tables' contents
        // as a continuous body of text, with just a new line for each row.
        txtSaveOptions->set_PreserveTableLayout(preserveTableLayout);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.PreserveTableLayout.txt", txtSaveOptions);

        String docText = System::IO::File::ReadAllText(ArtifactsDir + u"TxtSaveOptions.PreserveTableLayout.txt");

        if (preserveTableLayout)
        {
            ASSERT_EQ(String(u"Row 1, cell 1                                            Row 1, cell 2\r\n") +
                          u"Row 2, cell 1                                            Row 2, cell 2\r\n\r\n",
                      docText);
        }
        else
        {
            ASSERT_EQ(String(u"Row 1, cell 1\r\n") + u"Row 1, cell 2\r\n" + u"Row 2, cell 1\r\n" + u"Row 2, cell 2\r\n\r\n", docText);
        }
        //ExEnd
    }

    void MaxCharactersPerLine()
    {
        //ExStart
        //ExFor:TxtSaveOptions.MaxCharactersPerLine
        //ExSummary:Shows how to set maximum number of characters per line.
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(String(u"Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ") +
                       u"Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

        // Set 30 characters as maximum allowed per one line.
        auto saveOptions = MakeObject<TxtSaveOptions>();
        saveOptions->set_MaxCharactersPerLine(30);

        doc->Save(ArtifactsDir + u"TxtSaveOptions.MaxCharactersPerLine.txt", saveOptions);
        //ExEnd
    }
};

} // namespace ApiExamples
