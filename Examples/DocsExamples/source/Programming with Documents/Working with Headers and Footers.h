#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/HeaderFooterCollection.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Orientation.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/ParagraphAlignment.h>
#include <Aspose.Words.Cpp/ParagraphFormat.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/Tables/Cell.h>
#include <Aspose.Words.Cpp/Tables/CellFormat.h>
#include <Aspose.Words.Cpp/Tables/PreferredWidth.h>
#include <Aspose.Words.Cpp/Tables/Row.h>
#include <Aspose.Words.Cpp/Tables/Table.h>
#include <Aspose.Words.Cpp/Tables/TableCollection.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace Aspose::Words::Tables;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithHeadersAndFooters : public DocsExamplesBase
{
public:
    void CreateHeaderFooter()
    {
        //ExStart:CreateHeaderFooter
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Use HeaderPrimary and FooterPrimary
        // if you want to set header/footer for all document.
        // This header/footer type also responsible for odd pages.
        //ExStart:HeaderFooterType
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Header for page.");
        //ExEnd:HeaderFooterType

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"Footer for page.");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
        //ExEnd:CreateHeaderFooter
    }

    void DifferentFirstPage()
    {
        //ExStart:DifferentFirstPage
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify that we want different headers and footers for first page.
        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->Write(u"Header for the first page.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterFirst);
        builder->Write(u"Footer for the first page.");

        builder->MoveToSection(0);
        builder->Writeln(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.DifferentFirstPage.docx");
        //ExEnd:DifferentFirstPage
    }
    
    void OddEvenPages()
    {
        //ExStart:OddEvenPages
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Specify that we want different headers and footers for even and odd pages.            
        builder->get_PageSetup()->set_OddAndEvenPagesHeaderFooter(true);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderEven);
        builder->Write(u"Header for even pages.");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Header for odd pages.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterEven);
        builder->Write(u"Footer for even pages.");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"Footer for odd pages.");
               
        builder->MoveToSection(0);
        builder->Writeln(u"Page 1");
        builder->InsertBreak(BreakType::PageBreak);
        builder->Writeln(u"Page 2");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.OddEvenPages.docx");
        //ExEnd:OddEvenPages
    }

    void InsertImage()
    {
        //ExStart:InsertImage
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->InsertImage(ImagesDir + u"Logo.jpg", RelativeHorizontalPosition::RightMargin, 10,
            RelativeVerticalPosition::Page, 10, 50, 50, WrapType::Through);

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.InsertImage.docx");
        //ExEnd:InsertImage
    }

    void FontProps()
    {
        //ExStart:FontProps
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_Font()->set_Size(14);
        builder->Write(u"Header for page.");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.FontProps.docx");
        //ExEnd:FontProps
    }

    void PageNumbers()
    {
        //ExStart:PageNumbers
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Right);
        builder->Write(u"Page ");
        builder->InsertField(u"PAGE", u"");
        builder->Write(u" of ");
        builder->InsertField(u"NUMPAGES", u"");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.PageNumbers.docx");
        //ExEnd:PageNumbers
    }

    void LinkToPreviousHeaderFooter()
    {
        //ExStart:LinkToPreviousHeaderFooter
        //GistId:e26133cadfcbe798f1ae301e2941f5ff
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->get_PageSetup()->set_DifferentFirstPageHeaderFooter(true);

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderFirst);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Bold(true);
        builder->get_Font()->set_Size(14);
        builder->Write(u"Header for the first page.");

        builder->MoveToDocumentEnd();
        builder->InsertBreak(BreakType::SectionBreakNewPage);

        auto currentSection = builder->get_CurrentSection();
        auto pageSetup = currentSection->get_PageSetup();
        pageSetup->set_Orientation(Orientation::Landscape);
        // This section does not need a different first-page header/footer we need only one title page in the document,
        // and the header/footer for this page has already been defined in the previous section.
        pageSetup->set_DifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section
        // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
        // is different for the new section.
        currentSection->get_HeadersFooters()->LinkToPrevious(false);
        currentSection->get_HeadersFooters()->Clear();

        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->get_ParagraphFormat()->set_Alignment(ParagraphAlignment::Center);
        builder->get_Font()->set_Name(u"Arial");
        builder->get_Font()->set_Size(12);
        builder->Write(u"New Header for the first page.");

        doc->Save(ArtifactsDir + u"WorkingWithHeadersAndFooters.LinkToPreviousHeaderFooter.docx");
        //ExEnd:LinkToPreviousHeaderFooter
    }    

    //ExStart:CopyHeadersFootersFromPreviousSection
    //GistId:e26133cadfcbe798f1ae301e2941f5ff
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    void CopyHeadersFootersFromPreviousSection(SharedPtr<Section> section)
    {
        auto previousSection = System::ExplicitCast<Section>(section->get_PreviousSibling());

        if (previousSection == nullptr)
        {
            return;
        }

        section->get_HeadersFooters()->Clear();

        for (const auto& headerFooter : System::IterateOver<HeaderFooter>(previousSection->get_HeadersFooters()))
        {
            section->get_HeadersFooters()->Add((System::ExplicitCast<Node>(headerFooter))->Clone(true));
        }
    }
    //ExEnd:CopyHeadersFootersFromPreviousSection
};

}} // namespace DocsExamples::Programming_with_Documents
