#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/PaperSize.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/TextColumnCollection.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithSection : public DocsExamplesBase
{
public:
    void AddSection()
    {
        //ExStart:AddSection
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello1");
        builder->Writeln(u"Hello2");

        auto sectionToAdd = MakeObject<Section>(doc);
        doc->get_Sections()->Add(sectionToAdd);
        //ExEnd:AddSection
    }

    void DeleteSection()
    {
        //ExStart:DeleteSection
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello1");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello2");
        doc->AppendChild(MakeObject<Section>(doc));

        doc->get_Sections()->RemoveAt(0);
        //ExEnd:DeleteSection
    }

    void DeleteAllSections()
    {
        //ExStart:DeleteAllSections
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello1");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello2");
        doc->AppendChild(MakeObject<Section>(doc));

        doc->get_Sections()->Clear();
        //ExEnd:DeleteAllSections
    }

    void AppendSectionContent()
    {
        //ExStart:AppendSectionContent
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello1");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello22");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello3");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello45");

        // This is the section that we will append and prepend to.
        SharedPtr<Section> section = doc->get_Sections()->idx_get(2);

        // This copies the content of the 1st section and inserts it at the beginning of the specified section.
        SharedPtr<Section> sectionToPrepend = doc->get_Sections()->idx_get(0);
        section->PrependContent(sectionToPrepend);

        // This copies the content of the 2nd section and inserts it at the end of the specified section.
        SharedPtr<Section> sectionToAppend = doc->get_Sections()->idx_get(1);
        section->AppendContent(sectionToAppend);
        //ExEnd:AppendSectionContent
    }

    void CloneSection()
    {
        //ExStart:CloneSection
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
        //ExEnd:CloneSection
    }

    void CopySection()
    {
        //ExStart:CopySection
        auto srcDoc = MakeObject<Document>(MyDir + u"Document.docx");
        auto dstDoc = MakeObject<Document>();

        SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
        auto newSection = System::DynamicCast<Section>(dstDoc->ImportNode(sourceSection, true));
        dstDoc->get_Sections()->Add(newSection);

        dstDoc->Save(ArtifactsDir + u"WorkingWithSection.CopySection.docx");
        //ExEnd:CopySection
    }

    void DeleteHeaderFooterContent()
    {
        //ExStart:DeleteHeaderFooterContent
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
        section->ClearHeadersFooters();
        //ExEnd:DeleteHeaderFooterContent
    }

    void DeleteSectionContent()
    {
        //ExStart:DeleteSectionContent
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
        section->ClearContent();
        //ExEnd:DeleteSectionContent
    }

    void ModifyPageSetupInAllSections()
    {
        //ExStart:ModifyPageSetupInAllSections
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Writeln(u"Hello1");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello22");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello3");
        doc->AppendChild(MakeObject<Section>(doc));
        builder->Writeln(u"Hello45");

        // It is important to understand that a document can contain many sections,
        // and each section has its page setup. In this case, we want to modify them all.
        for (auto section : System::IterateOver<Section>(doc))
        {
            section->get_PageSetup()->set_PaperSize(PaperSize::Letter);
        }

        doc->Save(ArtifactsDir + u"WorkingWithSection.ModifyPageSetupInAllSections.doc");
        //ExEnd:ModifyPageSetupInAllSections
    }

    void SectionsAccessByIndex()
    {
        //ExStart:SectionsAccessByIndex
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
        section->get_PageSetup()->set_LeftMargin(90);
        // 3.17 cm
        section->get_PageSetup()->set_RightMargin(90);
        // 3.17 cm
        section->get_PageSetup()->set_TopMargin(72);
        // 2.54 cm
        section->get_PageSetup()->set_BottomMargin(72);
        // 2.54 cm
        section->get_PageSetup()->set_HeaderDistance(35.4);
        // 1.25 cm
        section->get_PageSetup()->set_FooterDistance(35.4);
        // 1.25 cm
        section->get_PageSetup()->get_TextColumns()->set_Spacing(35.4);
        // 1.25 cm
        //ExEnd:SectionsAccessByIndex
    }
};

}} // namespace DocsExamples::Programming_with_Documents
