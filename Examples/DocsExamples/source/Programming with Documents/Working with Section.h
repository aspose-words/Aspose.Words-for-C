#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/HeaderFooter.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/PageSetup.h>
#include <Aspose.Words.Cpp/PaperSize.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <Aspose.Words.Cpp/SectionCollection.h>
#include <Aspose.Words.Cpp/TextColumnCollection.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>

#include "DocsExamplesBase.h"
#include "../../../../packages/Aspose.Words.Cpp.22.1.0/build/native/include/Aspose.Words.Cpp/DocumentVisitor.h"

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
        //GistId:11904531c9095a3c413adf28dbe3fe8d
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
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");
        SharedPtr<Section> cloneSection = doc->get_Sections()->idx_get(0)->Clone();
        //ExEnd:CloneSection
    }

    void CopySection()
    {
        //ExStart:CopySection
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto srcDoc = MakeObject<Document>(MyDir + u"Document.docx");
        auto dstDoc = MakeObject<Document>();

        SharedPtr<Section> sourceSection = srcDoc->get_Sections()->idx_get(0);
        auto newSection = System::ExplicitCast<Section>(dstDoc->ImportNode(sourceSection, true));
        dstDoc->get_Sections()->Add(newSection);

        dstDoc->Save(ArtifactsDir + u"WorkingWithSection.CopySection.docx");
        //ExEnd:CopySection
    }

    void DeleteHeaderFooterContent()
    {
        //ExStart:DeleteHeaderFooterContent
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        SharedPtr<Section> section = doc->get_Sections()->idx_get(0);
        section->ClearHeadersFooters();
        //ExEnd:DeleteHeaderFooterContent
    }

    void DeleteHeaderFooterShapes()
    {
        //ExStart:DeleteHeaderFooterShapes
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto doc = MakeObject<Document>(MyDir + u"Document.docx");

        auto section = doc->get_Sections()->idx_get(0);
        section->DeleteHeaderFooterShapes();
        //ExEnd:DeleteHeaderFooterShapes
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
        //GistId:11904531c9095a3c413adf28dbe3fe8d
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
        for (const auto& section : System::IterateOver<Section>(doc))
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

    void SectionChildNodes()
    {
        //ExStart:SectionChildNodes
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->Write(u"Section 1");
        builder->MoveToHeaderFooter(HeaderFooterType::HeaderPrimary);
        builder->Write(u"Primary header");
        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);
        builder->Write(u"Primary footer");

        // A Section is a composite node and can contain child nodes,
        // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
        for (const auto& srcSection : System::IterateOver(doc->get_Sections()->LINQ_OfType<SharedPtr<Section>>()))
        {
            for (const auto& srcNode : System::IterateOver(srcSection->get_Body()))
            {
                switch (srcNode->get_NodeType())
                {
                    case NodeType::Body:
                    {
                        auto body = System::ExplicitCast<Body>(srcNode);
                        std::cout << "Body:" << std::endl;
                        std::cout << std::endl;
                        std::cout << "\t\"" << body->GetText().Trim() << "\"" << std::endl;                        
                        break;
                    }
                    case NodeType::HeaderFooter:
                    {
                        auto headerFooter = System::ExplicitCast<HeaderFooter>(srcNode);
                        auto headerFooterType = headerFooter->get_HeaderFooterType();

                        std::cout << "HeaderFooter type: " << static_cast<std::underlying_type<HeaderFooterType>::type>(headerFooterType) << std::endl;                            
                        std::cout << "\t\"" << headerFooter->GetText().Trim() << "\"" << std::endl;                        
                        break;
                    }
                    default:
                    {
                        throw System::Exception(u"Unexpected node type in a section.");
                    }
                }
            }
        }
        //ExEnd:SectionChildNodes
    }

    void EnsureMinimum()
    {
        //ExStart:EnsureMinimum
        //GistId:11904531c9095a3c413adf28dbe3fe8d
        auto doc = MakeObject<Document>();

        // If we add a new section like this, it will not have a body, or any other child nodes.
        doc->get_Sections()->Add(MakeObject<Section>(doc));
        // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
        doc->get_LastSection()->EnsureMinimum();

        doc->get_Sections()->idx_get(0)->get_Body()->get_FirstParagraph()->AppendChild(MakeObject<Run>(doc, u"Hello world!"));
        //ExEnd:EnsureMinimum
    }
};

}} // namespace DocsExamples::Programming_with_Documents
