#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Lists/List.h>
#include <Aspose.Words.Cpp/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Lists/ListFormat.h>
#include <Aspose.Words.Cpp/Lists/ListLevel.h>
#include <Aspose.Words.Cpp/Lists/ListLevelAlignment.h>
#include <Aspose.Words.Cpp/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <drawing/color.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithList : public DocsExamplesBase
{
public:
    void RestartListAtEachSection()
    {
        //ExStart:RestartListAtEachSection
        auto doc = MakeObject<Document>();

        doc->get_Lists()->Add(ListTemplate::NumberDefault);

        SharedPtr<List> list = doc->get_Lists()->idx_get(0);
        list->set_IsRestartAtEachSection(true);

        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->get_ListFormat()->set_List(list);

        for (int i = 1; i < 45; i++)
        {
            builder->Writeln(String::Format(u"List Item {0}", i));

            if (i == 15)
            {
                builder->InsertBreak(BreakType::SectionBreakNewPage);
            }
        }

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376.
        auto options = MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        doc->Save(ArtifactsDir + u"WorkingWithList.RestartListAtEachSection.docx", options);
        //ExEnd:RestartListAtEachSection
    }

    void SpecifyListLevel()
    {
        //ExStart:SpecifyListLevel
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a numbered list based on one of the Microsoft Word list templates
        // and apply it to the document builder's current paragraph.
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::NumberArabicDot));

        // There are nine levels in this list, let's try them all.
        for (int i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(String(u"Level ") + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the document builder's current paragraph.
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::BulletDiamonds));

        for (int i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(String(u"Level ") + i);
        }

        // This is a way to stop list formatting.
        builder->get_ListFormat()->set_List(nullptr);

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithList.SpecifyListLevel.docx");
        //ExEnd:SpecifyListLevel
    }

    void RestartListNumber()
    {
        //ExStart:RestartListNumber
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Create a list based on a template.
        SharedPtr<List> list1 = doc->get_Lists()->Add(ListTemplate::NumberArabicParenthesis);
        list1->get_ListLevels()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Red());
        list1->get_ListLevels()->idx_get(0)->set_Alignment(ListLevelAlignment::Right);

        builder->Writeln(u"List 1 starts below:");
        builder->get_ListFormat()->set_List(list1);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        // To reuse the first list, we need to restart numbering by creating a copy of the original list formatting.
        SharedPtr<List> list2 = doc->get_Lists()->AddCopy(list1);

        // We can modify the new list in any way, including setting a new start number.
        list2->get_ListLevels()->idx_get(0)->set_StartAt(10);

        builder->Writeln(u"List 2 starts below:");
        builder->get_ListFormat()->set_List(list2);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithList.RestartListNumber.docx");
        //ExEnd:RestartListNumber
    }
};

}} // namespace DocsExamples::Programming_with_Documents
