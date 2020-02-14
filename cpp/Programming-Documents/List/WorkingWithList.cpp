#include "stdafx.h"
#include "examples.h"

#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Lists/ListCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelAlignment.h>
#include <Aspose.Words.Cpp/Model/Lists/ListLevelCollection.h>
#include <Aspose.Words.Cpp/Model/Lists/ListTemplate.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlSaveOptions.h>
#include <Aspose.Words.Cpp/Model/Saving/OoxmlCompliance.h>
#include <Aspose.Words.Cpp/Model/Text/Font.h>
#include <Aspose.Words.Cpp/Model/Text/ListFormat.h>

using namespace Aspose::Words;
using namespace Aspose::Words::Lists;
using namespace Aspose::Words::Saving;

namespace
{
    void SpecifyListLevel(System::String const &outputDataDir)
    {
        // ExStart:SpecifyListLevel
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Create a numbered list based on one of the Microsoft Word list templates and
        // apply it to the current paragraph in the document builder.
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::NumberArabicDot));

        // There are 9 levels in this list, lets try them all.
        for (int32_t i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(System::String(u"Level ") + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the current paragraph in the document builder.
        builder->get_ListFormat()->set_List(doc->get_Lists()->Add(ListTemplate::BulletDiamonds));

        // There are 9 levels in this list, lets try them all.
        for (int32_t i = 0; i < 9; i++)
        {
            builder->get_ListFormat()->set_ListLevelNumber(i);
            builder->Writeln(System::String(u"Level ") + i);
        }

        // This is a way to stop list formatting. 
        builder->get_ListFormat()->set_List(nullptr);

        System::String outputPath = outputDataDir + u"WorkingWithList.SpecifyListLevel.doc";

        // Save the document to disk.
        builder->get_Document()->Save(outputPath);
        // ExEnd:SpecifyListLevel
        std::cout << "Document is saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void RestartListNumber(System::String const &outputDataDir)
    {
        // ExStart:RestartListNumber
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // Create a list based on a template.
        System::SharedPtr<List> list1 = doc->get_Lists()->Add(ListTemplate::NumberArabicParenthesis);
        // Modify the formatting of the list.
        list1->get_ListLevels()->idx_get(0)->get_Font()->set_Color(System::Drawing::Color::get_Red());
        list1->get_ListLevels()->idx_get(0)->set_Alignment(ListLevelAlignment::Right);

        builder->Writeln(u"List 1 starts below:");
        // Use the first list in the document for a while.
        builder->get_ListFormat()->set_List(list1);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        // Now I want to reuse the first list, but need to restart numbering.
        // This should be done by creating a copy of the original list formatting.
        System::SharedPtr<List> list2 = doc->get_Lists()->AddCopy(list1);

        // We can modify the new list in any way. Including setting new start number.
        list2->get_ListLevels()->idx_get(0)->set_StartAt(10);

        // Use the second list in the document.
        builder->Writeln(u"List 2 starts below:");
        builder->get_ListFormat()->set_List(list2);
        builder->Writeln(u"Item 1");
        builder->Writeln(u"Item 2");
        builder->get_ListFormat()->RemoveNumbers();

        System::String outputPath = outputDataDir + u"WorkingWithList.RestartListNumber.doc";

        // Save the document to disk.
        builder->get_Document()->Save(outputPath);
        // ExEnd:RestartListNumber
        std::cout << "Document is saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;
    }

    void SetRestartAtEachSection(System::String const &outputDataDir)
    {
        // ExStart:SetRestartAtEachSection
        System::SharedPtr<Document> doc = System::MakeObject<Document>();

        doc->get_Lists()->Add(ListTemplate::NumberDefault);

        System::SharedPtr<List> list = doc->get_Lists()->idx_get(0);

        // Set true to specify that the list has to be restarted at each section.
        list->set_IsRestartAtEachSection(true);

        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        builder->get_ListFormat()->set_List(list);

        for (int32_t i = 1; i < 45; i++)
        {
            builder->Writeln(System::String::Format(u"List Item {0}", i));

            // Insert section break.
            if (i == 15)
            {
                builder->InsertBreak(BreakType::SectionBreakNewPage);
            }
        }

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
        System::SharedPtr<OoxmlSaveOptions> options = System::MakeObject<OoxmlSaveOptions>();
        options->set_Compliance(OoxmlCompliance::Iso29500_2008_Transitional);

        System::String outputPath = outputDataDir + u"WorkingWithList.SetRestartAtEachSection.docx";

        // Save the document to disk.
        doc->Save(outputPath, options);
        // ExEnd:SetRestartAtEachSection
        std::cout << "Document is saved successfully." << std::endl << "File saved at " << outputPath.ToUtf8String() << std::endl;

    }
}

void WorkingWithList()
{
    std::cout << "WorkingWithList example started." << std::endl;
    // The path to the documents directory.
    System::String outputDataDir = GetOutputDataDir_WorkingWithList();
    RestartListNumber(outputDataDir);
    SpecifyListLevel(outputDataDir);
    SetRestartAtEachSection(outputDataDir);
    std::cout << "WorkingWithList example finished." << std::endl << std::endl;
}