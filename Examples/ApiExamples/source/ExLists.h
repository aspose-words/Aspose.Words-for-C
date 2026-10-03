#pragma once

#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Lists/List.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Lists;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExLists : public ApiExampleBase
{
    typedef ExLists ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void ApplyDefaultBulletsAndNumbers();
    void SpecifyListLevel();
    void NestedLists();
    void CreateCustomList();
    void RestartNumberingUsingListCopy();
    void CreateAndUseListStyle();
    void DetectBulletedParagraphs();
    void RemoveBulletsFromParagraphs();
    void ApplyExistingListToParagraphs();
    void ApplyNewListToParagraphs();
    //ExStart
    //ExFor:ListTemplate
    //ExSummary:Shows how to create a document that contains all outline headings list templates.
    void OutlineHeadingTemplates();
    //ExStart
    //ExFor:ListCollection
    //ExFor:ListCollection.AddCopy(List)
    //ExSummary:Shows how to create a document with a sample of all the lists from another document.
    void PrintOutAllLists();
    void ListDocument();
    void CreateListRestartAfterHigher();
    void GetListLabels();
    void CreatePictureBullet();
    void GetCustomNumberStyleFormat();
    void HasSameTemplate();
    void SetCustomNumberStyleFormat();
    void AddSingleLevelList();
    void RemoveTabStopFromListLevel();
    
protected:

    static void AddOutlineHeadingParagraphs(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list, System::String title);
    //ExEnd
    void TestOutlineHeadingTemplates(System::SharedPtr<Aspose::Words::Document> doc);
    static void AddListSample(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list);
    //ExEnd
    void TestPrintOutAllLists(System::SharedPtr<Aspose::Words::Document> listSourceDoc, System::SharedPtr<Aspose::Words::Document> outDoc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


