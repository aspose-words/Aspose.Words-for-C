#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
    void OutlineHeadingTemplates();
    void PrintOutAllLists();
    void ListDocument();
    void CreateListRestartAfterHigher();
    void GetListLabels();
    void CreatePictureBullet();
    void GetCustomNumberStyleFormat();
    void HasSameTemplate();
    void SetCustomNumberStyleFormat();
    void AddSingleLevelList();
    
protected:

    static void AddOutlineHeadingParagraphs(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list, System::String title);
    void TestOutlineHeadingTemplates(System::SharedPtr<Aspose::Words::Document> doc);
    static void AddListSample(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list);
    void TestPrintOutAllLists(System::SharedPtr<Aspose::Words::Document> listSourceDoc, System::SharedPtr<Aspose::Words::Document> outDoc);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


