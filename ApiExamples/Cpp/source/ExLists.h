#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { class DocumentBuilder; } }
namespace Aspose { namespace Words { namespace Lists { class List; } } }
namespace Aspose { namespace Words { class Document; } }

namespace ApiExamples {

class ExLists : public ApiExampleBase
{
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
    
protected:

    static void AddOutlineHeadingParagraphs(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list, System::String title);
    void TestOutlineHeadingTemplates(System::SharedPtr<Aspose::Words::Document> doc);
    static void AddListSample(System::SharedPtr<Aspose::Words::DocumentBuilder> builder, System::SharedPtr<Aspose::Words::Lists::List> list);
    void TestPrintOutAllLists(System::SharedPtr<Aspose::Words::Document> listSourceDoc, System::SharedPtr<Aspose::Words::Document> outDoc);
    
};

} // namespace ApiExamples


