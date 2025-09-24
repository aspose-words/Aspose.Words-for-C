#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Footnotes/FootnotePosition.h>
#include <Aspose.Words.Cpp/Model/Footnotes/EndnotePosition.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Notes;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExInlineStory : public ApiExampleBase
{
    typedef ExInlineStory ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void PositionFootnote(Aspose::Words::Notes::FootnotePosition footnotePosition);
    void PositionEndnote(Aspose::Words::Notes::EndnotePosition endnotePosition);
    void RefMarkNumberStyle();
    void NumberingRule();
    void StartNumber();
    void AddFootnote();
    void FootnoteEndnote();
    void AddComment();
    void InlineStoryRevisions();
    void InsertInlineStoryNodes();
    void DeleteShapes();
    void UpdateActualReferenceMarks();
    void EndnoteSeparator();
    void FootnoteSeparator();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


