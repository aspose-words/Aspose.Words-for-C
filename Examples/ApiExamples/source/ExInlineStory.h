#pragma once

#include <gtest/gtest.h>
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


