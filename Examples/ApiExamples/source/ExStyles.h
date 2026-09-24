#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExStyles : public ApiExampleBase
{
    typedef ExStyles ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Styles();
    void CreateStyle();
    void StyleCollection();
    void RemoveStylesFromStyleGallery();
    void ChangeTocsTabStops();
    void CopyStyleSameDocument();
    void CopyStyleDifferentDocument();
    void DefaultStyles();
    void ParagraphStyleBulletedList();
    void StyleAliases();
    void LockStyle();
    void StylePriority();
    void LinkedStyleName();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


