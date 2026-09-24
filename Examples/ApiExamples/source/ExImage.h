#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

/// <summary>
/// Mostly scenarios that deal with image shapes.
/// </summary>
class ExImage : public ApiExampleBase
{
    typedef ExImage ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void FromFile();
    void FromUrl();
    void FromStream();
    void CreateFloatingPageCenter();
    void CreateFloatingPositionSize();
    void InsertImageWithHyperlink();
    void CreateLinkedImage();
    void DeleteAllImages();
    void DeleteAllImagesPreOrder();
    void ScaleImage();
    void InsertWebpImage();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


