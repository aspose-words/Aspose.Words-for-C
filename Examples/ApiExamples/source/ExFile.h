#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExFile : public ApiExampleBase
{
    typedef ExFile ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void CatchFileCorruptedException();
    void DetectEncoding();
    void FileFormatToString();
    void DetectDocumentEncryption();
    void DetectDigitalSignatures();
    void SaveToDetectedFileFormat();
    void DetectFileFormat_SaveFormatToLoadFormat();
    void ExtractImages();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


