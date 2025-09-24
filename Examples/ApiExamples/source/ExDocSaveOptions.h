#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocSaveOptions : public ApiExampleBase
{
    typedef ExDocSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void SaveAsDoc();
    void TempFolder();
    void PictureBullets();
    void UpdateLastPrintedProperty(bool isUpdateLastPrintedProperty);
    void UpdateCreatedTimeProperty(bool isUpdateCreatedTimeProperty);
    void AlwaysCompressMetafiles(bool compressAllMetafiles);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


