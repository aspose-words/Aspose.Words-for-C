#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include "ApiExampleBase.h"

namespace ApiExamples {

class ExDigitalSignatureUtil : public ApiExampleBase
{
public:

    void LoadAndRemove();
    void SignDocument();
    void SignDocumentObfuscationBug();
    void IncorrectDecryptionPassword();
    void DecryptionPassword();
    void NoArgumentsForSing();
    void NoCertificateForSign();
    
};

} // namespace ApiExamples


