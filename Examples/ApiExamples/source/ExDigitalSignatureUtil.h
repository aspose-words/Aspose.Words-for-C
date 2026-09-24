#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDigitalSignatureUtil : public ApiExampleBase
{
    typedef ExDigitalSignatureUtil ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Load();
    void Remove();
    void RemoveSignatures();
    void SignDocument();
    void DecryptionPassword();
    void SignDocumentObfuscationBug();
    void IncorrectDecryptionPassword();
    void NoArgumentsForSing();
    void NoCertificateForSign();
    void XmlDsig();
    void SignDocumentWithOptions();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


