#pragma once

#include <system/guid.h>
#include <system/collections/list.h>
#include <system/array.h>
#include <gtest/gtest.h>
#include <cstdint>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExSignDocumentCustom : public ApiExampleBase
{
    typedef ExSignDocumentCustom ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    class Signee : public System::Object
    {
        typedef Signee ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        System::Guid get_PersonId() const;
        void set_PersonId(System::Guid value);
        System::String get_Name() const;
        void set_Name(System::String value);
        System::String get_Position() const;
        void set_Position(System::String value);
        System::ArrayPtr<uint8_t> get_Image() const;
        void set_Image(System::ArrayPtr<uint8_t> value);
        
        Signee(System::Guid guid, System::String name, System::String position, System::ArrayPtr<uint8_t> image);
        
    private:
    
        System::Guid pr_PersonId;
        System::String pr_Name;
        System::String pr_Position;
        System::ArrayPtr<uint8_t> pr_Image;
        
    };
    
    
public:

    //ExStart
    //ExFor:CertificateHolder
    //ExFor:SignatureLineOptions.Signer
    //ExFor:SignatureLineOptions.SignerTitle
    //ExFor:SignatureLine.Id
    //ExFor:SignOptions.SignatureLineId
    //ExFor:SignOptions.SignatureLineImage
    //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
    //ExSummary:Shows how to add a signature line to a document, and then sign it using a digital certificate.
    static void Sign();
    
protected:

    static System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee>>>& mSignees();
    
    /// <summary>
    /// Creates a copy of a source document signed using provided signee information and X509 certificate.
    /// </summary>
    static void SignDocument(System::String srcDocumentPath, System::String dstDocumentPath, System::SharedPtr<Aspose::Words::ApiExamples::ExSignDocumentCustom::Signee> signeeInfo, System::String certificatePath, System::String certificatePassword);
    static void CreateSignees();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


