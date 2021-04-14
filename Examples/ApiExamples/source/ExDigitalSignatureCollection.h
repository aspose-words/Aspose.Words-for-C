#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <iostream>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/CertificateHolder.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignature.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureCollection.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureType.h>
#include <Aspose.Words.Cpp/Model/DigitalSignatures/DigitalSignatureUtil.h>
#include <security/cryptography/x509_certificates/x500_distinguished_name.h>
#include <security/cryptography/x509_certificates/x509_certificate_2.h>
#include <system/collections/ienumerator.h>
#include <system/date_time.h>
#include <system/details/dispose_guard.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>

#include "ApiExampleBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::DigitalSignatures;

namespace ApiExamples {

class ExDigitalSignatureCollection : public ApiExampleBase
{
public:
    void GetEnumerator()
    {
        //ExStart
        //ExFor:DigitalSignatureCollection.GetEnumerator
        //ExSummary:Shows how to print all the digital signatures of a signed document.
        SharedPtr<DigitalSignatureCollection> digitalSignatures = DigitalSignatureUtil::LoadSignatures(MyDir + u"Digitally signed.docx");

        {
            SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<DigitalSignature>>> enumerator = digitalSignatures->GetEnumerator();
            while (enumerator->MoveNext())
            {
                SharedPtr<DigitalSignature> ds = enumerator->get_Current();

                if (ds != nullptr)
                {
                    std::cout << System::ObjectExt::ToString(ds) << std::endl;
                }
            }
        }
        //ExEnd

        ASSERT_EQ(1, digitalSignatures->get_Count());

        SharedPtr<DigitalSignature> signature = digitalSignatures->idx_get(0);

        ASSERT_TRUE(signature->get_IsValid());
        ASSERT_EQ(DigitalSignatureType::XmlDsig, signature->get_SignatureType());
        ASSERT_EQ(u"12/23/2010 02:14:40 AM", signature->get_SignTime().ToString(u"MM/dd/yyyy hh:mm:ss tt"));
        ASSERT_EQ(u"Test Sign", signature->get_Comments());

        ASSERT_EQ(signature->get_IssuerName(), signature->get_CertificateHolder()->get_Certificate()->get_IssuerName()->get_Name());
        ASSERT_EQ(signature->get_SubjectName(), signature->get_CertificateHolder()->get_Certificate()->get_SubjectName()->get_Name());

        ASSERT_EQ(String(u"CN=VeriSign Class 3 Code Signing 2009-2 CA, ") + u"OU=Terms of use at https://www.verisign.com/rpa (c)09, " +
                      u"OU=VeriSign Trust Network, " + u"O=\"VeriSign, Inc.\", " + u"C=US",
                  signature->get_IssuerName());

        ASSERT_EQ(String(u"CN=Aspose Pty Ltd, ") + u"OU=Digital ID Class 3 - Microsoft Software Validation v2, " + u"O=Aspose Pty Ltd, " + u"L=Lane Cove, " +
                      u"S=New South Wales, " + u"C=AU",
                  signature->get_SubjectName());
    }
};

} // namespace ApiExamples
