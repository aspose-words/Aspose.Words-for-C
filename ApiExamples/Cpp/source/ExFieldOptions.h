#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/globalization/culture_info.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdateCultureProvider.h>

#include "ApiExampleBase.h"

namespace Aspose { namespace Words { namespace Fields { class Field; } } }

namespace ApiExamples {

class ExFieldOptions : public ApiExampleBase
{
private:

    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a particular field.
    /// </summary>
    class FieldUpdateCultureProvider : public Aspose::Words::Fields::IFieldUpdateCultureProvider
    {
        typedef FieldUpdateCultureProvider ThisType;
        typedef Aspose::Words::Fields::IFieldUpdateCultureProvider BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Returns a CultureInfo object to be used during the field's update.
        /// </summary>
        System::SharedPtr<System::Globalization::CultureInfo> GetCulture(System::String name, System::SharedPtr<Aspose::Words::Fields::Field> field) override;
        
    };
    
    
public:

    void FieldOptionsCurrentUser();
    void FieldOptionsFileName();
    void FieldOptionsBidi();
    void FieldOptionsLegacyNumberFormat();
    void FieldOptionsPreProcessCulture();
    void FieldOptionsToaCategories();
    void FieldOptionsUseInvariantCultureNumberFormat();
    void DefineDateTimeFormatting();
    
};

} // namespace ApiExamples


