#pragma once

#include <system/string.h>
#include <system/globalization/culture_info.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Fields/IFieldUpdateCultureProvider.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Fields;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExFieldOptions : public ApiExampleBase
{
    typedef ExFieldOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
private:

    /// <summary>
    /// Provides a CultureInfo object that should be used during the update of a field.
    /// </summary>
    class FieldUpdateCultureProvider : public IFieldUpdateCultureProvider
    {
        typedef FieldUpdateCultureProvider ThisType;
        typedef IFieldUpdateCultureProvider BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Returns a CultureInfo object to be used during the field's update.
        /// </summary>
        System::SharedPtr<System::Globalization::CultureInfo> GetCulture(System::String name, System::SharedPtr<Aspose::Words::Fields::Field> field) override;
        
    };
    
    
public:

    void CurrentUser();
    void FileName();
    void Bidi();
    void LegacyNumberFormat();
    void PreProcessCulture();
    void TableOfAuthorityCategories();
    void UseInvariantCultureNumberFormat();
    //ExStart
    //ExFor:FieldOptions.FieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider
    //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
    //ExSummary:Shows how to specify a culture which parses date/time formatting for each field.
    void DefineDateTimeFormatting();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


