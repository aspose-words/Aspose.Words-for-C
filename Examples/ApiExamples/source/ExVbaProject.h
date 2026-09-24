#pragma once

#include <system/string.h>
#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReference.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Vba;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExVbaProject : public ApiExampleBase
{
    typedef ExVbaProject ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void CreateNewVbaProject();
    void CloneVbaProject();
    //ExStart
    //ExFor:VbaReference
    //ExFor:VbaReference.Type
    //ExFor:VbaReference.LibId
    //ExFor:VbaReferenceCollection
    //ExFor:VbaReferenceCollection.Item(Int32)
    //ExFor:VbaReferenceCollection.Count
    //ExFor:VbaReferenceCollection.RemoveAt(int)
    //ExFor:VbaReferenceCollection.Remove(VbaReference)
    //ExFor:VbaReferenceType
    //ExFor:VbaProject.References
    //ExSummary:Shows how to get/remove an element from the VBA reference collection.
    void RemoveVbaReference();
    //ExEnd
    void IsProtected();
    
protected:

    /// <summary>
    /// Returns string representing LibId path of a specified reference.
    /// </summary>
    static System::String GetLibIdPath(System::SharedPtr<Aspose::Words::Vba::VbaReference> reference);
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static System::String GetLibIdReferencePath(System::String libIdReference);
    /// <summary>
    /// Returns path from a specified identifier of an Automation type library.
    /// </summary>
    static System::String GetLibIdProjectPath(System::String libIdProject);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


