#pragma once

#include <cstdint>
#include <system/array.h>
#include <system/collections/ienumerable.h>

#include "DocsExamplesBase.h"
#include "LINQ Reporting Engine/Helpers/Data Source Objects/Client.h"
#include "LINQ Reporting Engine/Helpers/Data Source Objects/Contract.h"
#include "LINQ Reporting Engine/Helpers/Data Source Objects/Manager.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {
class Manager;
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects

using namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers {

class Common : public DocsExamplesBase
{
public:
    /// <summary>
    /// Return the first manager from Managers, which is an enumeration of instances of the Manager class.
    /// </summary>
    static System::SharedPtr<Data_Source_Objects::Manager> GetManager();
    /// <summary>
    /// Return an enumeration of instances of the Client class.
    /// </summary>
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Data_Source_Objects::Client>>> GetClients();
    /// <summary>
    ///  Return an enumeration of instances of the Manager class.
    /// </summary>
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Data_Source_Objects::Manager>>> GetManagers();
    /// <summary>
    ///  Return an enumeration of instances of the Contract class.
    /// </summary>
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Data_Source_Objects::Contract>>> GetContracts();

private:
    /// <summary>
    /// Return an array of photo bytes.
    /// </summary>
    static System::ArrayPtr<uint8_t> Photo();
};

}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers
